"""Daily PO purchase-line reconciliation, attached to the application's data.

No posting to billing and no outbound services. Review state is stored separately
from accounting data, under the same data directory as the application database.
"""
import cgi
import hashlib
import json
import sqlite3
import threading
from collections import Counter
from contextlib import closing
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import material_reconciliation as rec
import material_review_ui

LOCK = threading.RLock()
START_DATE = '2026-10-27'
PARSER_VERSION = 'detailed-preview-v1'
ISSUE_STATUSES = ('Open', 'In progress', 'Waiting on information', 'Resolved')


class ReviewError(ValueError):
    def __init__(self, message, status=400):
        super().__init__(message)
        self.status = status


def store(app):
    app.DATA_DIR.mkdir(parents=True, exist_ok=True)
    # Keep history in the main database so normal dashboard backups include it.
    con = sqlite3.connect(app.DB_PATH, timeout=30)
    con.row_factory = sqlite3.Row
    return con


def init_schema(con):
    rec.init_schema(con)
    con.executescript('''
      CREATE TABLE IF NOT EXISTS material_review_settings (key TEXT PRIMARY KEY, value TEXT NOT NULL);
      CREATE TABLE IF NOT EXISTS material_line_issues (
        purchase_key TEXT PRIMARY KEY, source_json TEXT NOT NULL,
        issue_group TEXT NOT NULL, description TEXT NOT NULL, status TEXT NOT NULL,
        notes TEXT NOT NULL, revision INTEGER NOT NULL, updated_at TEXT NOT NULL, updated_by TEXT NOT NULL
      );
      CREATE TABLE IF NOT EXISTS material_review_events (
        id INTEGER PRIMARY KEY AUTOINCREMENT, purchase_key TEXT NOT NULL,
        action TEXT NOT NULL, payload TEXT NOT NULL, actor TEXT NOT NULL, created_at TEXT NOT NULL
      );
    ''')
    con.execute('INSERT OR IGNORE INTO material_review_settings VALUES (?,?)', ('start_date', START_DATE))
    con.commit()


def actor_name(user):
    return user.get('username') or str(user.get('id') or 'Unknown')


def event(con, key, action, payload, actor):
    con.execute('INSERT INTO material_review_events (purchase_key,action,payload,actor,created_at) VALUES (?,?,?,?,?)',
                (key, action, json.dumps(payload), actor, datetime.now(ZoneInfo('America/Chicago')).isoformat(timespec='seconds')))


def invoice_sources(app, con):
    invoices = app.rows('''SELECT i.id,i.po_id,i.invoice_file,i.invoice_amount,i.uploaded_at,
        po.po_number,po.job_number,po.job_label,po.vendor,po.status AS po_status,
        COALESCE(NULLIF(po.customer_name,''),p.customer,'') AS customer_name
        FROM purchase_order_invoices i JOIN purchase_orders po ON po.id=i.po_id
        LEFT JOIN projects p ON p.id=po.project_id ORDER BY i.id''')
    documents = {}
    parsed_cache = []
    for invoice in invoices:
        path = (app.UPLOAD_DIR / invoice['invoice_file']).resolve()
        if path.parent != app.UPLOAD_DIR.resolve() or not path.is_file():
            invoice['sha256'] = 'missing:' + str(invoice['id'])
            documents[invoice['sha256']] = {'items': [], 'parser_error': 'Attached invoice file is unavailable in this copy of the application.'}
            continue
        sha = hashlib.sha256(path.read_bytes()).hexdigest()
        invoice['sha256'] = sha
        # Hash file content, not a filename or upload timestamp. Replaced files
        # must be parsed again; price/quantity corrections invalidate matches.
        key = rec.fingerprint([sha, PARSER_VERSION])
        old = con.execute('SELECT result_json FROM material_purchase_parse_cache WHERE cache_key=?', (key,)).fetchone()
        if old:
            documents[sha] = json.loads(old[0])
            continue
        try:
            result = app.import_vendor_invoice_pdf(path, 0, preview_details=True)
        except Exception as exc:
            result = {'items': [], 'parser_error': str(exc), 'parse_status': 'Needs parser review'}
        documents[sha] = result
        parsed_cache.append((key, path.name, json.dumps(result)))
    # Parse first, then write the cache in one short batch. OCR must not hold a
    # database write lock while other staff are editing POs or uploading files.
    if parsed_cache:
        con.executemany('INSERT OR REPLACE INTO material_purchase_parse_cache VALUES (?,?,?)', parsed_cache)
    return invoices, documents


def current(app, con):
    start = con.execute("SELECT value FROM material_review_settings WHERE key='start_date'").fetchone()[0]
    invoices, documents = invoice_sources(app, con)
    sales = rec.active_sales_rows(con)
    coverage = [r['day'] for r in con.execute('SELECT day FROM material_sales_coverage ORDER BY day')]
    catalog = app.rows('SELECT tablet_label,part_number,invoice_description FROM fieldwise_parts')
    all_buys = rec.compare_candidates(invoices, documents, sales, catalog, coverage)
    for b in all_buys:
        if not b.get('purchase_key'):
            b['purchase_key'] = rec.fingerprint(['unparsed', b['invoice_id'], b['invoice_file']])
    matches, stale = rec.valid_matches(con, all_buys, sales)
    buys = [b for b in all_buys if rec.in_forward_workflow(b, start)]
    allowed = {b['purchase_key'] for b in buys}
    used = Counter()
    for m in matches:
        used[m['purchase_key']] += m['bought_qty']
    issues = {r['purchase_key']: dict(r) for r in con.execute('SELECT * FROM material_line_issues')}
    today = datetime.now(ZoneInfo('America/Chicago')).date()
    for b in buys:
        b['base_state'] = b['state']
        b['confirmed_qty'] = used[b['purchase_key']]
        b['issue'] = issues.get(b['purchase_key'])
        b['matches'] = [m for m in matches if m['purchase_key'] == b['purchase_key']]
        if b['confirmed_qty'] and b['bought_qty']:
            b['state'] = 'Confirmed sold' if abs(b['confirmed_qty'] - b['bought_qty']) < 1e-8 else 'Partially confirmed'
        b.update(rec.purchase_timing(b, coverage, today))
        if b['invoice_date'] and str(b['invoice_date']) > today.isoformat():
            b['timing_state'] = 'Needs invoice date review — future purchase date'
        if b.get('issue') and b['issue']['status'] != 'Resolved':
            b['timing_state'] = 'Open investigation'
    return {'purchases': buys, 'sales': sales, 'coverage': coverage,
            'matches': [m for m in matches if m['purchase_key'] in allowed],
            'stale_matches': [m for m in stale if m['purchase_key'] in allowed],
            'orphan_issues': [i for k, i in issues.items() if k not in allowed],
            'start_date': start, 'excluded_review_rows': len(all_buys) - len(buys),
            'source_invoice_count': len({b['invoice_id'] for b in buys}),
            'purchase_snapshot': datetime.now(ZoneInfo('America/Chicago')).isoformat(timespec='seconds'),
            'states': dict(Counter(b['state'] for b in buys)),
            'timing_states': dict(Counter(b['timing_state'] for b in buys)),
            'imports': [dict(r) for r in con.execute('SELECT * FROM material_sales_imports ORDER BY id DESC LIMIT 30')]}


def save_issue(con, state, payload, actor):
    key = str(payload.get('purchase_key') or '')
    buy = next((b for b in state['purchases'] if b['purchase_key'] == key), None)
    old = con.execute('SELECT * FROM material_line_issues WHERE purchase_key=?', (key,)).fetchone()
    if not buy and old:
        # A replaced source must not erase or strand the investigation. Editing
        # its issue changes no quantities and does not relink the replacement.
        buy = json.loads(old['source_json'])
    if not buy:
        raise ReviewError('The purchase line changed. Refresh before saving the issue.', 409)
    if payload.get('expected_revision') != (old['revision'] if old else 0):
        raise ReviewError('This issue changed in another review. Refresh before saving.', 409)
    fields = {k: str(payload.get(k) or '').strip() for k in ('issue_group', 'description', 'notes', 'status')}
    if not fields['issue_group'] or not fields['description']:
        raise ReviewError('Enter an issue group and a description.')
    if len(fields['issue_group']) > 100 or any(len(fields[k]) > 4000 for k in ('description', 'notes')):
        raise ReviewError('Use a group of at most 100 characters and descriptions/notes of at most 4,000 characters.')
    if fields['status'] not in ISSUE_STATUSES:
        raise ReviewError('Choose a valid issue status.')
    rev = (old['revision'] if old else 0) + 1
    now = datetime.now(ZoneInfo('America/Chicago')).isoformat(timespec='seconds')
    con.execute('''INSERT INTO material_line_issues VALUES (?,?,?,?,?,?,?,?,?)
        ON CONFLICT(purchase_key) DO UPDATE SET source_json=excluded.source_json,
        issue_group=excluded.issue_group,description=excluded.description,status=excluded.status,
        notes=excluded.notes,revision=excluded.revision,updated_at=excluded.updated_at,updated_by=excluded.updated_by''',
        (key, json.dumps({k: v for k, v in buy.items() if k not in ('issue', 'candidates', 'matches')}),
         fields['issue_group'], fields['description'], fields['status'], fields['notes'], rev, now, actor))
    event(con, key, 'Issue updated', {'before': dict(old) if old else None, 'after': fields}, actor)
    return {'revision': rev}


def handle(app, handler, path, user):
    """Called after central application authentication/authorization."""
    try:
        if path == '/bought-vs-sold':
            return app.text_response(handler, material_review_ui.HTML)
        if path == '/bought-vs-sold-guide':
            return app.text_response(handler, (Path(__file__).resolve().parent / 'docs/bought-vs-sold-user-guide.html').read_text(encoding='utf-8'))
        actor = actor_name(user)
        with LOCK, closing(store(app)) as con:
            init_schema(con)
            if handler.command == 'GET':
                state = current(app, con)
                state['can_edit'] = app.can_edit_permission(user, 'po_review')
                state['events'] = [dict(r) for r in con.execute('SELECT * FROM material_review_events ORDER BY id DESC LIMIT 100')]
                con.commit()
                return app.json_response(handler, state)
            origin = handler.headers.get('Origin')
            if origin:
                from urllib.parse import urlparse
                if urlparse(origin).netloc != handler.headers.get('Host'):
                    raise ReviewError('Use this application to submit review changes.', 403)
            if int(handler.headers.get('Content-Length') or 0) > 20 * 1024 * 1024:
                raise ReviewError('Use an export smaller than 20 MB.')
            with con:
                if path.endswith('/upload'):
                    form = cgi.FieldStorage(fp=handler.rfile, headers=handler.headers,
                        environ={'REQUEST_METHOD': 'POST', 'CONTENT_TYPE': handler.headers.get('Content-Type')})
                    f = form['file'] if 'file' in form else None
                    if f is None or not getattr(f, 'filename', '') or Path(f.filename).suffix.lower() not in ('.xlsx', '.xlsm'):
                        raise ReviewError('Choose a Field Wise Excel export.')
                    raw = f.file.read()
                    digest = hashlib.sha256(raw).hexdigest()
                    folder = app.DATA_DIR / 'material-review-exports'
                    folder.mkdir(exist_ok=True)
                    export = folder / (digest + '.xlsx')
                    export.write_bytes(raw)
                    result = rec.import_sales_rows(con, rec.read_fieldwise_export(export),
                        form.getfirst('window_end'), Path(f.filename).name, digest, actor)
                else:
                    payload = app.parse_json(handler)
                    state = current(app, con)
                    if path.endswith('/confirm'):
                        for b in state['purchases']:
                            b['state'] = b['base_state']
                        result = {'id': rec.confirm_match(con, state['purchases'], state['sales'], payload, actor)}
                        event(con, payload['purchase_key'], 'Quantity confirmed', payload, actor)
                    elif path.endswith('/issue'):
                        result = save_issue(con, state, payload, actor)
                    elif path.endswith('/unconfirm'):
                        match = con.execute('SELECT * FROM material_confirmed_matches WHERE id=? AND active=1', (payload.get('id'),)).fetchone()
                        if not match:
                            raise ReviewError('Confirmation is no longer active.', 409)
                        con.execute('UPDATE material_confirmed_matches SET active=0 WHERE id=?', (match['id'],))
                        event(con, match['purchase_key'], 'Confirmation reversed', dict(match), actor)
                        result = {'ok': True}
                    else:
                        raise ReviewError('Not found', 404)
            return app.json_response(handler, result)
    except (ValueError, KeyError, TypeError) as exc:
        return app.json_response(handler, {'error': str(exc)}, getattr(exc, 'status', 400))


def integrate(html):
    anchor = '<button data-tab="closedPo" data-nav-area="po" data-nav-level="sub">Closed POs</button>'
    html = html.replace(anchor, anchor + '\n<button data-tab="boughtVsSold" data-nav-area="po" data-nav-level="sub">Bought vs Sold Review</button>', 1)
    html = html.replace('<section id="vendorPriceCatalog"', '''<section id="boughtVsSold" class="tab hidden">
      <iframe id="boughtVsSoldFrame" title="Bought vs Sold Review" style="width:100%;height:calc(100vh - 130px);min-height:750px;border:0;border-radius:8px"></iframe>
    </section><section id="vendorPriceCatalog"''', 1)
    html = html.replace("closedPo: 'po_review',", "closedPo: 'po_review', boughtVsSold: 'po_review',", 1)
    html = html.replace("'closedPo',", "'closedPo','boughtVsSold',", 1)
    html = html.replace("if (tabName === 'closedPo') loadClosedPos();", "if (tabName === 'closedPo') loadClosedPos();\n      if (tabName === 'boughtVsSold') { const f=document.getElementById('boughtVsSoldFrame'); if(!f.getAttribute('src')) f.src='/bought-vs-sold'; }", 1)
    return html
