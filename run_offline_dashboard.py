"""Launch an isolated dashboard copy with the downloaded October 27 data.

Re-running keeps local edits. It never refreshes from or writes to the live server,
and does not start the application's email/reminder background jobs.
"""
import json
import shutil
import sqlite3
from pathlib import Path
from http.server import ThreadingHTTPServer

import app
import material_reconciliation as rec
import material_review

ROOT = Path(__file__).resolve().parent
FOLDER = ROOT / 'outputs' / 'offline-dashboard'
DATA = FOLDER / 'data'
SOURCE = ROOT / 'outputs' / 'po-invoice-parser-review-20261027'
REVIEW = ROOT / 'outputs' / 'daily-material-comparison-20261027'


def prepare():
    DATA.mkdir(parents=True, exist_ok=True)
    (DATA / 'uploads').mkdir(exist_ok=True)
    (DATA / 'brand').mkdir(exist_ok=True)
    target = DATA / 'pm_tracker.sqlite3'
    if not target.exists():
        with sqlite3.connect(ROOT / 'tmp/po-invoice-parser-review-20261027/server-snapshot.sqlite3') as source:
            with sqlite3.connect(target) as dest:
                source.backup(dest)
    for path in (SOURCE / 'invoices').iterdir():
        dest = DATA / 'uploads' / path.name
        if path.is_file() and not dest.exists():
            shutil.copy2(path, dest)
    if (ROOT / 'data/brand').exists():
        for path in (ROOT / 'data/brand').iterdir():
            if path.is_file() and not (DATA / 'brand' / path.name).exists():
                shutil.copy2(path, DATA / 'brand' / path.name)
    state = DATA / 'material-review.sqlite3'
    if not state.exists():
        with sqlite3.connect(REVIEW / 'comparison.sqlite3') as source:
            with sqlite3.connect(state) as dest:
                source.backup(dest)
        config = json.loads((REVIEW / 'purchase-snapshot.json').read_text(encoding='utf-8'))
        with sqlite3.connect(state) as con:
            for sha, document in config['documents'].items():
                key = rec.fingerprint([sha, material_review.PARSER_VERSION])
                con.execute('INSERT OR REPLACE INTO material_purchase_parse_cache VALUES (?,?,?)',
                            (key, '', json.dumps(document)))
    return target


def main():
    target = prepare()
    app.DATA_DIR = DATA
    app.UPLOAD_DIR = DATA / 'uploads'
    app.BRAND_DIR = DATA / 'brand'
    app.DB_PATH = target
    app.HOST = '127.0.0.1'
    app.PORT = 8774
    # Keep this a local copy; no bid spreadsheet refresh or reminder jobs.
    app.seed_bid_tracker_from_workbook = lambda *a, **kw: None
    app.SMTP_HOST = ''
    app.INVENTORYCLOUD_API_TOKEN = ''
    app.init_db()
    with sqlite3.connect(target) as con:
        seeded = con.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name='material_review_settings'").fetchone()
        if not seeded:
            material_review.init_schema(con)
            seed_path = DATA / 'material-review.sqlite3'
            con.execute('ATTACH DATABASE ? AS seed', (str(seed_path),))
            tables = ('material_sales_imports', 'material_sales_tickets', 'material_sales_ticket_history',
                      'material_sales_coverage', 'material_purchase_parse_cache', 'material_confirmed_matches',
                      'material_line_issues', 'material_review_events')
            for table in tables:
                if con.execute('SELECT 1 FROM seed.sqlite_master WHERE type=\'table\' AND name=?', (table,)).fetchone():
                    con.execute(f'INSERT OR IGNORE INTO main.{table} SELECT * FROM seed.{table}')
            con.commit()
    banner = '<div style="padding:8px 16px;background:#fff0bd;color:#172231;font-weight:600">OFFLINE DASHBOARD · Local data copy · Changes stay on this PC · Reminder emails disabled</div>'
    app.HTML = app.HTML.replace('<body>', '<body>' + banner, 1)
    print('Offline dashboard: http://127.0.0.1:8774/?tab=boughtVsSold', flush=True)
    ThreadingHTTPServer((app.HOST, app.PORT), app.Handler).serve_forever()


if __name__ == '__main__':
    main()
