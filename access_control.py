"""Server-side route and document authorization. Unknown protected routes deny access."""
import re
from urllib.parse import unquote, urlparse, parse_qs

PROJECT_READ = ('projects','project_setup','fieldwise','review_exceptions','vendor_invoices',
                'customer_billing','customer_dashboard','nte_tracking','archived_projects')

READ = {
    'job-order-report': ('job_order_report','po_requests','po_review'),
    'fieldwise-audit-omissions': ('fieldwise_audit',),
    'po-invoice-audit-omissions': ('po_review',),
    'vendor-price-catalogs': ('vendor_price_catalog',),
    'fieldwise-parts-catalogs': ('vendor_price_catalog',),
    'fieldwise-parts': ('vendor_price_catalog','po_requests','po_review'),
    'projects': PROJECT_READ, 'subprojects': PROJECT_READ, 'change-orders': PROJECT_READ,
    'rate-sets': PROJECT_READ, 'internal-rates': ('project_setup',),
    'summary': ('projects',), 'master-detail': ('projects',), 'subproject-detail': ('projects',),
    'customer-dashboard': ('customer_dashboard',), 'customer-invoices': ('customer_billing',),
    'vendor-invoice-lines': ('vendor_invoices',), 'vendor-invoice-allocations': ('vendor_invoices',),
    'cost-records': ('review_exceptions','nte_tracking','fieldwise'),
    'bid-summary': ('bids',), 'bids': ('bids',),
    'texas-financial-summary': ('texas_ops',), 'texas-upload-reminders': ('texas_reminders',),
    'billing-invoice-reminders': ('customer_billing',), 'po-untouched-reminders': ('po_review',),
    'cog-categories': ('cog_setup','po_requests','po_review','project_setup'),
    'company-calendar': ('company_calendar',),
}
WRITE = {
    'projects': ('project_setup',), 'subprojects': ('project_setup',), 'change-orders': ('project_setup',),
    'internal-rates': ('project_setup',), 'bids': ('bids',),
    'fieldwise-audit': ('fieldwise_audit',), 'fieldwise-audit-omissions': ('fieldwise_audit',),
    'po-invoice-audit': ('po_review',), 'po-invoice-audit-omissions': ('po_review',),
    'vendor-price-catalogs': ('vendor_price_catalog',), 'fieldwise-parts-catalogs': ('vendor_price_catalog',),
    'import-fieldwise': ('fieldwise',), 'import-vendor-invoice': ('vendor_invoices',),
    'customer-invoices': ('customer_billing',), 'customer-invoice-allocations': ('customer_billing',),
    'invoices': ('vendor_invoices',), 'vendor-invoice': ('vendor_invoices',),
    'purchase-order-invoices': ('po_review',), 'cost-records': ('review_exceptions',),
    'texas-financial-import': ('texas_ops',), 'texas-financial-delete': ('texas_ops',),
    'texas-ap': ('texas_ops',), 'texas-ap-delete': ('texas_ops',),
    'texas-upload-reminders': ('texas_reminders',), 'billing-invoice-reminders': ('customer_billing',),
    'po-untouched-reminders': ('po_review',), 'company-calendar': ('company_calendar',),
}
ADMIN = ('users','role-permissions','cog-access-users','cog-categories')

def requirement(method, path, query=None):
    """Return permission alternatives; empty tuple means unknown (deny), None means signed-in."""
    query = query or {}
    if method == 'GET' and path in ('/', '/api/me', '/api/home-alerts'):
        return None
    if method == 'POST' and path in ('/api/logout', '/api/change-password'):
        return None
    if path in ('/server-health','/developer-revision'):
        return ('admin',)
    if path.startswith('/api/admin/') or path in ('/api/server-health','/api/developer-revision','/api/activity'):
        return ('admin',)
    if path in ('/nte-weekly-report','/nte-weekly-report.pdf','/nte-budget-addition-approval.pdf','/nte-budget-reallocation-approval.pdf'):
        return ('nte_tracking',) if method == 'GET' else ()
    if not path.startswith('/api/'):
        return ()
    key = path.split('/')[2]
    if key in ADMIN and (key != 'cog-categories' or method != 'GET'):
        return ('admin',)
    if key.startswith('nte-'):
        return ('nte_tracking',)
    if key == 'mcc-quotes':
        return ('mcc_quotes',)
    if key.startswith('mcc-'):
        return ('mcc_quotes','mcc_quote_setup') if method == 'GET' else ('mcc_quote_setup',)
    if key == 'purchase-orders':
        if method == 'GET':
            return ('po_requests','po_review')
        if method == 'POST' and path == '/api/purchase-orders':
            return ('po_requests',)
        if method == 'POST' and path.endswith(('/pickup','/pick-ticket','/pick-ticket/close')):
            return ('po_requests','po_review')
        return ('admin',) if method == 'DELETE' else ('po_review',)
    if key == 'home-alerts' and path == '/api/home-alerts/job-invoice-ack':
        return ('customer_billing',)
    if key == 'imports':
        source = query.get('source', [''])[0]
        if method == 'GET':
            return ('vendor_invoices',) if source == 'Vendor Invoice' else ('fieldwise',) if source in ('Field Wise','Field Wise PDF') else ('fieldwise','vendor_invoices')
        # The payload source is checked separately before any database mutation.
        return ('fieldwise','vendor_invoices')
    if key == 'projects' and method == 'POST' and path.endswith('/restore'):
        return ('archived_projects',)
    if key == 'projects' and method == 'GET' and query.get('status', [''])[0] == 'archived':
        return ('archived_projects',)
    return (READ if method == 'GET' else WRITE).get(key, ())

def document_name(path):
    for prefix in ('/uploads/','/pdf-viewer/','/pdf-page/'):
        if path.startswith(prefix):
            name = unquote(path[len(prefix):])
            if prefix == '/pdf-page/':
                name = name.rsplit('/',1)[0]
            if not name or '/' in name or '\\' in name or name in ('.','..'):
                return ''
            return name
    return None

def po_visible(app, user, po):
    return bool(po and (app.can_view_permission(user,'po_review') or
                (app.can_view_permission(user,'po_requests') and po['requested_by_user_id'] == user['id'])))

def document_allowed(app, user, name):
    """Resolve an exact stored filename to its owner. Unreferenced uploads are private."""
    if not name:
        return False
    # PO invoices are office documents; field staff may see only their attachments/pickups.
    pos = app.rows('SELECT * FROM purchase_orders WHERE attachment_file=? OR pickup_file=?', (name,name))
    if any(po_visible(app,user,po) for po in pos):
        return True
    references = (
        ('purchase_orders', ('invoice_file',), ('po_review',)),
        ('purchase_order_invoices', ('invoice_file',), ('po_review',)),
        ('customer_invoices', ('invoice_file',), ('customer_billing','customer_dashboard')),
        ('financial_reports', ('source_file',), ('texas_ops',)),
        ('nte_bucket_additions', ('support_file','signed_file'), ('nte_tracking',)),
        ('nte_bucket_reallocations', ('support_file','signed_file'), ('nte_tracking',)),
        ('vendor_invoice_allocations', ('source_file',), ('vendor_invoices',)),
        ('vendor_price_catalogs', ('source_file',), ('vendor_price_catalog',)),
        ('fieldwise_parts_catalogs', ('source_file',), ('vendor_price_catalog',)),
        ('mcc_quote_items', ('vendor_quote_file','filter_vendor_quote_file'), ('mcc_quotes','mcc_quote_setup')),
        ('mcc_quote_lines', ('vendor_quote_file',), ('mcc_quotes','mcc_quote_setup')),
        ('mcc_vfd_prices', ('vendor_quote_file','filter_vendor_quote_file'), ('mcc_quotes','mcc_quote_setup')),
        ('mcc_vfd_filter_prices', ('vendor_quote_file',), ('mcc_quotes','mcc_quote_setup')),
    )
    for table, columns, permissions in references:
        if any(app.can_view_permission(user,p) for p in permissions):
            clause = ' OR '.join(f'{col}=?' for col in columns)
            if app.one(f'SELECT 1 AS found FROM {table} WHERE {clause} LIMIT 1', (name,)*len(columns)):
                return True
    for row in app.rows('SELECT source FROM cost_records WHERE source_file=?', (name,)):
        permissions = ('vendor_invoices',) if row['source'] == 'Vendor Invoice' else ('fieldwise','review_exceptions','fieldwise_audit','nte_tracking')
        if any(app.can_view_permission(user,p) for p in permissions):
            return True
    return False

def filter_payload(app, handler, payload):
    """Remove mixed-feature data that the requester is not permitted to read."""
    path = urlparse(handler.path).path
    if path not in ('/api/home-alerts','/api/imports','/api/cost-records','/api/purchase-orders'):
        return payload
    user = app.current_user(handler)
    if path == '/api/home-alerts':
        tabs = {'review':'review_exceptions','texasOps':'texas_ops','bids':'bids','officePo':'po_review','billing':'customer_billing'}
        result = dict(payload)
        result['alerts'] = [a for a in payload.get('alerts',[]) if app.can_view_permission(user,tabs.get(a.get('tab'), 'admin'))]
        result['attention_count'] = sum(1 for a in result['alerts'] if a['severity'] != 'ok' and a['count'])
        return result
    if path in ('/api/imports','/api/cost-records'):
        def allowed(row):
            source = row.get('source')
            if source == 'Vendor Invoice':
                return app.can_view_permission(user,'vendor_invoices') or (path == '/api/cost-records' and app.can_view_permission(user,'review_exceptions'))
            return any(app.can_view_permission(user,p) for p in ('fieldwise','review_exceptions','nte_tracking'))
        return [r for r in payload if allowed(r)]
    if path == '/api/purchase-orders' and not app.can_view_permission(user,'po_review'):
        return [{k:v for k,v in po.items() if 'invoice' not in k} for po in payload if po['requested_by_user_id'] == user['id']]
    return payload

def authorize(app, handler):
    parsed = urlparse(handler.path)
    path, method = parsed.path, handler.command
    if app.session_security.https_host(handler) and not app.session_security.secure_request(handler):
        app.json_response(handler,{'error':'This hostname requires the configured HTTPS proxy.'},403)
        return False
    if method == 'GET' and (path in ('/login','/manifest.json','/service-worker.js','/offline') or path.startswith('/brand/')):
        return True
    if method == 'POST' and path == '/api/login':
        return True
    user = app.current_user(handler)
    if not user:
        if method == 'GET' and not path.startswith('/api/'):
            app.redirect_response(handler,'/login')
        else:
            app.json_response(handler,{'error':'Login required'},401)
        return False
    if user.get('must_change_password') and not (
        (method=='GET' and path in ('/','/api/me')) or
        (method=='POST' and path in ('/api/change-password','/api/logout'))
    ):
        if method=='GET' and not path.startswith('/api/'):
            app.redirect_response(handler,'/')
        else:
            app.json_response(handler,{'error':'Change your temporary password before accessing company data.','code':'password_change_required'},403)
        return False
    name = document_name(path)
    if name is not None:
        allowed = method == 'GET' and document_allowed(app,user,name)
        if not allowed:
            app.text_response(handler,'Not found','text/plain',404)
        return allowed
    if path.startswith('/po/'):
        po = app.one('SELECT * FROM purchase_orders WHERE id=?', (path.rsplit('/',1)[-1],))
        if method == 'GET' and po_visible(app,user,po):
            return True
        app.text_response(handler,'Not found','text/plain',404)
        return False
    permissions = requirement(method,path,parse_qs(parsed.query))
    checker = app.can_view_permission if method == 'GET' else app.can_edit_permission
    if permissions is not None and not any(checker(user,p) for p in permissions):
        app.json_response(handler,{'error':'You do not have permission to access this resource.'},403)
        return False
    # Check ownership BEFORE parsing uploads or writing files.
    if method == 'POST' and re.fullmatch(r'/api/purchase-orders/[^/]+/(pickup|pick-ticket|pick-ticket/close)', path):
        po = app.one('SELECT * FROM purchase_orders WHERE id=?', (path.split('/')[3],))
        if not po or (not app.can_edit_permission(user,'po_review') and po['requested_by_user_id'] != user['id']):
            app.json_response(handler,{'error':'PO not found.'},404)
            return False
    return True
