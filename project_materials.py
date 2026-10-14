"""Read-only material visibility; no changes to financial records."""
from functools import lru_cache


@lru_cache(maxsize=256)
def invoice_items(app, filename, modified):
    path = app.upload_attachment_path(filename)
    return app.import_vendor_invoice_pdf(path, None, preview_only=True)


def list_materials(app, project_id):
    projects = app.rows("SELECT id,project_code,name FROM projects ORDER BY project_code,name")
    if not project_id:
        return {'projects':projects,'items':[],'warnings':[]}
    if not any(str(p['id']) == str(project_id) for p in projects):
        raise ValueError('Project not found.')
    jobs = app.rows('''SELECT job_number,code,name FROM subprojects WHERE project_id=?
        UNION ALL SELECT job_number,co_number AS code,title AS name FROM change_orders WHERE project_id=?''',(project_id,project_id))
    items, warnings, seen = [], [], set()
    invoices = app.rows('''SELECT poi.invoice_file,poi.uploaded_by_username,po.po_number,po.requested_by_username,
        po.job_number,po.job_label,po.vendor,po.id AS po_id
        FROM purchase_order_invoices poi JOIN purchase_orders po ON po.id=poi.po_id WHERE po.project_id=? ORDER BY poi.uploaded_at DESC''',(project_id,))
    by_file = {r['invoice_file']:r for r in invoices}
    costs = app.rows('''SELECT c.id,c.item,c.description,c.qty,c.vendor,c.record_date,c.ticket_or_invoice,c.source_file,
        COALESCE(co.job_number,sp.job_number,'') AS job_number,
        COALESCE(co.title,sp.name,'') AS job_label FROM cost_records c
        LEFT JOIN subprojects sp ON sp.id=c.subproject_id LEFT JOIN change_orders co ON co.id=c.change_order_id
        WHERE c.project_id=? AND c.source='Vendor Invoice' AND c.cost_type='Material'
        ORDER BY c.record_date DESC,c.id''',(project_id,))
    for row in costs:
        if row['item'] in ('PO Invoice','Vendor Invoice Total'):
            continue
        linked=by_file.get(row['source_file'],{})
        items.append({**row,'unit':'','buyer':linked.get('requested_by_username') or 'Not recorded','po_number':linked.get('po_number') or '', 'uploaded_by':linked.get('uploaded_by_username') or ''})
        seen.add(row['source_file'])
    files = list(by_file)
    files.extend(r['source_file'] for r in costs if r['source_file'] and r['source_file'] not in files)
    for filename in files:
        if filename in seen: continue
        linked=by_file.get(filename,{})
        path=app.upload_attachment_path(filename)
        try:
            if not path: raise ValueError('Invoice file is missing.')
            parsed=invoice_items(app,filename,path.stat().st_mtime_ns)
            if not parsed['items']: raise ValueError('Item descriptions and quantities could not be extracted.')
            fallback=next((r for r in costs if r['source_file']==filename),{})
            for line in parsed['items']:
                items.append({**line,'vendor':linked.get('vendor') or parsed['vendor'], 'record_date':parsed['date'],
                    'ticket_or_invoice':parsed['invoice_number'],'job_number':linked.get('job_number') or fallback.get('job_number',''),
                    'job_label':linked.get('job_label') or fallback.get('job_label',''),'buyer':linked.get('requested_by_username') or 'Not recorded',
                    'po_number':linked.get('po_number') or '', 'uploaded_by':linked.get('uploaded_by_username') or ''})
        except Exception:
            warnings.append({'po_number':linked.get('po_number') or '', 'job_number':linked.get('job_number') or '', 'vendor':linked.get('vendor') or '', 'message':'Invoice items unavailable; office review or a clearer invoice is needed.'})
    # Field-facing result deliberately excludes invoice PDFs, prices and totals.
    allowed=('item','description','qty','unit','vendor','record_date','ticket_or_invoice','job_number','job_label','buyer','po_number','uploaded_by')
    return {'projects':projects,'jobs':jobs,'items':[{k:row.get(k,'') for k in allowed} for row in items],'warnings':warnings}


SCRIPT = '''
    let projectMaterialData = {items:[],warnings:[]};
    let materialsLoadSequence = 0;
    async function loadProjectMaterials() {
      const select=document.getElementById('materialsProject');
      const id=state.projectId || select.value || '';
      const sequence=++materialsLoadSequence;
      select.parentElement.classList.toggle('hidden', !isFieldPoOnly());
      document.getElementById('materialsCount').textContent='Loading invoice items...';
      try {
        const result=await api('/api/project-materials?project_id='+encodeURIComponent(id));
        if (sequence !== materialsLoadSequence) return;
        projectMaterialData=result;
        select.innerHTML='<option value="">Choose a project</option>'+projectMaterialData.projects.map(p=>`<option value="${p.id}">${htmlEscape(p.project_code+' — '+p.name)}</option>`).join('');
        select.value=String(id);
        const jobs=[...new Set(projectMaterialData.items.map(r=>r.job_number).filter(Boolean))].sort();
        document.getElementById('materialsJob').innerHTML='<option value="">All jobs / orders</option>'+jobs.map(j=>{
          const metadata=(projectMaterialData.jobs || []).filter(row=>String(row.job_number)===String(j));
          const details=[...new Set(metadata.map(row=>[row.code,row.name].filter(Boolean).join(' — ')).filter(Boolean))].join(' / ')
            || projectMaterialData.items.find(row=>row.job_number===j)?.job_label || '';
          return `<option value="${htmlEscape(j)}">${htmlEscape([j,details].filter(Boolean).join(' — '))}</option>`;
        }).join('');
        renderProjectMaterials();
      } catch(err) { document.getElementById('materialsCount').textContent=err.message; }
    }
    function renderProjectMaterials() {
      const query=document.getElementById('materialsSearch').value.trim().toLowerCase();
      const job=document.getElementById('materialsJob').value;
      const rows=projectMaterialData.items.filter(r=>(!job || r.job_number===job) && [r.item,r.description,r.vendor,r.buyer,r.job_number,r.job_label,r.po_number].join(' ').toLowerCase().includes(query));
      document.getElementById('materialsCount').textContent=`${rows.length} of ${projectMaterialData.items.length} invoice item lines`;
      document.getElementById('materialsWarnings').textContent=projectMaterialData.warnings.length ? `${projectMaterialData.warnings.length} invoice(s) need office review: `+projectMaterialData.warnings.map(w=>[w.po_number,w.job_number,w.vendor].filter(Boolean).join(' / ')).join('; ') : '';
      document.getElementById('materialsTable').innerHTML='<thead><tr><th>Date</th><th>Job / Order</th><th>Item / Description</th><th>Quantity</th><th>Vendor</th><th>Bought by (PO requester)</th><th>PO / Invoice</th></tr></thead><tbody>'+rows.map((r,index)=>`<tr><td>${htmlEscape(r.record_date)}</td><td>${htmlEscape(r.job_number)}<div class="muted">${htmlEscape(r.job_label)}</div></td><td><strong class="material-item-code">${htmlEscape(r.item)}</strong><div class="material-description-preview">${htmlEscape(r.description)}</div>${String(r.description || '').length > 0 ? `<button class="btn" type="button" data-material-description="${index}">View full description</button>` : ''}</td><td>${htmlEscape(r.qty)} ${htmlEscape(r.unit)}</td><td>${htmlEscape(r.vendor)}</td><td>${htmlEscape(r.buyer)}</td><td>${htmlEscape(r.po_number)}<div class="muted">${htmlEscape(r.ticket_or_invoice)}</div></td></tr>`).join('')+'</tbody>';
      document.querySelectorAll('[data-material-description]').forEach(btn => btn.onclick = () => {
        const row=rows[Number(btn.dataset.materialDescription)];
        document.getElementById('materialDescriptionTitle').textContent=row.item || 'Material description';
        document.getElementById('materialDescriptionText').textContent=row.description || '';
        document.getElementById('materialDescriptionDialog').showModal();
      });
      if (!rows.length) document.getElementById('materialsCount').textContent += ' — No readable invoice items match. Items appear after an invoice is uploaded and linked to this project.';
    }
    document.getElementById('materialDescriptionClose').onclick=()=>document.getElementById('materialDescriptionDialog').close();
    document.getElementById('materialsProject').onchange=()=>{
      state.projectId=Number(document.getElementById('materialsProject').value) || null;
      loadProjectMaterials();
    };
    document.getElementById('materialsRefresh').onclick=loadProjectMaterials;
    document.getElementById('materialsSearch').oninput=renderProjectMaterials;
    document.getElementById('materialsJob').onchange=renderProjectMaterials;
'''

