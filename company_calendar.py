"""Company scheduling storage and dashboard interface."""
from datetime import date, datetime, timedelta
import calendar
import math

KINDS = ('Project start', 'Project completion', 'Customer milestone', 'Company event', 'Time off', 'Other')
STATUSES = ('Tentative', 'Scheduled', 'Complete', 'Cancelled')
PREVIEW_USERNAME = 'rstewart@twinpeakselectrical.net'

def preview_allowed(user):
    return bool(user and user.get('active', 1) and str(user.get('username') or '').strip().lower() == PREVIEW_USERNAME)

def initialize(con):
    con.executescript('''
        CREATE TABLE IF NOT EXISTS company_calendar_events (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT NOT NULL, category TEXT NOT NULL, status TEXT NOT NULL,
            start_date TEXT NOT NULL, end_date TEXT NOT NULL,
            project_id INTEGER REFERENCES projects(id) ON DELETE SET NULL,
            owner TEXT NOT NULL DEFAULT '', location TEXT NOT NULL DEFAULT '',
            crew_count INTEGER NOT NULL DEFAULT 0,
            leadership_hours REAL NOT NULL DEFAULT 0, notes TEXT NOT NULL DEFAULT '',
            created_by INTEGER REFERENCES users(id) ON DELETE SET NULL,
            created_at TEXT NOT NULL, updated_at TEXT NOT NULL
        );
        CREATE INDEX IF NOT EXISTS company_calendar_dates ON company_calendar_events(start_date, end_date);
    ''')
    columns = {row[1] for row in con.execute('PRAGMA table_info(company_calendar_events)')}
    for name, definition in (('repeat_frequency', "TEXT NOT NULL DEFAULT 'None'"), ('repeat_interval', 'INTEGER NOT NULL DEFAULT 1'), ('repeat_until', "TEXT NOT NULL DEFAULT ''")):
        if name not in columns:
            con.execute(f'ALTER TABLE company_calendar_events ADD COLUMN {name} {definition}')

def occurrences(events, month):
    first = date.fromisoformat(month + '-01')
    last = first.replace(day=calendar.monthrange(first.year, first.month)[1])
    result = []
    for event in events:
        start = date.fromisoformat(event['start_date'])
        duration = date.fromisoformat(event['end_date']) - start
        frequency = event.get('repeat_frequency', 'None')
        until = date.fromisoformat(event['repeat_until']) if frequency != 'None' else start
        interval = event.get('repeat_interval', 1)
        for n in range(4000):
            if frequency in ('Daily', 'Weekly'):
                current = start + timedelta(days=n * interval * (7 if frequency == 'Weekly' else 1))
            elif frequency in ('Monthly', 'Yearly'):
                offset = n * interval * (12 if frequency == 'Yearly' else 1)
                year, m = divmod(start.year * 12 + start.month - 1 + offset, 12)
                if year > 9999:
                    break
                current = date(year, m+1, min(start.day, calendar.monthrange(year, m+1)[1]))
            else:
                current = start
            if current > until or current > last:
                break
            if current + duration >= first:
                result.append(dict(event, start_date=current.isoformat(), end_date=(current+duration).isoformat(),
                                   series_start_date=event['start_date'], series_end_date=event['end_date']))
            if frequency == 'None':
                break
    return sorted(result, key=lambda e: (e['start_date'], e['title'], e['id']))

def save(con, data, actor):
    title = str(data.get('title') or '').strip()
    if not title or len(title) > 200:
        raise ValueError('Enter a title of 1 to 200 characters.')
    category, status = data.get('category'), data.get('status')
    if category not in KINDS or status not in STATUSES:
        raise ValueError('Choose a valid event type and status.')
    try:
        start = date.fromisoformat(str(data.get('start_date')))
        end = date.fromisoformat(str(data.get('end_date')))
        crew = int(str(data.get('crew_count') or '0'))
        hours = float(data.get('leadership_hours') or 0)
    except (ValueError, TypeError):
        raise ValueError('Enter valid dates, a whole crew count, and leadership hours.')
    if end < start:
        raise ValueError('End date must be on or after the start date.')
    frequency = data.get('repeat_frequency') or 'None'
    if frequency not in ('None', 'Daily', 'Weekly', 'Monthly', 'Yearly'):
        raise ValueError('Choose a valid repeat frequency.')
    interval, until = 1, ''
    if frequency != 'None':
        try:
            interval = int(str(data.get('repeat_interval') or '1'))
            until_date = date.fromisoformat(str(data.get('repeat_until') or ''))
        except (ValueError, TypeError):
            raise ValueError('Enter a whole repeat interval and a repeat end date.')
        if not 1 <= interval <= 52 or not start <= until_date <= start + timedelta(days=3650):
            raise ValueError('Repeat interval must be 1–52, with a repeat end date within ten years of the start.')
        until = until_date.isoformat()
    if not 0 <= crew <= 10000 or not math.isfinite(hours) or not 0 <= hours <= 10000:
        raise ValueError('Crew and leadership hours must be between 0 and 10,000.')
    project_id = data.get('project_id') or None
    if project_id and not con.execute('SELECT id FROM projects WHERE id = ?', (project_id,)).fetchone():
        raise ValueError('The selected project no longer exists.')
    values = (title, category, status, start.isoformat(), end.isoformat(), project_id,
              str(data.get('owner') or '').strip(), str(data.get('location') or '').strip(),
              crew, hours, str(data.get('notes') or '').strip(), frequency, interval, until)
    now = datetime.now().isoformat(timespec='seconds')
    if data.get('id'):
        cur = con.execute('''UPDATE company_calendar_events SET
            title=?, category=?, status=?, start_date=?, end_date=?, project_id=?,
            owner=?, location=?, crew_count=?, leadership_hours=?, notes=?, repeat_frequency=?, repeat_interval=?, repeat_until=?, updated_at=? WHERE id=?''',
            (*values, now, data['id']))
        if not cur.rowcount:
            raise LookupError('Calendar event not found.')
        return int(data['id'])
    return con.execute('''INSERT INTO company_calendar_events
        (title,category,status,start_date,end_date,project_id,owner,location,crew_count,
         leadership_hours,notes,repeat_frequency,repeat_interval,repeat_until,created_by,created_at,updated_at)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''', (*values, actor['id'], now, now)).lastrowid

SECTION = r'''
<style>
 #calendarEditor { width:min(760px,calc(100vw - 32px)); max-height:calc(100dvh - 48px); overflow:auto; margin:auto; padding:24px; background:var(--panel); color:var(--text); border:1px solid var(--line); border-radius:12px; box-shadow:0 24px 80px rgba(0,0,0,.4); }
 #calendarEditor:not([open]) { display:none; }
 #calendarEditor::backdrop { background:rgba(15,23,42,.6); }
 #calendarEditor .calendar-fields { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:12px; }
 @media(max-width:600px) { #calendarEditor { padding:16px; } #calendarEditor .calendar-fields { grid-template-columns:1fr; } }
</style>
<section id="companyCalendar" class="tab hidden">
 <div class="panel">
  <div class="section-head"><div><h2>Work Calendar</h2><p class="muted">Plan project dates, customer milestones, company events, and resource commitments.</p></div><button class="btn" id="calendarNew">Add event</button></div>
  <div style="display:flex;gap:8px;flex-wrap:wrap;align-items:center">
   <button class="btn" id="calendarPrev" aria-label="Previous month">Previous</button><button class="btn" id="calendarToday">Today</button><button class="btn" id="calendarNext" aria-label="Next month">Next</button>
   <label>Month <input type="month" id="calendarMonth"></label>
   <label>View <select id="calendarView"><option>Month</option><option>Agenda</option></select></label>
   <label>Event type <select id="calendarCategory"><option value="">All types</option></select></label>
   <label>Project <select id="calendarProject"><option value="">All projects</option></select></label>
   <label>Status <select id="calendarStatus"><option value="">All statuses</option></select></label>
  </div>
  <p class="muted">Crew and leadership hours are planned per day, including weekends. Totals exclude completed and cancelled events. Filters also apply to totals.</p>
  <p id="calendarMessage" role="status"></p><div id="calendarSummary"></div>
  <div style="overflow:auto"><div id="calendarContent"></div></div>
 </div>
 <dialog id="calendarEditor" aria-labelledby="calendarEditorTitle">
  <div class="section-head"><h2 id="calendarEditorTitle">Add event</h2><button class="btn" type="button" id="calendarClose" aria-label="Close event details">&#215;</button></div>
  <form id="calendarForm">
   <input type="hidden" name="id">
   <label>Event title <input name="title" required maxlength="200"></label>
   <div class="calendar-fields">
    <label>Event type <select name="category"></select></label><label>Status <select name="status"></select></label>
    <label>Start date <input type="date" name="start_date" required></label><label>End date <input type="date" name="end_date" required></label>
    <label>Repeat <select name="repeat_frequency"><option>None</option><option>Daily</option><option>Weekly</option><option>Monthly</option><option>Yearly</option></select></label>
    <label>Repeat every <input name="repeat_interval" type="number" min="1" max="52" step="1" value="1"><span class="muted">Use 2 with Weekly for every other week.</span></label>
    <label>Repeat through <input name="repeat_until" type="date"></label>
    <label>Project <select name="project_id"><option value="">Company / no project</option></select></label>
    <label>Owner / leadership <input name="owner" placeholder="Person or team responsible"></label>
    <label>Location <input name="location"></label>
    <label>Crew needed per day <input name="crew_count" type="number" min="0" max="10000" step="1" value="0"></label>
    <label>Leadership hours per day <input name="leadership_hours" type="number" min="0" max="10000" step="0.25" value="0"></label>
   </div>
   <p class="muted" id="calendarRepeatHelp">Repeats keep the same event duration. Monthly dates that do not exist use the last day of that month. Editing or changing status applies to the entire series.</p>
   <label>Notes / scheduling dependencies <textarea name="notes" rows="3"></textarea></label>
   <button class="btn" type="submit" id="calendarSave">Save event</button> <button class="btn" type="button" id="calendarCancel">Close</button>
   <p id="calendarFormMessage" role="status"></p>
  </form>
 </dialog>
</section>
'''

SCRIPT = r'''
    const calendarKinds = ['Project start','Project completion','Customer milestone','Company event','Time off','Other'];
    const calendarStatuses = ['Tentative','Scheduled','Complete','Cancelled'];
    let calendarEvents = [], calendarRequest = 0;
    const calEl = id => document.getElementById(id);
    const calDate = d => `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
    const calForm = calEl('calendarForm');
    const calOptions = items => items.map(x => `<option>${htmlEscape(x)}</option>`).join('');
    calEl('calendarCategory').insertAdjacentHTML('beforeend', calOptions(calendarKinds));
    calEl('calendarStatus').insertAdjacentHTML('beforeend', calOptions(calendarStatuses));
    calForm.elements.category.innerHTML = calOptions(calendarKinds);
    calForm.elements.status.innerHTML = calOptions(calendarStatuses);
    calEl('calendarMonth').value = calDate(new Date()).slice(0,7);
    async function loadCompanyCalendar() {
      const request = ++calendarRequest;
      try {
        const data = await api('/api/company-calendar?month='+encodeURIComponent(calEl('calendarMonth').value));
        if (request !== calendarRequest) return;
        calendarEvents = data.events;
        const opts = data.projects.map(p => `<option value="${p.id}">${htmlEscape(p.project_code+' — '+p.name)}</option>`).join('');
        const selected = calEl('calendarProject').value;
        calEl('calendarProject').innerHTML = '<option value="">All projects</option>'+opts;
        calEl('calendarProject').value = selected;
        const editedProject = calForm.elements.project_id.value;
        calForm.elements.project_id.innerHTML = '<option value="">Company / no project</option>'+opts;
        calForm.elements.project_id.value = editedProject;
        calEl('calendarNew').classList.toggle('hidden', !canEditPermission('company_calendar'));
        renderCompanyCalendar();
        calEl('calendarMessage').textContent = '';
      } catch (err) { calEl('calendarMessage').textContent = err.message; }
    }
    function renderCompanyCalendar() {
      const month = calEl('calendarMonth').value;
      if (!month) return;
      const [year, m] = month.split('-').map(Number), first = new Date(year,m-1,1), last = new Date(year,m,0);
      const events = calendarEvents.filter(e => (!calEl('calendarCategory').value || e.category === calEl('calendarCategory').value) && (!calEl('calendarProject').value || String(e.project_id) === calEl('calendarProject').value) && (!calEl('calendarStatus').value || e.status === calEl('calendarStatus').value));
      const inMonth = events.filter(e => e.start_date <= calDate(last) && e.end_date >= calDate(first));
      const active = e => !['Complete','Cancelled'].includes(e.status);
      let peakCrew=0, peakHours=0;
      for(let day=1; day<=last.getDate(); day++) {
        const key=calDate(new Date(year,m-1,day)), daily=inMonth.filter(e=>active(e)&&e.start_date<=key&&e.end_date>=key);
        peakCrew=Math.max(peakCrew,daily.reduce((n,e)=>n+e.crew_count,0));
        peakHours=Math.max(peakHours,daily.reduce((n,e)=>n+e.leadership_hours,0));
      }
      calEl('calendarSummary').textContent = `${inMonth.length} events this month · Peak daily crew: ${peakCrew} · Peak daily leadership hours: ${peakHours}`;
      const eventButton = e => `<button type="button" class="btn" data-calendar-event="${e.id}" style="display:block;width:100%;text-align:left;margin:4px 0;white-space:normal;border-left:4px solid ${['#2563eb','#16a34a','#d97706','#9333ea','#dc2626','#64748b'][calendarKinds.indexOf(e.category)]}">${htmlEscape(e.title)}<br><small>${htmlEscape(e.category)} · ${htmlEscape(e.status)}</small></button>`;
      if(calEl('calendarView').value === 'Agenda') {
        calEl('calendarContent').innerHTML = inMonth.length ? inMonth.map(e => `<div style="padding:12px;border-bottom:1px solid #94a3b8">${eventButton(e)}<div>${htmlEscape(e.start_date)} to ${htmlEscape(e.end_date)} · ${htmlEscape(e.project_name || 'Company')} · ${htmlEscape(e.owner || 'Owner unassigned')}</div><div>Crew/day: ${e.crew_count} · Leadership hours/day: ${e.leadership_hours} · ${htmlEscape(e.location)}</div><p style="white-space:pre-wrap">${htmlEscape(e.notes)}</p></div>`).join('') : '<p>No events match this month and filters.</p>';
      } else {
        let cells = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'].map(x=>`<div style="padding:8px;font-weight:bold">${x}</div>`).join('');
        for(let n=0;n<first.getDay();n++) cells+='<div></div>';
        for(let day=1;day<=last.getDate();day++) {
          const key=calDate(new Date(year,m-1,day)), daily=inMonth.filter(e=>e.start_date<=key&&e.end_date>=key), planned=daily.filter(active);
          cells+=`<div style="min-height:130px;padding:8px;border:1px solid #94a3b8;${key===calDate(new Date())?'outline:2px solid #2563eb;outline-offset:-2px;':''}"><strong>${day}</strong>${daily.map(eventButton).join('')}${planned.length?`<small>Crew: ${planned.reduce((n,e)=>n+e.crew_count,0)} · Leadership: ${planned.reduce((n,e)=>n+e.leadership_hours,0)}h</small>`:''}</div>`;
        }
        calEl('calendarContent').innerHTML=`<div style="display:grid;grid-template-columns:repeat(7,minmax(140px,1fr));margin-top:12px">${cells}</div>`;
      }
      calEl('calendarContent').querySelectorAll('[data-calendar-event]').forEach(b=>b.onclick=()=>editCalendarEvent(calendarEvents.find(e=>e.id===Number(b.dataset.calendarEvent))));
    }
    function editCalendarEvent(event) {
      calForm.reset();
      const defaults = {id:'',title:'',category:'Project start',status:'Scheduled',start_date:calDate(new Date()),end_date:calDate(new Date()),project_id:'',owner:'',location:'',crew_count:0,leadership_hours:0,notes:'',repeat_frequency:'None',repeat_interval:1,repeat_until:''};
      const values = event ? {...defaults,...event,start_date:event.series_start_date || event.start_date,end_date:event.series_end_date || event.end_date} : defaults;
      Object.entries(values).forEach(([key,value])=>{if(calForm.elements.namedItem(key)) calForm.elements.namedItem(key).value=value ?? '';});
      const editable=canEditPermission('company_calendar');
      Array.from(calForm.elements).forEach(el=>{if(!['calendarCancel','calendarSave'].includes(el.id))el.disabled=!editable;});
      calEl('calendarSave').classList.toggle('hidden',!editable);
      calEl('calendarEditorTitle').textContent=event?(values.repeat_frequency !== 'None'?'Recurring event — edit series':'Event details'):'Add event';
      updateCalendarRepeat();
      calEl('calendarFormMessage').textContent='';
      calEl('calendarEditor').showModal();
      if(editable) calForm.elements.title.focus();
      else calEl('calendarClose').focus();
    }
    calEl('calendarNew').onclick=async()=>{if(await confirmDiscard()) {markSaved();editCalendarEvent();}};
    async function closeCalendarEditor() {
      if(hasUnsavedChanges && !window.confirm('Discard your unsaved event changes?')) return;
      calEl('calendarEditor').close();markSaved();
    }
    calEl('calendarCancel').onclick=closeCalendarEditor;
    calEl('calendarClose').onclick=closeCalendarEditor;
    calEl('calendarEditor').addEventListener('cancel',event=>{event.preventDefault();closeCalendarEditor();});
    function updateCalendarRepeat() {
      const repeating=calForm.elements.repeat_frequency.value !== 'None';
      calForm.elements.repeat_until.required=repeating;
      calForm.elements.repeat_until.disabled=!repeating || !canEditPermission('company_calendar');
      calForm.elements.repeat_interval.disabled=!repeating || !canEditPermission('company_calendar');
      calEl('calendarRepeatHelp').classList.toggle('hidden',!repeating);
    }
    calForm.elements.repeat_frequency.onchange=updateCalendarRepeat;
    calForm.elements.start_date.onchange=()=>{if(calForm.elements.end_date.value<calForm.elements.start_date.value)calForm.elements.end_date.value=calForm.elements.start_date.value;};
    calForm.onsubmit=async e=>{
      e.preventDefault(); calEl('calendarSave').disabled=true;
      try {
        await api('/api/company-calendar',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(formDataObj(calForm))});
        markSaved();calEl('calendarEditor').close();await loadCompanyCalendar();
      } catch(err) { calEl('calendarFormMessage').textContent=err.message; }
      finally {calEl('calendarSave').disabled=false;}
    };
    ['calendarView','calendarCategory','calendarProject','calendarStatus'].forEach(id=>calEl(id).onchange=renderCompanyCalendar);
    calEl('calendarMonth').onchange=loadCompanyCalendar;
    function moveCalendarMonth(delta) {const [y,m]=calEl('calendarMonth').value.split('-').map(Number);calEl('calendarMonth').value=calDate(new Date(y,m-1+delta,1)).slice(0,7);loadCompanyCalendar();}
    calEl('calendarPrev').onclick=()=>moveCalendarMonth(-1);
    calEl('calendarNext').onclick=()=>moveCalendarMonth(1);
    calEl('calendarToday').onclick=()=>{calEl('calendarMonth').value=calDate(new Date()).slice(0,7);loadCompanyCalendar();};
'''

def integrate(html):
    html = html.replace('<button data-tab="dashboard" data-nav-area="main"', '<button data-tab="companyCalendar" data-nav-area="main" data-nav-level="main">Work Calendar</button>\n        <button data-tab="dashboard" data-nav-area="main"', 1)
    html = html.replace('<section id="home" class="tab">', SECTION + '\n<section id="home" class="tab">', 1)
    html = html.replace("dashboard: 'projects',", "companyCalendar: 'company_calendar',\n      dashboard: 'projects',", 1)
    html = html.replace("if (!key || state.currentUser?.role === 'Admin') return true;", "if (key === 'company_calendar') return !!state.currentUser?.company_calendar_preview;\n      if (!key || state.currentUser?.role === 'Admin') return true;")
    html = html.replace("if (tabName === 'home') loadHomeAlerts();", "if (tabName === 'companyCalendar') loadCompanyCalendar();\n      if (tabName === 'home') loadHomeAlerts();", 1)
    html = html.replace("return target.closest('.cost-filter')", "return target.closest('#companyCalendar') && !target.closest('#calendarForm') || target.closest('.cost-filter')", 1)
    position = html.rfind('    updateCoPricingFields();')
    html = html[:position] + SCRIPT + '\n' + html[position:]
    return html
