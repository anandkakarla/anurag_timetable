"""v71 = v70 + Faculty busy optional Room column: EXISTING_LOAD_GAP charges a class next to existing load only when it is in a different building from that existing-load room (blank Room = unknown, always charged as before).
v70 = v69 + MAX_LAB_PERIODS_PER_DAY: lab periods per section/batch per day above N are penalised per period (2+2 fine, 3+2 penalised once).
v69 = v68 + MAX_LABS_PER_DAY: at most N lab blocks per section/batch per day; courses listed with LAB_DAY_LIMIT_EXCLUDE (scope Course) are not counted.
v68 = v67, display only: existing-load blocks are labelled "Existing load" (or the Faculty busy row's optional Note, e.g. "MBA: Data Mgmt & SQL") instead of "First-year class".
v67 = v66 + CONSECUTIVE_DAYS_PER_COURSE: the days used by a course's labs must be consecutive (e.g. Wed-Thu, not Mon-Thu).
v66 = v65 + SAME_LAB_ROOM: all sessions of a lab course for one section (every batch, every slot) should use the same room.
v65 = v64 + --stop-at-first: stop as soon as the first complete timetable is found (for quick feasibility tests).
v64 = v63 + MAX_DAYS_PER_COURSE (all lab sessions of a course, every section, on at most N days) + FACULTY_SLOTS_LAB_EXEMPT (the teacher's allowed-slot rule does not apply to labs).
v63 = v62 + NO_FULL_DAY_LABS: a section/batch with a lab block in the morning (P1-P4) and another in the afternoon on the same day is penalised.
v62 = v61 + COURSE_ALLOWED_SLOTS: a course (scope Course, target = course code) may only be placed in the listed Day:Pn slots.
v61 = v60 + MAX_LABS_PER_MORNING (a section/batch has at most N lab blocks in P1-P4 per day) + Department/Section rows override the Global row of the same rule for that department/section (different weights per department).
v60 = v59 + --free-depts: with --hint, every class outside the listed departments is fixed to its time and room in the hint timetable; only the listed departments are re-solved. --free-faculty: classes taught by these teachers are also left free.
v59 = v58 + NO_FACULTY_FLOOR_HOP / NO_SECTION_FLOOR_HOP: back-to-back periods on different floors of the same building (A-408 -> A-114) are penalised like a building change.
v58 = v57 + PREFER_ONE_SUBJECT_PER_DAY: soft companion of MAX_SAME_SUBJECT_PER_DAY (keep that one hard at 2, penalise every theory period of a subject beyond 1 per section per day).
v57 = v56 + saves during the run: first timetable found -> <name>_first and <name>; then every SAVE_EVERY seconds if improved -> <name>. Ctrl+C stops the search and saves the best timetable found.
v56 = v55 + two soft rules: NO_SECTION_ROOM_SWITCH (floating section changes theory room between back-to-back periods) and FACULTY_BACK_TO_BACK (teacher teaches two back-to-back periods, existing load included; one continuous lab block is not counted). P5/P6 idle weight for Mech/Civil S6 = 10000.
v55 = v54 solver, display only: sections taught together share one cell (AIML III-A/B/C), no course codes in elective tags.
v54 simplified: fixed faculty assignments, no staffing pools or selection.

AIML has one fixed PSE faculty per section, round robin from the owner's list.
Other faculty teams and all constraint weights remain as supplied.

Usage (PowerShell or terminal):
    python solve_timetable.py                    # solve -> timetable_output.xlsx + .html
    python solve_timetable.py --excel            # solve -> Excel only
    python solve_timetable.py --html             # solve -> HTML only
    python solve_timetable.py --html-from-excel  # no solve: HTML from timetable_output.xlsx (or give a path)
    Other options: --input workload_all_departments.xlsx, --seconds, --out, --name, --hint, --check-only

Input workbook sheets (see its README sheet):
    Settings        Days, Period times, Title, Scope, Basis, Seconds, Workers
    Activities      one row per class (lab batch, section or merged AI group)
    Faculty busy    fixed commitments (first-year classes) that block faculty
    Faculty limits  only confirmed daily/weekly limits
    Not scheduled   record-only rows (MOOCs, unresolved alternatives) - not timetabled
    Clarifications  record-only questions - copied to the output
    Assumptions     record-only notes - copied to the output

Rules enforced: no faculty, student-group or room double booking; a whole section
clashes with each of its batches (SEC:B1, SEC:B2) but batches of one section may run
together; multi-period sessions never cross a break; faculty busy times are blocked;
rows sharing a Parallel group run at the same time; rows sharing an Elective slot
(pooled sections split across several subjects) all run at the same time and block
every pooled section as a whole for that time; confirmed faculty limits.
The previous output (if found) is used as a warm start only; every class may move.
Requires: pip install ortools openpyxl
"""
import argparse,json,re,html,sys,datetime
from pathlib import Path
from collections import defaultdict,Counter
from openpyxl import Workbook,load_workbook
from openpyxl.styles import Alignment,Font,PatternFill,Border,Side
from openpyxl.utils import get_column_letter
from ortools.sat.python import cp_model

HERE=Path(__file__).resolve().parent
ACT_HEADERS=['Activity ID','Department','Semester','Course code','Course','Component','Student groups','Faculty IDs','Faculty names',
    'Periods/week','Sessions/week','Periods/session','Eligible rooms','Parallel group','Elective slot','Source','Assumptions']
INPUT_HEADERS=ACT_HEADERS
DAYS_SHORT={'Monday':'Mon','Tuesday':'Tue','Wednesday':'Wed','Thursday':'Thu','Friday':'Fri','Saturday':'Sat'}
class InputError(ValueError):pass

# ----------------------------------------------------------------- 1. read input
def split(v):return [x.strip() for x in re.split(r'[;,]',str(v or '')) if x.strip()]
def minutes(v):
    m=re.fullmatch(r'(\d{1,2}):(\d{2})',str(v).strip())
    if not m:raise InputError(f'Invalid time {v!r}; use HH:MM')
    return 60*int(m[1])+int(m[2])
def sheet_rows(wb,name,required=True):
    if name not in wb.sheetnames:
        if required:raise InputError(f'Missing sheet {name!r}')
        return []
    rows=list(wb[name].iter_rows(values_only=True))
    hi=next((i for i,r in enumerate(rows) if r and any(c is not None for c in r)),None)
    if hi is None:return []
    head=[str(c).strip() if c is not None else '' for c in rows[hi]]
    return [dict(zip(head,r)) for r in rows[hi+1:] if r and any(c not in (None,'') for c in r)]
def as_int(v,field,rid):
    try:
        f=float(v)
        if f>0 and f.is_integer():return int(f)
    except (TypeError,ValueError):pass
    raise InputError(f'{rid}: {field} must be a positive whole number (got {v!r})')

# ----------------------------------------------------------------- 1b. coordinator workbook -> classes
# The workload workbook (from the intake page) is the only input. This expands its
# course grid into the classes the solver places: one per section (or merged group),
# two per batch lab (B1, B2, synchronised with a swap partner), one per elective subject.
SEC_SHEET=re.compile(r'^(.+?)\s+S(\d+)$')
def _num(v):
    try:return float(v)
    except (TypeError,ValueError):return 0
def _cell(v):
    t=str(v if v is not None else '').strip()
    if t.endswith('.0') and t[:-2].isdigit():t=t[:-2]
    return t
def _secs(v):return [x for x in re.split(r'[+,;/ ]+',_cell(v).upper()) if x]
def _type(v,name,code):
    t=_cell(v).lower()
    if t:return 'L' if re.match(r'^(lab|l$|p$|practical)',t) else 'IP' if re.match(r'^(ip|pse|project)',t) else 'T'
    return 'L' if re.search(r'\blab(oratory)?\b',name or '',re.I) else 'IP' if _is_ip(code,name) else 'T'
def _is_ip(code,name):return bool(re.match(r'^(EVA2201|ESE3201)$',code or '',re.I) or re.search(r'integrated project|problem solving|\bPSE\b|\bIP[- ]?I',name or '',re.I))
def _default_block(p):return p if p<=3 else 3 if p%3==0 else 2 if p%2==0 else 1
def _slug(v):return re.sub(r'[^A-Z0-9]+','',str(v).upper())

WB_INFO={'facdept':{},'roompools':{}}
def expand_workbook(wb,errors):
    rows=lambda n:sheet_rows(wb,n,False)
    depts=[]; sections={}; home={}; rooms={}; names={}
    for r in rows('Sections'):
        d=_cell(r.get('Department')); s=int(_num(r.get('Semester')))
        if not d or not s:continue
        if d not in depts:depts.append(d)
        sections[d,s]=_secs(r.get('Sections'))
        for part in re.split(r'[;\n]+',_cell(r.get('Home rooms'))):
            if '=' in part:k,v=part.split('=',1);home[d,s,k.strip().upper()]=split(v)
    for r in rows('Rooms'):
        rooms[_cell(r.get('Department'))]=(split(r.get('Classrooms')),split(r.get('Labs')))
    for r in rows('Faculty'):
        fid=_cell(r.get('Employee ID'))
        if fid:names.setdefault(fid,_cell(r.get('Name')));WB_INFO['facdept'].setdefault(fid,_cell(r.get('Department')))
    for d in depts:
        cr=list(dict.fromkeys(rooms.get(d,([],[]))[0]+[x for (dd,_,_),v in home.items() if dd==d for x in v]))
        WB_INFO['roompools'][d]={'classrooms':cr,'labs':rooms.get(d,([],[]))[1]}
    cat={}
    for r in rows('Curriculum'):
        cat[_cell(r.get('Department')),int(_num(r.get('Semester'))),_cell(r.get('Code')).upper()]=(int(_num(r.get('Theory periods'))),int(_num(r.get('Lab periods'))))
    merged=defaultdict(list)
    for r in rows('Merged classes'):
        d,s,code=_cell(r.get('Department')),int(_num(r.get('Semester'))),_cell(r.get('Code')).upper();g=sorted(_secs(r.get('Sections')))
        if d and code and len(g)>1:merged[d,s,code,'T' if _type(r.get('Type'),'',code)=='T' else 'L'].append(g)
    # Lab batches: a lab listed alone runs as two batches at separate times. A lab listed with
    # one or more partners ("Swaps with": A; B) rotates: with n labs in the group the section
    # splits into n batches, and in each rotation slot every batch is in a different lab.
    batches={}
    for r in rows('Lab batches'):
        d,s,a=_cell(r.get('Department')),int(_num(r.get('Semester'))),_cell(r.get('Lab code')).upper()
        if not (d and a):continue
        grp=sorted({a,*[x.upper() for x in split(r.get('Swaps with'))]})
        for code in grp:
            if len(grp)>1 or (d,s,code) not in batches:batches[d,s,code]=grp
    acts=[];used=set()
    def uid(b):
        k,i=b,2
        while k in used:k=f'{b}-{i}';i+=1
        used.add(k);return k
    def add(d,s,code,name,kind,groups,fac,hours,dur,rms,par='',es='',note='',src=''):
        aid=uid(f'{d}-S{s}-{code}-{ {"Theory":"T","Lab":"L","IP":"IP"}[kind]}-{"".join(g.split(":")[0].rsplit("-",1)[1] for g in groups)}'+('-'+groups[0].split(':')[1] if ':' in groups[0] and not es else ''))
        if hours%dur:errors.append(f'{d} S{s} {code}: {hours} periods not divisible into {dur}-period sessions');return
        f=fac or ['TBC-FAC-'+aid]
        acts.append({'id':aid,'dept':d,'semester':s,'code':code,'name':name,'kind':kind,'groups':groups,'faculty':f,
            'faculty_names':'; '.join(names.get(x,x) for x in fac) if fac else 'TBC faculty','hours':hours,'sessions':hours//dur,'duration':dur,
            'rooms':(list(dict.fromkeys(rms))+['TBC-ROOM-'+aid]) if SOFT_ROOM_WEIGHT else (list(dict.fromkeys(rms)) or ['TBC-ROOM-'+aid]),'parallel':par,'eslot':es,'source':src,'assumptions':note,
            'course_id':f'{d}-S{s}-{code}-{kind}'})
    for sheet in wb.sheetnames:
        m=SEC_SHEET.match(sheet)
        if not m or m[1] not in depts:continue
        d,s=m[1],int(m[2]);secs=sections.get((d,s),[]);cls,labs=rooms.get(d,([],[]))
        root=lambda x:f'{x.split(".")[0]}-S{s}-{x.split(".")[1]}' if '.' in x else f'{d}-S{s}-{x}'  # "ECM.A" = section A of ECM (cross-department merge)
        for i,r in enumerate(rows(sheet),2):
            code=_cell(r.get('Code')).upper();name=_cell(r.get('Course name'))
            if not code and not name:continue
            if not code:errors.append(f'{sheet} row {i}: course code missing');continue
            t=_type(r.get('Type'),name,code);lab=t!='T'
            p=int(_num(r.get('Periods')))
            if not p and not _cell(r.get('Periods')):
                c=cat.get((d,s,code)) or cat.get((d,s,code.rstrip('P')))
                p=(c[1] if lab else c[0]) if c else 0
            if not p:errors.append(f'{sheet} row {i} ({code}): periods missing');continue
            blk=(int(_num(r.get('Lab block'))) or _default_block(p)) if lab else 1
            groups=merged.get((d,s,code,'L' if lab else 'T'),[]);ing={x for g in groups for x in g}
            allocs=[]
            for g in groups:
                lead=next((x for x in g if _cell(r.get(x))),None)
                if lead:allocs.append((g,_cell(r.get(lead))))  # merged group with no faculty on this row: not on this row
            for x in secs:
                if x not in ing and _cell(r.get(x)):allocs.append(([x],_cell(r.get(x))))
            rr=split(r.get('Rooms'));note=_cell(r.get('Note'));src=f'{sheet} row {i}'
            for g,c in allocs:
                fac=[] if c.upper()=='TBC' else split(c);grp=[root(x) for x in g]
                if not lab:add(d,s,code,name,'Theory',grp,fac,p,1,rr or [y for x in g for y in home.get((d,s,x),[])] or cls,par=_cell(r.get('Parallel group')),note=note,src=src);continue
                kind='IP' if t=='IP' else 'Lab';lr=rr or ([] if kind=='IP' else labs)
                if kind=='Lab' and (d,s,code) in batches:
                    k=batches[d,s,code];nb=max(2,len(k))
                    # "a; b | c" in a section cell = different teachers per batch (batch 1 | batch 2 | ...).
                    per=[[] if x.strip().upper()=='TBC' else split(x) for x in c.split('|')] if '|' in c else None
                    for b in range(1,nb+1):
                        par=''
                        if len(k)>1 and len(g)==1:par=f'{grp[0]}|{"~".join(k)}|{(b-1+k.index(code))%nb+1}'
                        add(d,s,code,name,kind,[x+f':B{b}' for x in grp],per[min(b,len(per))-1] if per else fac,p,blk,lr,par=par,note=note,src=src)
                else:add(d,s,code,name,kind,grp,fac,p,blk,lr,par=_cell(r.get('Parallel group')),note=note,src=src)
    pools=defaultdict(list)
    for i,r in enumerate(rows('Electives'),2):
        d=_cell(r.get('Department'));s=int(_num(r.get('Semester')));slot=_cell(r.get('Slot'))
        if d and slot:pools[d,s,slot,tuple(sorted(_secs(r.get('Pooled sections'))))].append((i,r))
    for (d,s,slot,secs),items in pools.items():
        sl=_slug(slot) or 'SLOT';es=f'{d}-S{s}-{sl}-{"".join(secs)}';cls=rooms.get(d,([],[]))[0]
        per=int(_num(items[0][1].get('Periods/week')))
        if not per:errors.append(f'Electives {d} S{s} {slot}: periods/week missing');continue
        for j,(i,r) in enumerate(items):
            code=_cell(r.get('Subject code')).upper() or f'SUB{j+1}'
            add(d,s,code,_cell(r.get('Subject name')),'Theory',[f'{d}-S{s}-{x}:{sl}-{_slug(code)}' for x in secs],split(r.get('Faculty IDs')),per,1,
                split(r.get('Room')) or [y for x in secs for y in home.get((d,s,x),[])]+cls,es=es,note=_cell(r.get('Note')),src=f'Electives row {i}')
    clar=[[d,'',rooms_note] for d,rooms_note in ((_cell(r.get('Department')),_cell(r.get('Notes'))) for r in rows('Rooms')) if rooms_note]
    return acts,clar

def read_input(path):
    wb=load_workbook(path,data_only=True);errors=[]
    st={str(r['Setting']).strip():r['Value'] for r in sheet_rows(wb,'Settings') if r.get('Setting')}
    days=split(st.get('Days'));times=[tuple(minutes(x) for x in p.split('-')) for p in split(st.get('Period times'))]
    if not days or not times:raise InputError('Settings: Days and Period times are required')
    for i,(a,b) in enumerate(times):
        if a>=b or (i and a<times[i-1][1]):raise InputError('Settings: period times must be increasing and non-overlapping')
    acts=[];seen=set();wclar=[]
    if 'Activities' not in wb.sheetnames:
        if 'Sections' not in wb.sheetnames:raise InputError('Not a timetable workbook: it needs the Sections sheet and course sheets such as "AI S6" (or an Activities sheet).')
        acts,wclar=expand_workbook(wb,errors)
    for r in (sheet_rows(wb,'Activities') if 'Activities' in wb.sheetnames else []):
        rid=str(r.get('Activity ID') or '').strip()
        if not rid:errors.append('Activities: a row has no Activity ID');continue
        if rid in seen:errors.append(f'{rid}: duplicate Activity ID');continue
        seen.add(rid)
        try:
            a={'id':rid,'dept':str(r.get('Department') or '').strip(),'semester':as_int(r.get('Semester'),'Semester',rid),
               'code':str(r.get('Course code') or '').strip(),'name':str(r.get('Course') or '').strip(),'kind':str(r.get('Component') or 'Theory').strip(),
               'groups':split(r.get('Student groups')),'faculty':split(r.get('Faculty IDs')),'faculty_names':str(r.get('Faculty names') or ''),
               'hours':as_int(r.get('Periods/week'),'Periods/week',rid),'sessions':as_int(r.get('Sessions/week'),'Sessions/week',rid),
               'duration':as_int(r.get('Periods/session'),'Periods/session',rid),'rooms':split(r.get('Eligible rooms')),
               'parallel':str(r.get('Parallel group') or '').strip(),'eslot':str(r.get('Elective slot') or '').strip(),'source':str(r.get('Source') or ''),'assumptions':str(r.get('Assumptions') or '')}
        except InputError as e:errors.append(str(e));continue
        for k,label in [('groups','Student groups'),('faculty','Faculty IDs'),('rooms','Eligible rooms')]:
            if not a[k]:errors.append(f'{rid}: {label} is empty (use a TBC placeholder)')
        if a['hours']!=a['sessions']*a['duration']:errors.append(f'{rid}: Periods/week {a["hours"]} != Sessions/week x Periods/session')
        if a['duration']>len(times):errors.append(f'{rid}: session longer than the day')
        a['course_id']=f"{a['dept']}-S{a['semester']}-{a['code']}-{a['kind']}"
        acts.append(a)
    if wclar is not None and 'Activities' not in wb.sheetnames:
        for a in acts:
            if a['duration']>len(times):errors.append(f"{a['dept']} S{a['semester']} {a['code']}: a {a['duration']}-period session is longer than the day")
    par=defaultdict(list)
    for a in acts:
        if a['parallel']:par[a['parallel']].append(a)
    for g,m in par.items():
        if len(m)<2 or len({(a['sessions'],a['duration']) for a in m})!=1:errors.append(f'Parallel group {g}: needs 2+ rows with identical sessions and periods/session')
    es=defaultdict(list)
    for a in acts:
        if a['eslot']:es[a['eslot']].append(a)
    for g,m in es.items():
        if len({(a['sessions'],a['duration']) for a in m})!=1:errors.append(f'Elective slot {g}: all subjects need the same Sessions/week and Periods/session')
        if any(a['parallel'] for a in m):errors.append(f'Elective slot {g}: do not also set a Parallel group on its rows')
    busy=[]
    dayof={d[:3].lower():d for d in days}
    for r in sheet_rows(wb,'Faculty busy',False):
        r['Day']=dayof.get(str(r.get('Day') or '').strip()[:3].lower())
        if not r['Day']:errors.append(f'Faculty busy: unknown day in row {r}');continue
        tm=lambda v:minutes(v.strftime('%H:%M') if hasattr(v,'strftime') else v)
        component = str(r.get('Component') or 'Unknown').strip()
        exception = r.get('Daily total exception')
        if exception not in (None, ''):
            exception = as_int(exception, 'Daily total exception', 'Faculty busy')
            if component.lower() not in PROJECT_COMPONENTS or exception != 6:
                raise InputError('Faculty busy: Daily total exception is supported only for project work (Project, JOE, DOE) with value 6')
        else:
            exception = None
        busy.append({'faculty':split(r.get('Faculty IDs')),'day':r['Day'],
                     'start':tm(r['Start']),'end':tm(r['End']),
                     'component':component,'daily_total_exception':exception,'note':str(r.get('Note') or '').strip(),'room':str(r.get('Room') or '').strip()})
    limits={}
    for r in sheet_rows(wb,'Faculty limits',False):
        faculty_id = _cell(r.get('Faculty ID'))
        if not faculty_id or faculty_id in limits:
            raise InputError('Faculty limits: missing or duplicate Faculty ID')
        def opt(value):
            return as_int(value, 'limit', faculty_id) if value not in (None, '') else None
        limits[faculty_id] = {'daily': opt(r.get('Daily limit')), 'weekly': opt(r.get('Weekly limit'))}
    rec=lambda n,keys:[{k:(r.get(h) if r.get(h) is not None else '') for k,h in keys} for r in sheet_rows(wb,n,False)]
    not_sched=rec('Not scheduled',[('status','Status'),('dept','Department'),('semester','Semester'),('id','Ref'),('code','Course code'),('course','Course'),('groups','Groups'),('faculty','Faculty'),('periods','Periods'),('reason','Reason')])
    clar=[[r.get('Department',''),r.get('Year',''),r.get('Question','')] for r in sheet_rows(wb,'Clarifications',False)]+wclar
    notes=[[r.get('Activity',''),r.get('Field',''),r.get('Value',''),r.get('Reason','')] for r in sheet_rows(wb,'Assumptions',False)]
    section_busy = []
    for row in sheet_rows(wb, 'Section busy', False):
        day = dayof.get(str(row.get('Day') or '').strip()[:3].lower())
        start = minutes(str(row.get('Start'))[:5])
        end = minutes(str(row.get('End'))[:5])
        if day is None or start >= end:
            raise InputError('Section busy: invalid day or time interval')
        section = f"{_cell(row.get('Department'))}-S{int(_num(row.get('Semester')))}-{_cell(row.get('Section')).upper()}"
        section_busy.append(dict(section=section, day=day, start=start, end=end, label=_cell(row.get('Label')) or 'Fixed commitment'))
    constraints = read_constraints(wb, acts, busy, len(times))
    # Home-room preference comes from the input, not department-specific code.
    home_rooms = {}
    for row in sheet_rows(wb, 'Sections', False):
        department = _cell(row.get('Department'))
        semester = int(_num(row.get('Semester')))
        for part in re.split(r'[;\n]+', _cell(row.get('Home rooms'))):
            if '=' in part:
                section, room_list = part.split('=', 1)
                home_rooms[f'{department}-S{semester}-{section.strip().upper()}'] = split(room_list)
    for activity in acts:
        activity['home_rooms'] = sorted({room for group in activity['groups']
            for room in home_rooms.get(group.split(':', 1)[0], [])})
    fixed_sessions = []
    fixed_seen = set()
    by_id = {activity['id']: activity for activity in acts}
    day_aliases = {day[:3].lower(): day for day in days}
    for row in sheet_rows(wb, 'Fixed sessions', False):
        activity_id = _cell(row.get('Activity ID'))
        if activity_id not in by_id:
            raise InputError(f'Fixed sessions: unknown Activity ID {activity_id!r}')
        session = as_int(row.get('Session'), 'Session', activity_id)
        period = as_int(row.get('Period'), 'Period', activity_id)
        day = day_aliases.get(str(row.get('Day') or '').strip()[:3].lower())
        room = _cell(row.get('Room'))
        activity = by_id[activity_id]
        if (activity_id, session) in fixed_seen:
            raise InputError(f'Fixed sessions: duplicate {activity_id} session {session}')
        fixed_seen.add((activity_id, session))
        if not day or session > activity['sessions'] or period + activity['duration'] - 1 > len(times):
            raise InputError(f'Fixed sessions: invalid session/day/period for {activity_id}')
        if room and room not in activity['rooms']:
            raise InputError(f'Fixed sessions: ineligible room {room!r} for {activity_id}')
        if any(times[q][1] != times[q + 1][0] for q in range(period - 1, period + activity['duration'] - 2)):
            raise InputError(f'Fixed sessions: session crosses a break for {activity_id}')
        fixed_sessions.append(dict(activity_id=activity_id, session=session, day=day, period=period, room=room))
    if errors:raise InputError('Input errors:\n  '+'\n  '.join(errors[:60])+(f'\n  ... {len(errors)-60} more' if len(errors)>60 else ''))
    return {'days':days,'times':times,'activities':acts,'busy':busy,'faculty_limits':limits,'not_scheduled':not_sched,'clarifications':clar,'notes':notes,'settings':st,'facdept':WB_INFO['facdept'],'roompools':WB_INFO['roompools'],'constraints':constraints,'fixed_sessions':fixed_sessions,'section_busy':section_busy}

def read_hint(path):
    """Previous output workbook (Sessions sheet) -> warm-start hints."""
    try:
        return [{'activity_id':str(r['Activity ID']),'session':int(r['Session']),'day':r['Day'],'period':int(r['Period']),'room':r['Room'],'duration':int(r['Duration'])}
                for r in sheet_rows(load_workbook(path,read_only=True,data_only=True),'Sessions')]
    except Exception as e:
        print('No warm start used:',e);return []

def from_hints(data,hints):
    """Rebuild a result from a previous output's Sessions sheet (no solving)."""
    acts={a['id']:a for a in data['activities']};times=data['times'];got=defaultdict(set);sch=[]
    for h in hints:
        a=acts.get(h['activity_id'])
        if not a or h['session'] in got[a['id']]:continue
        got[a['id']].add(h['session']);p=h['period']
        sch.append({'activity_id':a['id'],'course_id':a.get('course_id'),'course':a['name'],'groups':a['groups'],'faculty':a['faculty'],'room':h['room'],
            'day':h['day'],'period':p,'duration':a['duration'],'start':times[p-1][0],'end':times[p+a['duration']-2][1],'session':h['session']})
    un=[{'activity_id':a['id'],'session':k,'periods':a['duration'],'reason':'Not placed in the last solver run'} for a in data['activities'] for k in range(1,a['sessions']+1) if k not in got[a['id']]]
    res={'status':'RENDERED','schedule':sch,'unplaced':un,'seconds':0,'required_periods':sum(a['hours'] for a in data['activities']),'placed_periods':sum(s['duration'] for s in sch)}
    res['unplaced_periods']=res['required_periods']-res['placed_periods'];return res

"""Constraint definitions and handlers to embed in the standalone solver."""

# Only P4 and P7 may be free (0-based periods 0,1,2,4,5 must be occupied unless a fixed block covers them).
MUST_FILL_PERIODS = (0, 1, 2, 4, 5)
# Sections with fewer periods than must-fill slots: P5/P6 become top-priority soft for them.
SOFT_FILL_PERIODS = (4, 5)
SOFT_FILL_SECTIONS = {'Mech-S6-A', 'Mech-S6-B', 'CIVIL-S6-A', 'CIVIL-S6-B'}
SOFT_FILL_WEIGHT = 10000
TBC_PERIODS = {4, 6, 7}  # 1-based periods allowed for courses that have no real teacher yet
SOFT_ROOM_WEIGHT = 0
TBCS_CODES = {'EAEXX07', 'EAEXX08'}
# PSE-II venues (OIA "PSE II Rooms" list): two sections at a time. Any other class
# using one of these rooms takes the whole room.
SHARED_ROOM_SECTIONS = {'FAB ZONE A': 2, 'FAB ZONE B': 2, 'FAB ZONE C': 2, 'I-306': 2, 'I-506': 2}
# Training venues reserved from Section busy labels such as "TT: DA-Python (I-306, T2)".
RESERVED_TRAINING_ROOMS = {'I-306', 'I-405', 'FAB ZONE C'}
def training_room(label):
    m = re.search(r'\(([^,()]+?)\s*(?:,|\))', str(label or ''))
    room = m.group(1).strip() if m else ''
    return room if room.upper() in RESERVED_TRAINING_ROOMS else None
def room_demand(activity, room):
    capacity = SHARED_ROOM_SECTIONS[str(room).strip().upper()]
    if activity['code'] != 'ESE3201':
        return capacity
    return min(capacity, len({g.split(':', 1)[0] for g in activity['groups']}))
def building(room):
    room = str(room or '').strip().upper()
    if not room or room.startswith('TBC'):
        return None
    m = re.match(r'^([A-Z])-', room)
    return m.group(1) if m else room
def floor_loc(room):
    """'A-408' -> 'A:4', 'C-G2' -> 'C:0', 'C-05' -> 'C:0'; rooms without a floor number (FAB zones, named labs) -> None."""
    room = str(room or '').strip().upper()
    m = re.match(r'^([A-Z])-(G|\d)', room)
    if not m:
        return None
    return m.group(1) + ':' + ('0' if m.group(2) == 'G' or len(re.match(r'^[A-Z]-(\w+)', room).group(1)) < 3 else m.group(2))
LUNCH_START = 12 * 60 + 40  # 12:40, end of P4
def is_placeholder(f):
    return str(f).startswith('TBC') or bool(re.match(r'^(DM|DGD|TP)\d', str(f)))
# name: allowed types, value kind, scopes, enabled, type, value, weight, description
CONSTRAINT_SPECS = {
    'P1_NO_IDLE': ('HARD', None, 'Global Department Section', False, 'HARD', None, 1, 'Require P1 occupancy for every student group; recorded section commitments are exempt.'),
    'COURSE_ALLOWED_SLOTS': ('HARD', 'slots', 'Course', False, 'HARD', None, 1, 'This course may only be placed in the listed Day:Pn slots.'),
    'FACULTY_ALLOWED_SLOTS': ('HARD', 'slots', 'Faculty', False, 'HARD', None, 1, 'Teaching allowed only in listed Day:Pn slots.'),
    'FACULTY_ALLOWED_DAYS': ('HARD', 'days', 'Faculty', False, 'HARD', None, 1, 'Restrict generated teaching to the listed working days.'),
    'NO_SECTION_OVERLAP': ('HARD', None, 'Global', True, 'HARD', None, 1, 'Prevent section and batch clashes.'),
    'NO_FACULTY_CLASH': ('HARD', None, 'Global', True, 'HARD', None, 1, 'Prevent simultaneous faculty assignments.'),
    'NO_ROOM_CLASH': ('HARD', None, 'Global', True, 'HARD', None, 1, 'Prevent simultaneous use of a physical room.'),
    'EXISTING_SCHEDULE_CLASH': ('HARD', None, 'Global', True, 'HARD', None, 1, 'Respect Faculty busy intervals.'),
    'FIXED_SLOT': ('HARD', None, 'Global', True, 'HARD', None, 1, 'Honor optional Fixed sessions rows: Activity ID, Session, Day, Period, Room.'),
    'PERIODS_PER_WEEK': ('HARD', None, 'Global', False, 'HARD', None, 1, 'Enable for complete scheduling. Disabled permits explicitly reported unplaced sessions.'),
    'MAX_FACULTY_THEORY_PER_DAY': ('HARD SOFT', 'integer', 'Global Faculty', True, 'HARD', 4, 10, 'Daily theory cap. Existing commitments with unknown component count as theory conservatively.'),
    'MAX_FACULTY_CONSECUTIVE': ('HARD SOFT', 'integer', 'Global Faculty', True, 'SOFT', 3, 1000, 'Teacher teaches at most this many periods in a row (lunch breaks a run); labs, theory and existing load all count.'),
    'MAX_FACULTY_TOTAL_PER_DAY': ('HARD SOFT', 'integer', 'Global Faculty', True, 'HARD', 5, 10, 'Daily total cap. Every lab/IP/PSE period counts. Includes existing commitments.'),
    'MAX_FACULTY_PERIODS_PER_WEEK': ('HARD SOFT', 'integer', 'Global Faculty', False, 'HARD', None, 10, 'Optional confirmed weekly cap. Enter Value before enabling.'),
    'MAX_SAME_SUBJECT_PER_DAY': ('HARD SOFT', 'integer', 'Global Department Section Course', True, 'SOFT', 1, 8, 'Theory periods per section/course/day.'),
    'NO_CONSECUTIVE_THEORY_SAME_SUBJECT': ('HARD SOFT', None, 'Global Department Section Course', True, 'SOFT', None, 6, 'Avoid adjacent theory of the same subject within a teaching half-day.'),
    'COMPACT_SCHEDULE': ('HARD SOFT', None, 'Global Department Section Course', True, 'SOFT', None, 8, 'Avoid student gaps between first and last class of the day.'),
    'FACULTY_COMPACT_SCHEDULE': ('HARD SOFT', None, 'Global Faculty', True, 'SOFT', None, 3, 'Avoid faculty gaps, including existing commitments.'),
    'BALANCE_FACULTY_DAYS': ('SOFT', None, 'Global Faculty', True, 'SOFT', None, 3, 'Minimize the difference between highest and lowest daily faculty load.'),
    'PREFER_HOME_ROOM': ('HARD SOFT', None, 'Global Department Section Course', True, 'SOFT', None, 2, 'Prefer declared section home rooms for theory where available.'),
    'MINIMIZE_DISTINCT_ROOMS': ('SOFT', None, 'Global Department Section Course', True, 'SOFT', None, 2, 'Prefer fewer distinct theory rooms per section during the week.'),
    'THEORY_TIME_OF_DAY_BALANCE': ('SOFT', None, 'Global Department Section Course', True, 'SOFT', None, 2, 'Balance theory between morning and afternoon in proportion to available periods.'),
    'MAX_SAME_THEORY_SUBJECT_IN_P1': ('HARD SOFT', 'integer', 'Global Department Section Course', True, 'SOFT', 1, 2, 'Limit repeated first-period assignment for the same theory subject.'),
    'AVOID_7TH_HOUR_THEORY': ('HARD SOFT', 'integer', 'Global Department Section Course', False, 'SOFT', 7, 1, 'Optional theory avoidance in the period numbered by Value.'),
    'NO_FACULTY_BLOCK_HOP': ('HARD SOFT', None, 'Global Department Course', True, 'SOFT', None, 500, 'A teacher should not change building between back-to-back periods (lunch is a break).'),
    'NO_SECTION_BLOCK_HOP': ('HARD SOFT', None, 'Global Department Section Course', True, 'SOFT', None, 500, 'A section/batch should not change building between back-to-back periods (lunch is a break).'),
    'PREFER_ONE_SUBJECT_PER_DAY': ('SOFT', 'integer', 'Global Department Section Course', True, 'SOFT', 1, 5000, 'Soft: theory periods per section/course/day above this value are penalised (each extra period).'),
    'NO_FACULTY_FLOOR_HOP': ('HARD SOFT', None, 'Global Department Faculty', True, 'SOFT', None, 5000, 'A teacher should not teach back-to-back periods on different floors of the same building (lunch is a break).'),
    'NO_SECTION_FLOOR_HOP': ('HARD SOFT', None, 'Global Department Section', True, 'SOFT', None, 5000, 'A section should not have back-to-back periods on different floors of the same building (lunch is a break).'),
    'CONSECUTIVE_DAYS_PER_COURSE': ('HARD', None, 'Course', False, 'HARD', None, 1, "The days used by this course's labs must be consecutive."),
    'MAX_DAYS_PER_COURSE': ('HARD', 'integer', 'Course', False, 'HARD', 2, 1, 'All lab sessions of this course (all sections and batches) fall on at most this many days.'),
    'FACULTY_SLOTS_LAB_EXEMPT': ('HARD', None, 'Faculty', False, 'HARD', None, 1, "This teacher's FACULTY_ALLOWED_SLOTS rule does not apply to labs."),
    'SAME_LAB_ROOM': ('HARD SOFT', None, 'Global Department Section', True, 'SOFT', None, 2500, 'All sessions of a lab course for one section (every batch) use the same room; each room mismatch is penalised.'),
    'NO_FULL_DAY_LABS': ('HARD SOFT', None, 'Global Department Section', True, 'SOFT', None, 5000, 'A section/batch should not have labs in both the morning and the afternoon of the same day.'),
    'MAX_LAB_PERIODS_PER_DAY': ('HARD SOFT', 'integer', 'Global Department Section', True, 'SOFT', 4, 5000, 'Lab periods per section/batch per day above this value are penalised, each extra period.'),
    'MAX_LABS_PER_DAY': ('HARD SOFT', 'integer', 'Global Department Section', True, 'HARD', 1, 5000, 'At most this many lab blocks per section/batch per day (skills courses listed in LAB_DAY_LIMIT_EXCLUDE are not counted).'),
    'LAB_DAY_LIMIT_EXCLUDE': ('HARD', None, 'Course', False, 'HARD', None, 1, 'This course is not counted by MAX_LABS_PER_DAY.'),
    'MAX_LABS_PER_MORNING': ('HARD SOFT', 'integer', 'Global Department Section', True, 'SOFT', 1, 5000, 'At most this many lab blocks per section/batch in the morning (P1-P4) of a day.'),
    'NO_SECTION_ROOM_SWITCH': ('HARD SOFT', None, 'Global Department Section', True, 'SOFT', None, 5000, 'A section/batch should not change theory room between back-to-back periods (lunch is a break).'),
    'FACULTY_BACK_TO_BACK': ('SOFT', None, 'Global Department Faculty', True, 'SOFT', None, 100, 'Avoid a teacher teaching two back-to-back periods (existing load counts; one continuous lab block does not).'),
    'EXISTING_LOAD_GAP': ('HARD SOFT', None, 'Global Department Course', True, 'SOFT', None, 20, 'Avoid a new class directly next to a teacher existing-load period (same half-day).'),
    'NO_BACKTOBACK_LABS': ('HARD SOFT', None, 'Global Department Section Course', True, 'SOFT', None, 1, 'Avoid adjacent separate lab blocks for the same student batch.'),
}


def default_constraints():
    return [dict(name=name, enabled=spec[3], type=spec[4], value=spec[5],
                 weight=spec[6], scope='Global', target='')
            for name, spec in CONSTRAINT_SPECS.items() if name not in ('FACULTY_ALLOWED_DAYS', 'FACULTY_ALLOWED_SLOTS', 'P1_NO_IDLE', 'NO_SECTION_ROOM_SWITCH', 'FACULTY_BACK_TO_BACK', 'PREFER_ONE_SUBJECT_PER_DAY', 'NO_FACULTY_FLOOR_HOP', 'NO_SECTION_FLOOR_HOP', 'MAX_LABS_PER_MORNING', 'COURSE_ALLOWED_SLOTS', 'NO_FULL_DAY_LABS', 'MAX_DAYS_PER_COURSE', 'FACULTY_SLOTS_LAB_EXEMPT', 'SAME_LAB_ROOM', 'CONSECUTIVE_DAYS_PER_COURSE', 'MAX_LABS_PER_DAY', 'LAB_DAY_LIMIT_EXCLUDE', 'MAX_LAB_PERIODS_PER_DAY')]


def read_constraints(wb, activities, busy, periods):
    if 'Constraints' not in wb.sheetnames:
        return default_constraints()
    parsed = []
    seen = set()
    faculty = {f for a in activities for f in a['faculty']}
    faculty.update(f for b in busy for f in b['faculty'])
    targets = {
        'Department': {a['dept'] for a in activities},
        'Faculty': faculty,
        'Section': {g.split(':', 1)[0] for a in activities for g in a['groups']},
        'Course': {a['code'] for a in activities},
    }
    for row_number, row in enumerate(sheet_rows(wb, 'Constraints'), 2):
        name = str(row.get('Name') or '').strip().upper()
        label = f'Constraints row {row_number} ({name})'
        if name not in CONSTRAINT_SPECS:
            raise InputError(f'{label}: unknown constraint name')
        spec = CONSTRAINT_SPECS[name]
        enabled = str(row.get('Enabled') or '').strip().upper()
        if enabled not in ('YES', 'NO'):
            raise InputError(f'{label}: Enabled must be YES or NO')
        kind = str(row.get('Type') or '').strip().upper()
        if kind not in spec[0].split():
            raise InputError(f'{label}: supported Type is {spec[0]}')
        scope = str(row.get('Scope') or 'Global').strip().title()
        target = _cell(row.get('Target'))
        if scope not in spec[2].split():
            raise InputError(f'{label}: supported scopes are {spec[2]}')
        if scope == 'Global' and target:
            raise InputError(f'{label}: Global scope requires an empty Target')
        if scope != 'Global' and target not in targets[scope]:
            raise InputError(f'{label}: unknown {scope} target {target!r}')
        key = (name, scope, target)
        if key in seen:
            raise InputError(f'{label}: duplicate name/scope/target')
        seen.add(key)
        value = row.get('Value')
        if spec[1] == 'integer':
            if value in (None, ''):
                if enabled == 'YES':
                    raise InputError(f'{label}: Value is required when enabled')
                value = None
            else:
                value = as_int(value, 'Value', label)
                if name == 'AVOID_7TH_HOUR_THEORY' and value > periods:
                    raise InputError(f'{label}: period number exceeds periods/day')
        elif spec[1] == 'slots':
            settings = {r['Setting']: r['Value'] for r in sheet_rows(wb, 'Settings')}
            day_names = split(settings.get('Days'))
            lookup = {d.lower(): d for d in day_names}
            lookup.update({d[:3].lower(): d for d in day_names})
            value_slots = []
            for token in split(value):
                slot = re.fullmatch(r'([^:]+):P(\d+)', token, re.I)
                if not slot or slot[1].strip().lower() not in lookup or not 1 <= int(slot[2]) <= periods:
                    raise InputError(f'{label}: use configured Day:Pn slots, e.g. Monday:P3')
                value_slots.append((lookup[slot[1].strip().lower()], int(slot[2])))
            if enabled == 'YES' and not value_slots:
                raise InputError(f'{label}: at least one allowed slot is required')
            value = list(dict.fromkeys(value_slots))
        elif spec[1] == 'days':
            settings = {r['Setting']: r['Value'] for r in sheet_rows(wb, 'Settings')}
            day_names = split(settings.get('Days'))
            lookup = {d.lower(): d for d in day_names}
            lookup.update({d[:3].lower(): d for d in day_names})
            raw_days = split(value)
            if enabled == 'YES' and not raw_days:
                raise InputError(f'{label}: allowed days are required')
            if any(d.lower() not in lookup for d in raw_days):
                raise InputError(f'{label}: allowed days must be configured working days')
            value = list(dict.fromkeys(lookup[d.lower()] for d in raw_days))
        elif value not in (None, ''):
            raise InputError(f'{label}: this constraint does not accept a Value')
        weight = as_int(row.get('Weight'), 'Weight', label)
        parsed.append(dict(name=name, enabled=enabled == 'YES', type=kind,
                           value=value, weight=weight, scope=scope, target=target))
    missing = (set(CONSTRAINT_SPECS) - {'FACULTY_ALLOWED_DAYS', 'FACULTY_ALLOWED_SLOTS', 'P1_NO_IDLE', 'NO_SECTION_ROOM_SWITCH', 'FACULTY_BACK_TO_BACK', 'PREFER_ONE_SUBJECT_PER_DAY', 'NO_FACULTY_FLOOR_HOP', 'NO_SECTION_FLOOR_HOP', 'MAX_LABS_PER_MORNING', 'COURSE_ALLOWED_SLOTS', 'NO_FULL_DAY_LABS', 'MAX_DAYS_PER_COURSE', 'FACULTY_SLOTS_LAB_EXEMPT', 'SAME_LAB_ROOM', 'CONSECUTIVE_DAYS_PER_COURSE', 'MAX_LABS_PER_DAY', 'LAB_DAY_LIMIT_EXCLUDE', 'MAX_LAB_PERIODS_PER_DAY'}) - {r['name'] for r in parsed}
    if missing:
        raise InputError('Constraints: missing rows (use Enabled=NO to disable): ' + ', '.join(sorted(missing)))
    return parsed


def rule_enabled(data, name):
    return any(r['name'] == name and r['enabled']
               for r in data.get('constraints', default_constraints()))


PROJECT_COMPONENTS = ('project', 'joe', 'joe lab', 'doe', 'doe lab')
PROJECT_DAY_CAP = 6

def existing_daily_exceptions(data):
    """A day with existing project work (Project, JOE, DOE) may carry 6 periods in total."""
    exceptions = {}
    project_periods = defaultdict(set)
    for busy in data['busy']:
        is_project = str(busy.get('component') or '').lower() in PROJECT_COMPONENTS
        limit = busy.get('daily_total_exception')
        if limit is not None and (not is_project or limit != PROJECT_DAY_CAP):
            raise InputError('Only existing project work (Project, JOE, DOE) may have a daily total exception, and it must be 6')
        if is_project:
            for period, (start, end) in enumerate(data['times']):
                if overlaps(start, end, busy['start'], busy['end']):
                    for faculty in busy['faculty']:
                        project_periods[faculty, busy['day']].add(period)
                        exceptions[faculty, busy['day']] = PROJECT_DAY_CAP
    return exceptions


def existing_cap_warnings(data):
    """Explain fixed loads that already exceed a configured hard limit."""
    rules = data.get('constraints', default_constraints())
    overrides = {(r['name'], r['target']) for r in rules if r['scope'] == 'Faculty'}
    occupied = defaultdict(set)
    theory = defaultdict(set)
    for busy in data['busy']:
        for period, (start, end) in enumerate(data['times']):
            if overlaps(start, end, busy['start'], busy['end']):
                for faculty in busy['faculty']:
                    slot = (busy['day'], period)
                    occupied[faculty].add(slot)
                    if str(busy.get('component') or 'Unknown').lower() not in ('lab', 'practical', 'ip', 'pse', 'non-teaching') + PROJECT_COMPONENTS:
                        theory[faculty].add(slot)
    exceptions = existing_daily_exceptions(data)
    warnings = []
    for rule in rules:
        if not rule['enabled'] or rule['type'] != 'HARD' or not rule['name'].startswith('MAX_FACULTY_'):
            continue
        faculty_ids = [rule['target']] if rule['scope'] == 'Faculty' else sorted(occupied)
        for faculty in faculty_ids:
            if rule['scope'] == 'Global' and (rule['name'], faculty) in overrides:
                continue
            slots = theory[faculty] if rule['name'] == 'MAX_FACULTY_THEORY_PER_DAY' else occupied[faculty]
            loads = Counter(day for day, period in slots)
            if rule['name'] == 'MAX_FACULTY_PERIODS_PER_WEEK':
                loads = {'week': len(slots)}
            for day, value in loads.items():
                limit = rule['value']
                if rule['name'] == 'MAX_FACULTY_TOTAL_PER_DAY':
                    limit = max(limit, exceptions.get((faculty, day), limit))
                if value > limit:
                    warnings.append(f"{faculty} {day}: existing {value} periods exceeds {rule['name']}={rule['value']}")
    return warnings


def apply_constraint_handlers(model, data, entries, student_atoms, busy_slots):
    """Build actual hard bounds and weighted violation variables for every rule."""
    days, times = data['days'], data['times']
    count = len(times)
    penalties = []
    daily_exceptions = existing_daily_exceptions(data)
    day_cache = {}
    period_cache = {}
    hop_cache = {}
    all_faculty = {f for a in data['activities'] for f in a['faculty']}
    all_faculty.update(busy_slots)
    theory_busy = defaultdict(set)
    for b in data['busy']:
        component = str(b.get('component') or 'Unknown').lower()
        if component in ('lab', 'practical', 'ip', 'pse', 'non-teaching') + PROJECT_COMPONENTS:
            continue
        for p, (start, end) in enumerate(times):
            if overlaps(start, end, b['start'], b['end']):
                for f in b['faculty']:
                    theory_busy[f].add(days.index(b['day']) * count + p)

    def is_theory(activity):
        return activity['kind'].lower() in ('theory', 'lecture', 'tutorial')

    def match(rule, activity):
        scope, target = rule['scope'], rule['target']
        if scope == 'Global':
            return True
        if scope == 'Department':
            return activity['dept'] == target
        if scope == 'Course':
            return activity['code'] == target
        return target in {g.split(':', 1)[0] for g in activity['groups']}

    def on_day(index, day):
        key = (index, day)
        if key not in day_cache:
            a, _, present, start, _, _ = entries[index]
            equal = model.new_bool_var(f'day_eq_{index}_{day}')
            model.add_linear_constraint(start, day * count, (day + 1) * count - 1).only_enforce_if(equal)
            # Division gives the reverse implication without a disjunction.
            day_number = model.new_int_var(0, len(days) - 1, f'day_number_{index}_{day}')
            model.add_division_equality(day_number, start, count)
            model.add(day_number != day).only_enforce_if(equal.Not())
            active = model.new_bool_var(f'day_active_{index}_{day}')
            model.add(active <= present)
            model.add(active <= equal)
            model.add(active >= present + equal - 1)
            day_cache[key] = active
        return day_cache[key]

    def at_period(index, day, period):
        key = (index, day, period)
        if key not in period_cache:
            a, _, present, start, _, _ = entries[index]
            slot = day * count + period
            before = model.new_bool_var(f'before_{index}_{slot}')
            after = model.new_bool_var(f'after_{index}_{slot}')
            model.add(start <= slot).only_enforce_if(before)
            model.add(start > slot).only_enforce_if(before.Not())
            model.add(start + a['duration'] > slot).only_enforce_if(after)
            model.add(start + a['duration'] <= slot).only_enforce_if(after.Not())
            active = model.new_bool_var(f'period_active_{index}_{slot}')
            model.add(active <= present)
            model.add(active <= before)
            model.add(active <= after)
            model.add(active >= present + before + after - 2)
            period_cache[key] = active
        return period_cache[key]

    def maximum(values, label):
        if not values:
            return 0
        used = model.new_bool_var(label)
        model.add_max_equality(used, values)
        return used

    def bound(rule, expression, limit, upper):
        if rule['type'] == 'HARD':
            model.add(expression <= limit)
        else:
            excess = model.new_int_var(0, upper, f'violation_{len(penalties)}')
            model.add(excess >= expression - limit)
            penalties.append((excess, rule['weight'], upper))

    def forbid_gap(rule, occupancy, allowed=()):
        for p in range(1, count - 1):
            if p + 1 in allowed:
                continue
            earlier = maximum(occupancy[:p], f'earlier_{len(penalties)}_{p}')
            later = maximum(occupancy[p + 1:], f'later_{len(penalties)}_{p}')
            bound(rule, earlier + later - occupancy[p], 1, 1)

    morning = [p for p in range(count) if p == 0 or all(times[q][1] == times[q + 1][0] for q in range(p))]
    afternoon = [p for p in range(count) if p not in morning]
    if rule_enabled(data, 'PERIODS_PER_WEEK'):
        for _, _, present, _, _, _ in entries:
            model.add(present == 1)
    if rule_enabled(data, 'FIXED_SLOT'):
        fixed_entries = {(entry[0]['id'], entry[1] + 1): entry for entry in entries}
        for fixed in data.get('fixed_sessions', []):
            entry = fixed_entries[fixed['activity_id'], fixed['session']]
            model.add(entry[2] == 1)
            model.add(entry[3] == days.index(fixed['day']) * count + fixed['period'] - 1)
            if fixed['room']:
                for room, use in entry[4]:
                    model.add(use == int(room == fixed['room']))

    overrides = {(r['name'], r['target']) for r in data.get('constraints', []) if r['scope'] == 'Faculty'}
    scoped_depts = {(r['name'], r['target']) for r in data.get('constraints', []) if r['scope'] == 'Department' and r['enabled']}
    scoped_secs = {(r['name'], r['target']) for r in data.get('constraints', []) if r['scope'] == 'Section' and r['enabled']}
    for rule in data.get('constraints', default_constraints()):
        name = rule['name']
        if not rule['enabled']:
            continue
        if name.startswith('MAX_FACULTY_') or name in ('FACULTY_COMPACT_SCHEDULE', 'BALANCE_FACULTY_DAYS'):
            selected = [rule['target']] if rule['scope'] == 'Faculty' else sorted(all_faculty)
            for faculty in selected:
                if rule['scope'] == 'Global' and (name, faculty) in overrides:
                    continue
                own = [i for i, entry in enumerate(entries) if faculty in entry[0]['faculty']]
                if name == 'MAX_FACULTY_PERIODS_PER_WEEK':
                    load = sum(entries[i][0]['duration'] * entries[i][2] for i in own) + len(busy_slots[faculty])
                    bound(rule, load, rule['value'], len(days) * count + sum(entries[i][0]['duration'] for i in own))
                    continue
                if name == 'MAX_FACULTY_CONSECUTIVE':
                    # At most `value` periods in a row (lunch breaks a run), counting labs,
                    # theory and existing load. With value 3 this only bites on P1-P4.
                    if not own or is_placeholder(faculty):
                        continue
                    limit = int(rule['value'])
                    occupied = busy_slots[faculty]
                    for day in range(len(days)):
                        for start in range(count - limit):
                            window = range(start, start + limit + 1)
                            if any(times[k][1] != times[k + 1][0] for k in range(start, start + limit)):
                                continue
                            fixed = sum(1 for p in window if day * count + p in occupied)
                            terms = [at_period(i, day, p) for p in window if day * count + p not in occupied for i in own]
                            if terms and fixed + len(terms) > limit:
                                bound(rule, fixed + sum(terms), limit, fixed + len(terms))
                    continue
                loads = []
                for day in range(len(days)):
                    theory = name == 'MAX_FACULTY_THEORY_PER_DAY'
                    fixed = theory_busy[faculty] if theory else busy_slots[faculty]
                    occupied = sum(day * count <= p < (day + 1) * count for p in fixed)
                    indices = [i for i in own if not theory or is_theory(entries[i][0])]
                    load = sum(entries[i][0]['duration'] * on_day(i, day) for i in indices) + occupied
                    if name.startswith('MAX_FACULTY_'):
                        limit = rule['value']
                        if name == 'MAX_FACULTY_TOTAL_PER_DAY':
                            limit = max(limit, daily_exceptions.get((faculty, days[day]), limit))
                        # Existing overload is accepted, but closes the whole day
                        # to every new component, even when this cap is SOFT.
                        if occupied > limit:
                            for i in own:
                                model.add(on_day(i, day) == 0)
                        bound(rule, load, max(limit, occupied), count + sum(entries[i][0]['duration'] for i in indices))
                    elif name == 'FACULTY_COMPACT_SCHEDULE':
                        occupancy = [1 if day * count + p in busy_slots[faculty] else maximum([at_period(i, day, p) for i in own], f'faculty_used_{faculty}_{day}_{p}') for p in range(count)]
                        forbid_gap(rule, occupancy)
                    else:
                        daily = model.new_int_var(0, count + sum(entries[i][0]['duration'] for i in own), f'faculty_load_{faculty}_{day}')
                        model.add(daily == load)
                        loads.append(daily)
                if name == 'BALANCE_FACULTY_DAYS':
                    upper = count + sum(entries[i][0]['duration'] for i in own)
                    high = model.new_int_var(0, upper, f'high_{faculty}')
                    low = model.new_int_var(0, upper, f'low_{faculty}')
                    model.add_max_equality(high, loads)
                    model.add_min_equality(low, loads)
                    bound(rule, high - low, 0, upper)
            continue

        if name in ('FACULTY_ALLOWED_DAYS', 'FACULTY_ALLOWED_SLOTS', 'COURSE_ALLOWED_SLOTS', 'FACULTY_SLOTS_LAB_EXEMPT'):
            continue  # Enforced in legal start domains before session variables are created.
        own = [i for i, entry in enumerate(entries) if match(rule, entry[0])]
        if rule['scope'] == 'Global':  # a Department/Section row of the same rule replaces the Global one there
            own = [i for i in own if (name, entries[i][0]['dept']) not in scoped_depts
                   and not any((name, g.split(':', 1)[0]) in scoped_secs for g in entries[i][0]['groups'])]
        elif rule['scope'] == 'Department':
            own = [i for i in own if not any((name, g.split(':', 1)[0]) in scoped_secs for g in entries[i][0]['groups'])]
        if name == 'SAME_LAB_ROOM':
            groups = defaultdict(list)
            for i in own:
                a = entries[i][0]
                if a['kind'].lower() == 'lab' and len(entries[i][4]) > 1:
                    for root in sorted({g.split(':', 1)[0] for g in a['groups']}):
                        groups[(a['dept'], a['code'], root)].append(i)
            for key, indices in groups.items():
                if len(indices) < 2:
                    continue
                first = dict(entries[indices[0]][4])
                for i in indices[1:]:
                    other = dict(entries[i][4])
                    for room in set(first) | set(other):
                        u, v = first.get(room, 0), other.get(room, 0)
                        if rule['type'] == 'HARD':
                            model.add(u == v)
                        else:
                            bound(rule, u - v, 0, 1)
                            bound(rule, v - u, 0, 1)
            continue
        if name == 'CONSECUTIVE_DAYS_PER_COURSE':
            labs = [i for i in own if entries[i][0]['kind'].lower() == 'lab']
            if labs:
                used = [maximum([on_day(i, day) for i in labs], f'ccday_{rule["target"]}_{day}') for day in range(len(days))]
                for d1 in range(len(days)):
                    for d2 in range(d1 + 2, len(days)):
                        model.add(used[d1] + used[d2] <= 1)
            continue
        if name == 'MAX_DAYS_PER_COURSE':
            labs = [i for i in own if entries[i][0]['kind'].lower() == 'lab']
            if labs:
                used = [maximum([on_day(i, day) for i in labs], f'cday_{rule["target"]}_{day}') for day in range(len(days))]
                model.add(sum(used) <= int(rule['value']))
            continue
        if name == 'NO_FULL_DAY_LABS':
            labs = defaultdict(list)
            for i in own:
                if entries[i][0]['kind'].lower() == 'lab':
                    for group in student_atoms(entries[i][0]):
                        labs[group].append(i)
            am = [p for p in range(count) if times[p][1] <= LUNCH_START]
            pm = [p for p in range(count) if times[p][0] >= LUNCH_START]
            for group, indices in labs.items():
                if len(indices) < 2:
                    continue
                for day in range(len(days)):
                    a_m = maximum([at_period(i, day, p) for i in indices for p in am], f'fdl_am_{group}_{day}')
                    p_m = maximum([at_period(i, day, p) for i in indices for p in pm], f'fdl_pm_{group}_{day}')
                    bound(rule, a_m + p_m, 1, 2)
            continue
        if name == 'MAX_LAB_PERIODS_PER_DAY':
            labs = defaultdict(list)
            for i in own:
                if entries[i][0]['kind'].lower() == 'lab':
                    for group in student_atoms(entries[i][0]):
                        labs[group].append(i)
            for group, indices in labs.items():
                if sum(entries[i][0]['duration'] for i in indices) <= int(rule['value']):
                    continue
                for day in range(len(days)):
                    occ = [at_period(i, day, p) for i in indices for p in range(count)]
                    if rule['type'] == 'HARD':
                        model.add(sum(occ) <= int(rule['value']))
                    else:
                        bound(rule, sum(occ), int(rule['value']), len(occ))
            continue
        if name == 'LAB_DAY_LIMIT_EXCLUDE':
            continue
        if name == 'MAX_LABS_PER_DAY':
            excluded = {r['target'] for r in data.get('constraints', []) if r['name'] == 'LAB_DAY_LIMIT_EXCLUDE' and r['enabled']}
            labs = defaultdict(list)
            for i in own:
                a = entries[i][0]
                if a['kind'].lower() == 'lab' and a['code'] not in excluded:
                    for group in student_atoms(a):
                        labs[group].append(i)
            for group, indices in labs.items():
                if len(indices) <= int(rule['value']):
                    continue
                for day in range(len(days)):
                    flags = [on_day(i, day) for i in indices]
                    if rule['type'] == 'HARD':
                        model.add(sum(flags) <= int(rule['value']))
                    else:
                        bound(rule, sum(flags), int(rule['value']), len(flags))
            continue
        if name == 'MAX_LABS_PER_MORNING':
            labs = defaultdict(list)
            for i in own:
                a = entries[i][0]
                if a['kind'].lower() == 'lab':
                    for group in student_atoms(a):
                        labs[group].append(i)
            morning_p = [p for p in range(count) if times[p][1] <= LUNCH_START] or list(range(count // 2))
            for group, indices in labs.items():
                if len(indices) <= int(rule['value']):
                    continue
                for day in range(len(days)):
                    flags = [maximum([at_period(i, day, p) for p in morning_p], f'mlab_{i}_{day}') for i in indices]
                    bound(rule, sum(flags), int(rule['value']), len(flags))
            continue
        if name == 'P1_NO_IDLE':
            grouped = defaultdict(list)
            for i in own:
                for group in student_atoms(entries[i][0]):
                    grouped[group].append(i)
            for group, indices in grouped.items():
                root = group.split(':', 1)[0]
                for day, day_name in enumerate(days):
                  for pp in MUST_FILL_PERIODS:
                    fixed = any(b['section'] == root and b['day'] == day_name and overlaps(times[pp][0], times[pp][1], b['start'], b['end']) for b in data.get('section_busy', []))
                    if not fixed:
                        used = maximum([at_period(i, day, pp) for i in indices], f'p{pp+1}_required_{group}_{day}')
                        if pp in SOFT_FILL_PERIODS and root in SOFT_FILL_SECTIONS:
                            empty = model.new_int_var(0, 1, f'p{pp+1}_empty_{group}_{day}')
                            model.add(empty == 1 - used)
                            penalties.append((empty, SOFT_FILL_WEIGHT, 1))  # top priority, but allowed where the section lacks hours
                        else:
                            model.add(used == 1)
        elif name == 'PREFER_HOME_ROOM':
            for i in own:
                a = entries[i][0]
                homes = set(a.get('home_rooms', []))
                if is_theory(a) and homes:
                    for room, use in entries[i][4]:
                        if room not in homes:
                            bound(rule, use, 0, 1)
        elif name == 'AVOID_7TH_HOUR_THEORY':
            for i in own:
                if is_theory(entries[i][0]):
                    for day in range(len(days)):
                        bound(rule, at_period(i, day, rule['value'] - 1), 0, 1)
        elif name in ('MAX_SAME_SUBJECT_PER_DAY', 'PREFER_ONE_SUBJECT_PER_DAY', 'NO_CONSECUTIVE_THEORY_SAME_SUBJECT', 'MAX_SAME_THEORY_SUBJECT_IN_P1', 'THEORY_TIME_OF_DAY_BALANCE'):
            groups = defaultdict(list)
            for i in own:
                a = entries[i][0]
                if is_theory(a):
                    for group in student_atoms(a):
                        groups[group, a['code']].append(i)
            for key, indices in groups.items():
                if name == 'MAX_SAME_THEORY_SUBJECT_IN_P1':
                    bound(rule, sum(at_period(i, day, 0) for i in indices for day in range(len(days))), rule['value'], len(indices))
                elif name == 'THEORY_TIME_OF_DAY_BALANCE':
                    if afternoon:
                        am = sum(at_period(i, day, p) for i in indices for day in range(len(days)) for p in morning)
                        pm = sum(at_period(i, day, p) for i in indices for day in range(len(days)) for p in afternoon)
                        upper = count * sum(entries[i][0]['duration'] for i in indices)
                        difference = model.new_int_var(0, upper, f'time_balance_{len(penalties)}')
                        model.add_abs_equality(difference, len(afternoon) * am - len(morning) * pm)
                        bound(rule, difference, 0, upper)
                else:
                    for day in range(len(days)):
                        if name in ('MAX_SAME_SUBJECT_PER_DAY', 'PREFER_ONE_SUBJECT_PER_DAY'):
                            bound(rule, sum(entries[i][0]['duration'] * on_day(i, day) for i in indices), rule['value'], sum(entries[i][0]['duration'] for i in indices))
                        else:
                            for p in range(count - 1):
                                if times[p][1] == times[p + 1][0]:
                                    left = maximum([at_period(i, day, p) for i in indices], f'left_{key}_{day}_{p}')
                                    right = maximum([at_period(i, day, p + 1) for i in indices], f'right_{key}_{day}_{p}')
                                    bound(rule, left + right, 1, 1)
        elif name in ('NO_FACULTY_BLOCK_HOP', 'NO_SECTION_BLOCK_HOP', 'NO_FACULTY_FLOOR_HOP', 'NO_SECTION_FLOOR_HOP'):
            floor_rule = name.endswith('FLOOR_HOP')
            locf = floor_loc if floor_rule else building  # location used by this rule
            owners = defaultdict(list)
            for i in own:
                a = entries[i][0]
                if not any(locf(room) for room, _ in entries[i][4]):
                    continue
                keys = [f for f in a['faculty'] if not is_placeholder(f)] if name in ('NO_FACULTY_BLOCK_HOP', 'NO_FACULTY_FLOOR_HOP') else student_atoms(a)
                for k in keys:
                    owners[k].append(i)
            def in_building(i, day, p, b):
                key = (name[:6], floor_rule, i, day, p, b)
                if key not in hop_cache:
                    rooms = entries[i][4]
                    blds = {locf(room) for room, _ in rooms}
                    if b not in blds:
                        hop_cache[key] = None
                    elif len(blds) == 1:
                        hop_cache[key] = at_period(i, day, p)
                    else:
                        at = at_period(i, day, p)
                        chosen = sum(use for room, use in rooms if locf(room) == b)
                        z = model.new_bool_var(f'bld_{int(floor_rule)}_{i}_{day}_{p}_{b}')
                        model.add(z <= at)
                        model.add(z <= chosen)
                        model.add(z >= at + chosen - 1)
                        hop_cache[key] = z
                return hop_cache[key]
            for k, indices in owners.items():
                blds = sorted({locf(room) for i in indices for room, _ in entries[i][4]} - {None})
                if len(blds) < 2:
                    continue
                for day in range(len(days)):
                    for p in range(count - 1):
                        if times[p][1] != times[p + 1][0]:
                            continue  # lunch (or any break) separates these periods
                        here = {b: [v for v in (in_building(i, day, p, b) for i in indices) if v is not None] for b in blds}
                        nxt = {b: [v for v in (in_building(i, day, p + 1, b) for i in indices) if v is not None] for b in blds}
                        for b in blds:
                            other = [v for b2 in blds if b2 != b and (not floor_rule or b2[0] == b[0]) for v in nxt[b2]]
                            if here[b] and other:
                                bound(rule, sum(here[b]) + sum(other), 1, 1)
        elif name == 'NO_SECTION_ROOM_SWITCH':
            groups = defaultdict(list)
            for i in own:
                if is_theory(entries[i][0]):
                    for group in student_atoms(entries[i][0]):
                        groups[group].append(i)
            def in_room(i, day, p, r):
                key = ('room', i, day, p, r)
                if key not in hop_cache:
                    rooms = entries[i][4]
                    if r not in {room for room, _ in rooms}:
                        hop_cache[key] = None
                    elif len(rooms) == 1:
                        hop_cache[key] = at_period(i, day, p)
                    else:
                        at = at_period(i, day, p)
                        use = next(u for room, u in rooms if room == r)
                        z = model.new_bool_var(f'inroom_{i}_{day}_{p}_{r}')
                        model.add(z <= at)
                        model.add(z <= use)
                        model.add(z >= at + use - 1)
                        hop_cache[key] = z
                return hop_cache[key]
            for group, indices in groups.items():
                rms = sorted({room for i in indices for room, _ in entries[i][4]})
                if len(rms) < 2:
                    continue  # one room only: nothing can switch
                for day in range(len(days)):
                    for p in range(count - 1):
                        if times[p][1] != times[p + 1][0]:
                            continue
                        here = {r: [v for v in (in_room(i, day, p, r) for i in indices) if v is not None] for r in rms}
                        nxt = {r: [v for v in (in_room(i, day, p + 1, r) for i in indices) if v is not None] for r in rms}
                        for r in rms:
                            other = [v for r2 in rms if r2 != r for v in nxt[r2]]
                            if here[r] and other:
                                bound(rule, sum(here[r]) + sum(other), 1, 1)
        elif name == 'FACULTY_BACK_TO_BACK':
            by_faculty = defaultdict(list)
            for i in own:
                for f in entries[i][0]['faculty']:
                    if not is_placeholder(f):
                        by_faculty[f].append(i)
            for f, indices in by_faculty.items():
                occupied = busy_slots.get(f, set())
                for day in range(len(days)):
                    occ = [1 if day * count + p in occupied else maximum([at_period(i, day, p) for i in indices], f'b2b_used_{f}_{day}_{p}') for p in range(count)]
                    for p in range(count - 1):
                        if times[p][1] != times[p + 1][0] or (isinstance(occ[p], int) and isinstance(occ[p + 1], int)):
                            continue
                        cont = []
                        for i in indices:
                            if entries[i][0]['duration'] > 1:
                                c = model.new_bool_var(f'b2b_cont_{i}_{day}_{p}')
                                model.add(c <= at_period(i, day, p))
                                model.add(c <= at_period(i, day, p + 1))
                                cont.append(c)
                        bound(rule, occ[p] + occ[p + 1] - sum(cont), 1, 1)
        elif name == 'EXISTING_LOAD_GAP':
            by_faculty = defaultdict(list)
            for i in own:
                for f in entries[i][0]['faculty']:
                    if not is_placeholder(f) and busy_slots.get(f):
                        by_faculty[f].append(i)
            slot_bld = defaultdict(set)  # (faculty, slot) -> buildings of existing load; None = unknown
            for b in data['busy']:
                for q, (st, en) in enumerate(times):
                    if overlaps(st, en, b['start'], b['end']):
                        for f in b['faculty']:
                            slot_bld[(f, days.index(b['day']) * count + q)].add(building(b.get('room')) if b.get('room') else None)
            for f, indices in by_faculty.items():
                occupied = busy_slots[f]
                near = set()
                for slot in occupied:
                    day, p = divmod(slot, count)
                    blds = slot_bld.get((f, slot)) or {None}
                    bld = None if (None in blds or len(blds) != 1) else next(iter(blds))
                    for q in (p - 1, p + 1):
                        if 0 <= q < count and times[min(p, q)][1] == times[max(p, q)][0] and day * count + q not in occupied:
                            near.add((day, q, bld))
                unknown = {(d, q) for d, q, b in near if b is None}
                for day, q, bld in sorted(near, key=lambda t: (t[0], t[1], t[2] or '')):
                    if bld is not None and (day, q) in unknown:
                        continue  # already charged unconditionally
                    for i in indices:
                        if bld is None:
                            bound(rule, at_period(i, day, q), 0, 1)
                            continue
                        rooms = entries[i][4]
                        if not any(building(room) for room, _ in rooms):
                            bound(rule, at_period(i, day, q), 0, 1)  # room unknown: charge as before
                            continue
                        other = [use for room, use in rooms if building(room) not in (bld, None)]
                        if not other:
                            continue  # every room choice is in the same building
                        at = at_period(i, day, q)
                        if len(other) == len(rooms):
                            bound(rule, at, 0, 1)
                        else:
                            z = model.new_bool_var(f'elg_{i}_{day}_{q}')
                            for use in other:
                                model.add(z >= at + use - 1)
                            bound(rule, z, 0, 1)
        elif name in ('COMPACT_SCHEDULE', 'MINIMIZE_DISTINCT_ROOMS', 'NO_BACKTOBACK_LABS'):
            groups = defaultdict(list)
            for i in own:
                for group in student_atoms(entries[i][0]):
                    groups[group].append(i)
            for group, indices in groups.items():
                if name == 'COMPACT_SCHEDULE':
                    for day in range(len(days)):
                        occupancy = [maximum([at_period(i, day, p) for i in indices], f'student_used_{group}_{day}_{p}') for p in range(count)]
                        raw = str(data['settings'].get('Allowed student gap periods') or '')
                        allowed = {int(v.strip()) for v in raw.split(';') if v.strip()}
                        if any(p < 1 or p > count for p in allowed):
                            raise InputError('Allowed student gap periods must be valid period numbers')
                        forbid_gap(rule, occupancy, allowed)
                elif name == 'MINIMIZE_DISTINCT_ROOMS':
                    room_uses = defaultdict(list)
                    for i in indices:
                        if is_theory(entries[i][0]):
                            for room, use in entries[i][4]:
                                room_uses[room].append(use)
                    for room, uses in room_uses.items():
                        bound(rule, maximum(uses, f'room_used_{group}_{room}'), 0, 1)
                    # Floating sections (several room choices): also count distinct rooms per
                    # day and per half-day, so a room change mid-morning or mid-afternoon costs
                    # more than one at lunch, and one room all day costs least.
                    theory = [i for i in indices if is_theory(entries[i][0])]
                    if len(room_uses) > 1 and any(len(entries[i][4]) > 1 for i in theory):
                        for day in range(len(days)):
                            for half_name, half in (('am', morning), ('pm', afternoon), ('day', list(range(count)))):
                                per_room = defaultdict(list)
                                for i in theory:
                                    inside = maximum([at_period(i, day, p) for p in half], f'in_{half_name}_{i}_{day}_{group}')
                                    for room, use in entries[i][4]:
                                        if len(entries[i][4]) == 1:
                                            per_room[room].append(inside)
                                        else:
                                            z = model.new_bool_var(f'room_{half_name}_{i}_{day}_{room}_{group}')
                                            model.add(z <= use)
                                            model.add(z <= inside)
                                            model.add(z >= use + inside - 1)
                                            per_room[room].append(z)
                                if len(per_room) > 1:
                                    weight = rule['weight'] if half_name == 'day' else max(1, rule['weight'] // 2)
                                    for room, uses in per_room.items():
                                        penalties.append((maximum(uses, f'room_{half_name}_used_{group}_{day}_{room}'), weight, 1))
                else:
                    indices = [i for i in indices if entries[i][0]['kind'].lower() in ('lab', 'practical')]
                    for day in range(len(days)):
                        for p in range(count - 1):
                            if times[p][1] != times[p + 1][0]:
                                continue
                            left = []
                            right = []
                            for i in indices:
                                duration = entries[i][0]['duration']
                                if p + 1 >= duration:
                                    # Whole blocks are atomic. Ending here and
                                    # starting next period must be different labs.
                                    end_here = model.new_bool_var(f'lab_end_{i}_{day}_{p}')
                                    start = entries[i][3]
                                    target = day * count + p + 1 - duration
                                    equal = model.new_bool_var(f'lab_end_eq_{i}_{day}_{p}')
                                    model.add(start == target).only_enforce_if(equal)
                                    model.add(start != target).only_enforce_if(equal.Not())
                                    model.add_min_equality(end_here, [equal, entries[i][2]])
                                    left.append(end_here)
                                # At p+1 but absent at p means this block starts.
                                begins = model.new_bool_var(f'lab_begin_{i}_{day}_{p}')
                                model.add_min_equality(begins, [at_period(i, day, p + 1), at_period(i, day, p).Not()])
                                right.append(begins)
                            bound(rule, maximum(left, f'lab_left_{group}_{day}_{p}') + maximum(right, f'lab_right_{group}_{day}_{p}'), 1, 1)
    return penalties


# ----------------------------------------------------------------- 2. solve
def overlaps(a,b,c,d):return a<d and c<b
def atom_map(activities):
    """A whole section conflicts with each of its batches; disjoint batches may run together."""
    by_root=defaultdict(set)
    for a in activities:
        if a.get('eslot'):continue
        for g in a['groups']:
            if ':' in g:by_root[g.split(':',1)[0]].add(g)
    def atoms(groups):
        out=set()
        for g in groups:out.update(by_root[g] if ':' not in g and by_root[g] else [g])
        return out
    return atoms

def build_cp_model(data, hints=()):
    days,times=data['days'],data['times'];n=len(times);acts=data['activities']
    atoms=atom_map(acts);model=cp_model.CpModel();resources=defaultdict(list);shared_rooms=defaultdict(list);entries=[]
    leaders={}
    for x in acts:
        if x['eslot']:leaders.setdefault(x['eslot'],x['id'])
    def student_atoms(x):
        # Elective slot: one member (the leader) books the pooled sections as a whole;
        # the others are synchronised with it, so they need no student booking.
        if x['eslot']:return atoms(sorted({g.split(':',1)[0] for g in x['groups']})) if leaders[x['eslot']]==x['id'] else set()
        return atoms(x['groups'])
    busy_slots=defaultdict(set)
    for b in data['busy']:
        for q,(s,e) in enumerate(times):
            if overlaps(s,e,b['start'],b['end']):
                for f in b['faculty']:busy_slots[f].add(days.index(b['day'])*n+q)
    section_busy_slots = defaultdict(set)
    for b in data.get('section_busy', []):
        for q, (start, end) in enumerate(times):
            if overlaps(start, end, b['start'], b['end']):
                section_busy_slots[b['section']].add(days.index(b['day']) * n + q)
    allowed_faculty_slots = {r['target']: set(r['value']) for r in data.get('constraints', [])
                             if r['name'] == 'FACULTY_ALLOWED_SLOTS' and r['enabled']}
    allowed_faculty_days = {r['target']: set(r['value']) for r in data.get('constraints', [])
                            if r['name'] == 'FACULTY_ALLOWED_DAYS' and r['enabled']}
    lab_exempt = {r['target'] for r in data.get('constraints', []) if r['name'] == 'FACULTY_SLOTS_LAB_EXEMPT' and r['enabled']}
    allowed_course_slots = {r['target']: set(r['value']) for r in data.get('constraints', [])
                            if r['name'] == 'COURSE_ALLOWED_SLOTS' and r['enabled']}
    pinned_ids={x['activity_id'] for x in data.get('fixed_sessions',[])}
    for a in acts:
        previous=None;dur=a['duration']
        # Courses with no real teacher yet (TBC / DM / DGD placeholders) may only use P4, P6 and P7,
        # so the departments can staff them later from teachers free at those times. Pinned sessions keep their slot.
        tbc_only=a['id'] not in pinned_ids and dur<=2 and all(is_placeholder(f) for f in a['faculty'])
        # TBCS (Technical and Business Communication Skills, EAEXX07) and CSE-C3's soft-skills
        # equivalent (EAEXX08) also only in P4, P6 and P7.
        tbc_only=tbc_only or (a['id'] not in pinned_ids and a['code'] in TBCS_CODES)
        legal=[di*n+p for di in range(len(days)) for p in range(n-dur+1)
               if (not tbc_only or all(q+1 in TBC_PERIODS for q in range(p,p+dur)))
               and all(days[di] in allowed_faculty_days.get(f, set(days)) for f in a['faculty'])
               and all(f not in allowed_faculty_slots or (f in lab_exempt and a['kind'].lower() == 'lab') or all((days[di], q + 1) in allowed_faculty_slots[f] for q in range(p,p+dur)) for f in a['faculty'])
               and (a['code'] not in allowed_course_slots or all((days[di], q + 1) in allowed_course_slots[a['code']] for q in range(p,p+dur)))
               and not any(set(range(di*n+p,di*n+p+dur)) & section_busy_slots[g.split(':',1)[0]] for g in a['groups'])
               and all(times[k][1]==times[k+1][0] for k in range(p,p+dur-1))
               and (not rule_enabled(data, 'EXISTING_SCHEDULE_CLASH') or not any(set(range(di*n+p,di*n+p+dur))&busy_slots[f] for f in a['faculty']))]
        for sn in range(a['sessions']):
            on=model.new_bool_var(f'on_{a["id"]}_{sn}')
            st=model.new_int_var_from_domain(cp_model.Domain.from_values(legal or [0]),f'st_{a["id"]}_{sn}')
            if not legal:model.add(on==0)
            en=model.new_int_var(0,len(days)*n+dur,f'en_{a["id"]}_{sn}')
            iv=model.new_optional_interval_var(st,dur,en,on,f'iv_{a["id"]}_{sn}')
            for f in a['faculty']:resources['F',f].append(iv)
            for g in student_atoms(a):resources['G',g].append(iv)
            rooms=[]
            for ri,room in enumerate(a['rooms']):
                use=model.new_bool_var(f'r_{a["id"]}_{sn}_{ri}');rooms.append((room,use))
                room_iv=model.new_optional_interval_var(st,dur,en,use,f'ri_{a["id"]}_{sn}_{ri}')
                if str(room).strip().upper() in SHARED_ROOM_SECTIONS:
                    shared_rooms[room].append((room_iv,room_demand(a,room)))
                else:
                    resources['R',room].append(room_iv)
            model.add(sum(v for _,v in rooms)==on)
            if previous:  # identical sessions in order (symmetry breaking)
                ps,po=previous;model.add(on<=po);model.add(st>=ps+dur).only_enforce_if(on)
            previous=(st,on);entries.append((a,sn,on,st,rooms,bool(legal)))
    par=defaultdict(list)
    for a,sn,on,st,rr,lg in entries:
        if a['parallel']:par[a['parallel'],sn].append((on,st))
        if a['eslot']:par['ES:'+a['eslot'],sn].append((on,st))
    for items in par.values():
        for on2,st2 in items[1:]:model.add(items[0][0]==on2);model.add(items[0][1]==st2).only_enforce_if(on2)
    if rule_enabled(data, 'NO_ROOM_CLASH'):
        reserved = set()
        for b in data.get('section_busy', []):
            room = training_room(b['label'])
            if not room:
                continue
            slots = [q for q, (s0, e0) in enumerate(data['times']) if overlaps(s0, e0, b['start'], b['end'])]
            for q in slots:
                key = (room, b['day'], q)
                if key in reserved:
                    continue
                reserved.add(key)
                t = data['days'].index(b['day']) * len(data['times']) + q
                iv = model.new_fixed_size_interval_var(t, 1, f'train_{room}_{b["day"]}_{q}')
                if room.upper() in SHARED_ROOM_SECTIONS:
                    shared_rooms[room].append((iv, SHARED_ROOM_SECTIONS[room.upper()]))
                else:
                    resources['R', room].append(iv)
        for room, items in shared_rooms.items():
            model.add_cumulative([iv for iv, _ in items], [d for _, d in items], SHARED_ROOM_SECTIONS[str(room).strip().upper()])
    resource_rules = {'F': 'NO_FACULTY_CLASH', 'G': 'NO_SECTION_OVERLAP', 'R': 'NO_ROOM_CLASH'}
    for (resource_type, resource_id), intervals in resources.items():
        if rule_enabled(data, resource_rules[resource_type]):
            model.add_no_overlap(intervals)
    for f,lim in data['faculty_limits'].items():
        own=[(a,on,st) for a,sn,on,st,rr,lg in entries if f in a['faculty']];occ=busy_slots[f]
        if lim.get('weekly') is not None:model.add(sum(a['duration']*on for a,on,st in own)+len(occ)<=lim['weekly'])
        if lim.get('daily') is not None:
            for di in range(len(days)):
                terms=[]
                for k,(a,on,st) in enumerate(own):
                    dv=model.new_int_var(0,len(days)-1,f'dv_{f}_{di}_{k}');model.add_division_equality(dv,st,n)
                    same=model.new_bool_var(f'd_{f}_{di}_{k}');model.add(dv==di).only_enforce_if(same);model.add(dv!=di).only_enforce_if(same.Not())
                    at=model.new_bool_var(f'at_{f}_{di}_{k}');model.add_implication(at,on);model.add_implication(at,same);model.add(at>=on+same-1)
                    terms.append(a['duration']*at)
                model.add(sum(terms)+sum(di*n<=p<(di+1)*n for p in occ)<=lim['daily'])
    h={(x['activity_id'],x['session']-1):x for x in hints}
    used=0;free=data.get('free_depts');frozen=0;unfrozen=[]
    for a,sn,on,st,rr,lg in entries:
        x=h.get((a['id'],sn))
        ok=bool(x and x['day'] in days and x['duration']==a['duration'] and x['room'] in a['rooms'])
        if ok:
            model.add_hint(on,1);model.add_hint(st,days.index(x['day'])*n+x['period']-1);used+=1
            for room,v in rr:model.add_hint(v,int(room==x['room']))
        if free is not None and a['dept'] not in free and not (set(a['faculty']) & data.get('free_faculty', set())):
            if ok:
                model.add(on==1);model.add(st==days.index(x['day'])*n+x['period']-1)
                for room,v in rr:model.add(v==int(room==x['room']))
                frozen+=1
            else:unfrozen.append(f"{a['id']} s{sn+1}")
    if free is not None:
        print(f'Frozen from the hint timetable: {frozen} sessions outside {sorted(free)}'+(f' (classes of {len(data["free_faculty"])} listed teachers also free)' if data.get('free_faculty') else '')+'; re-solving the rest.',flush=True)
        if unfrozen:print(f'  {len(unfrozen)} sessions outside those departments could not be frozen (changed or missing in the hint) and are free:',', '.join(unfrozen[:20])+(' ...' if len(unfrozen)>20 else ''),flush=True)
    print(f'Model: {len(acts)} activities, {len(entries)} sessions, {sum(a["hours"] for a in acts)} periods; warm-start hints {used}',flush=True)
    penalties = apply_constraint_handlers(model, data, entries, student_atoms, busy_slots)
    provisional_rooms = [use for *_, room_choices, legal in entries for room, use in room_choices if room.startswith('TBC')]
    # Losing a period always costs more than every room/quality penalty combined.
    room_weight = max(1, SOFT_ROOM_WEIGHT)
    coverage_weight = 1 + room_weight * len(provisional_rooms) + sum(weight * upper for _, weight, upper in penalties)
    placed = sum(activity['duration'] * present for activity, _, present, _, _, _ in entries)
    model.maximize(coverage_weight * placed - room_weight * sum(provisional_rooms) - sum(weight * violation for violation, weight, _ in penalties))
    return model, entries


class SearchProgress(cp_model.CpSolverSolutionCallback):
    """Prints a line each time CP-SAT finds a better timetable: time, score (higher is better: placed periods
    minus soft-rule penalties), the best score still possible, and the gap between them. Gap 0 = proven optimal."""
    EVERY=100      # print every 100th improved timetable ...
    QUIET_SECS=60  # ... or when this many seconds have passed since the last line
    SAVE_EVERY=900  # seconds between interim saves (15 min)
    def __init__(self,on_save=None):
        super().__init__();self.n=0;self.last=-1e9;self.on_save=on_save;self.saved_at=None;self.saved_n=0
    def on_solution_callback(self):
        self.n+=1;obj=self.objective_value;bd=self.best_objective_bound
        if self.on_save and (self.n==1 or (self.wall_time-self.saved_at>=self.SAVE_EVERY and self.n>self.saved_n)):
            try:
                self.on_save(self.value,self.n==1,self.wall_time,obj)
                print(f"  [{self.wall_time:7.1f}s] saved timetable (solution #{self.n}, score {obj:,.0f})"+(' as first solution' if self.n==1 else ''),flush=True)
            except Exception as e:print('  interim save failed:',e,flush=True)
            self.saved_at=self.wall_time;self.saved_n=self.n
        if not (self.n==1 or self.n%self.EVERY==0 or self.wall_time-self.last>=self.QUIET_SECS):return
        self.last=self.wall_time
        gap=abs(bd-obj)
        print(f"  [{self.wall_time:7.1f}s] solution #{self.n:<3} score {obj:>14,.0f}   best possible {bd:>14,.0f}   gap {gap:>10,.0f} ({100.0*gap/max(1.0,abs(obj)):.4f}%)",flush=True)

def solve(data, seconds, workers=4, hints=(), on_save=None):
    days, times = data['days'], data['times']
    n = len(times)
    acts = data['activities']
    model, entries = build_cp_model(data, hints)
    validation_error = model.validate()
    if validation_error:
        raise InputError('Constraint model is invalid: ' + validation_error)
    collect=lambda value,ok,status_name,wall,quiet=False:_collect(data,entries,value,ok,status_name,wall,quiet)
    sv=cp_model.CpSolver();sv.parameters.max_time_in_seconds=seconds;sv.parameters.num_search_workers=workers;sv.parameters.random_seed=20261104
    print(f'Searching for up to {seconds}s with {workers} workers. Progress (first, every 100th improvement, or at least once a minute):',flush=True)
    sv.parameters.catch_sigint_signal=True
    if data.get('stop_at_first'):sv.parameters.stop_after_first_solution=True  # Ctrl+C: stop searching, keep the best timetable found
    def interim(value,first,wall,obj):
        r=collect(value,True,'FEASIBLE (saved during search)',wall,quiet=True);r['objective']=obj
        on_save(r,first)
    progress=SearchProgress(interim if on_save else None)
    print('Saves: first timetable at once, then every 15 min if improved. Ctrl+C stops and saves the best so far.',flush=True)
    status=sv.solve(model,progress)
    ok=status in (cp_model.FEASIBLE,cp_model.OPTIMAL)
    res=collect(sv.value,ok,sv.status_name(status),sv.wall_time)
    if ok:
        gap=abs(sv.best_objective_bound-sv.objective_value)
        print(f'Finished after {progress.n} improvements: {sv.status_name(status)}, score {sv.objective_value:,.0f}, best possible {sv.best_objective_bound:,.0f}, gap {gap:,.0f}',flush=True)
        res['objective']=sv.objective_value;res['bound']=sv.best_objective_bound;res['gap']=gap
    return res

def _collect(data,entries,value,ok,status_name,wall,quiet=False):
    days, times = data['days'], data['times']
    n = len(times)
    acts = data['activities']
    sv=None
    res={'status':status_name,'schedule':[],'unplaced':[],'seconds':wall}
    for a,sn,on,st,rr,lg in entries:
        if not ok or not value(on):
            res['unplaced'].append({'activity_id':a['id'],'session':sn+1,'periods':a['duration'],'reason':
                'No solution found within the time limit' if not ok else 'No legal start (breaks / faculty busy times)' if not lg else
                'Cannot fit with the other classes of its faculty, students and rooms'});continue
        k=value(st);di,p=divmod(k,n);room=next(r for r,v in rr if value(v))
        res['schedule'].append({'activity_id':a['id'],'course_id':a['course_id'],'course':a['name'],'groups':a['groups'],'faculty':a['faculty'],'room':room,
            'day':days[di],'period':p+1,'duration':a['duration'],'start':times[p][0],'end':times[p+a['duration']-1][1],'session':sn+1})
    res['required_periods']=sum(a['hours'] for a in acts);res['placed_periods']=sum(s['duration'] for s in res['schedule'])
    res['unplaced_periods']=sum(u['periods'] for u in res['unplaced'])
    if ok:
        times=data['times'];contig=lambda p:times[p][1]==times[p+1][0]
        occ_f=defaultdict(dict);occ_s=defaultdict(dict)
        for x in res['schedule']:
            b=building(x['room'])
            for q in range(x['period']-1,x['period']-1+x['duration']):
                for f in x['faculty']:
                    if not is_placeholder(f):occ_f[f,x['day']][q]=b
                for g in x['groups']:occ_s[g,x['day']][q]=b
        def hops(occ):return sum(1 for d in occ.values() for p in range(n-1) if contig(p) and d.get(p) and d.get(p+1) and d[p]!=d[p+1])
        busy=defaultdict(set)
        for b in data['busy']:
            for q,(s0,e0) in enumerate(times):
                if overlaps(s0,e0,b['start'],b['end']):
                    for f in b['faculty']:busy[f,b['day']].add(q)
        gaps=sum(1 for (f,d),qs in occ_f.items() for q in qs if any(q2 in busy[f,d] and contig(min(q,q2)) for q2 in (q-1,q+1)))
        res['hops']={'faculty_block_hops':hops(occ_f),'section_block_hops':hops(occ_s),'next_to_existing_load':gaps}
        if not quiet:print(f"Building hops: faculty {res['hops']['faculty_block_hops']}, sections {res['hops']['section_block_hops']}; classes next to existing load {gaps}",flush=True)
    audit(data,res);res['audit']='Independent resource checks passed for enabled clash rules; breaks respected; every required period placed or listed unplaced. Configured hard bounds enforced by CP-SAT.'
    return res

# ----------------------------------------------------------------- 3. independent check
def audit(data,res):
    acts={a['id']:a for a in data['activities']};atoms=atom_map(data['activities']);T=data['times'];seen=set();cnt=Counter()
    for s in res['schedule']:
        a=acts[s['activity_id']];key=(a['id'],s['session']);assert key not in seen;seen.add(key)
        for rule in data.get('constraints', []):
            if rule['name'] == 'FACULTY_ALLOWED_DAYS' and rule['enabled'] and rule['target'] in a['faculty']:
                assert s['day'] in rule['value'], ('Faculty outside allowed days', rule['target'], s['day'])
        for rule in data.get('constraints', []):
            if rule['name'] == 'FACULTY_ALLOWED_SLOTS' and rule['enabled'] and rule['target'] in a['faculty'] and not (a['kind'].lower() == 'lab' and any(x['name'] == 'FACULTY_SLOTS_LAB_EXEMPT' and x['enabled'] and x['target'] == rule['target'] for x in data.get('constraints', []))):
                assert all((s['day'], q) in rule['value'] for q in range(s['period'], s['period'] + a['duration'])), ('Faculty outside allowed slots', rule['target'])
        p=s['period']-1;d=a['duration'];assert s['room'] in a['rooms'] and p+d<=len(T)
        assert all(T[q][1]==T[q+1][0] for q in range(p,p+d-1)),'session crosses a break'
        for b in (data['busy'] if rule_enabled(data, 'EXISTING_SCHEDULE_CLASH') else []):assert not (s['day']==b['day'] and overlaps(s['start'],s['end'],b['start'],b['end']) and set(a['faculty'])&set(b['faculty'])),'busy time used'
        for b in data.get('section_busy', []):
            assert not (b['section'] in {g.split(':',1)[0] for g in a['groups']} and s['day'] == b['day'] and overlaps(s['start'],s['end'],b['start'],b['end'])), 'section busy time used'
        cnt[a['id']]+=d
    for u in res['unplaced']:key=(u['activity_id'],u['session']);assert key not in seen;seen.add(key);cnt[u['activity_id']]+=u['periods']
    assert all(cnt[a['id']]==a['hours'] for a in data['activities']),'period reconciliation failed'
    def stu(x):
        a=acts[x['activity_id']]
        return atoms(sorted({g.split(':',1)[0] for g in a['groups']})) if a['eslot'] else atoms(a['groups'])
    for g in {a['eslot'] for a in data['activities'] if a['eslot']}:
        pos=[sorted((s['session'],s['day'],s['period']) for s in res['schedule'] if s['activity_id']==a['id']) for a in data['activities'] if a['eslot']==g]
        assert all(p==pos[0] for p in pos),('Elective slot not synchronised',g)
    slots=defaultdict(list)
    for s in res['schedule']:
        for q in range(s['period'],s['period']+s['duration']):slots[s['day'],q].append(s)
    for ss in slots.values():
        for i,x in enumerate(ss):
            for y in ss[i+1:]:
                if rule_enabled(data, 'NO_FACULTY_CLASH'):assert not set(x['faculty'])&set(y['faculty']),('Faculty clash',x['activity_id'],y['activity_id'])
                ex,ey=acts[x['activity_id']]['eslot'],acts[y['activity_id']]['eslot']
                if rule_enabled(data, 'NO_SECTION_OVERLAP') and not (ex and ex==ey):assert not stu(x)&stu(y),('Student clash',x['activity_id'],y['activity_id'])
                if rule_enabled(data, 'NO_ROOM_CLASH') and str(x['room']).strip().upper() not in SHARED_ROOM_SECTIONS:assert x['room']!=y['room'],('Room clash',x['activity_id'],y['activity_id'])
    for g in {a['parallel'] for a in data['activities'] if a['parallel']}:
        pos=[sorted((s['session'],s['day'],s['period']) for s in res['schedule'] if s['activity_id']==a['id']) for a in data['activities'] if a['parallel']==g]
        assert all(p==pos[0] for p in pos),('Parallel group not synchronised',g)
    return True

# ----------------------------------------------------------------- 4. export (Excel + HTML)
def hm(m):return f'{m//60:02d}:{m%60:02d}'
def root(g):return g.split(':',1)[0]
def batch(g):return g.split(':',1)[1] if ':' in g else ''
def sec_key(r):
    m=re.match(r'(.+)-S(\d)-(.+)$',r);return (m[1],int(m[2]),m[3]) if m else (r,0,r)

def faculty_names(a):
    ids=a['faculty'];names=[x.strip() for x in re.split(r'[;]',a.get('faculty_names') or '') if x.strip()]
    if len(names)==len(ids):return dict(zip(ids,names))
    return {i:(names[0] if len(ids)==1 and names else i) for i in ids}

def build(model,result):
    acts={a['id']:a for a in model['activities']}
    times=model['times'];days=model['days']
    fname={}
    for a in model['activities']:fname.update(faculty_names(a))
    sessions=[]
    for s in result['schedule']:
        a=acts[s['activity_id']]
        sessions.append(dict(s,dept=a['dept'],semester=a['semester'],code=a['code'],kind=a['kind'],
            fac_names=[fname.get(f,f) for f in a['faculty']],roots=sorted({root(g) for g in a['groups']}),batches=sorted({batch(g) for g in a['groups'] if batch(g)})))
    unplaced=[]
    for u in result['unplaced']:
        a=acts[u['activity_id']]
        unplaced.append(dict(u,dept=a['dept'],semester=a['semester'],code=a['code'],course=a['name'],kind=a['kind'],groups=', '.join(a['groups']),
            faculty=', '.join(fname.get(f,f) for f in a['faculty']),faculty_ids=', '.join(a['faculty'])))
    return acts,fname,sessions,unplaced

def cell_text(s,focus_root=None):
    t=f"{s['course']} [{s['code']}]"
    if s['batches']:t=' / '.join(s['batches'])+': '+t
    pass  # no 'merged' note in cells
    t+='\n'+'; '.join(s['fac_names'])+'\n'+s['room']
    return t

THIN=Side(style='thin',color='B8C4D3');BORDER=Border(left=THIN,right=THIN,top=THIN,bottom=THIN)
FIXEDBLK=PatternFill('solid',fgColor='FCE4CC');THEORY=PatternFill('solid',fgColor='EAF1FB');HEAD=PatternFill('solid',fgColor='203B60');SUB=PatternFill('solid',fgColor='E7EDF6');LUNCH=PatternFill('solid',fgColor='EEEEEE')
TBC=PatternFill('solid',fgColor='FFF4D6');LAB=PatternFill('solid',fgColor='E6F4EA');PSE=PatternFill('solid',fgColor='F1E8FA')
WRAP=Alignment(wrap_text=True,vertical='top');CENTER=Alignment(horizontal='center',vertical='center',wrap_text=True)

def table(ws,r0,headers,rows,widths=None):
    for j,h in enumerate(headers,1):
        c=ws.cell(r0,j,h);c.font=Font(bold=True,color='FFFFFF');c.fill=HEAD;c.alignment=CENTER;c.border=BORDER
    for i,row in enumerate(rows,r0+1):
        for j,v in enumerate(row,1):
            c=ws.cell(i,j,v);c.alignment=WRAP;c.border=BORDER
    if widths:
        for j,w in enumerate(widths,1):ws.column_dimensions[get_column_letter(j)].width=max(ws.column_dimensions[get_column_letter(j)].width or 0,w)
    return r0+len(rows)+1

def write_excel(model,result,acts,fname,sessions,unplaced,path,meta):
    wb=Workbook();ws=wb.active;ws.title='Summary'
    times=model['times'];days=model['days'];n=len(times)
    ws['A1']=meta['title'];ws['A1'].font=Font(bold=True,size=15,color='203B60')
    ws['A2']='Placeholder faculty/rooms (TBC) are unverified. Every required period is either placed or listed as unplaced.';ws['A2'].font=Font(italic=True,color='8A5A00')
    items=[('Generated',meta['generated']),('Input basis',meta['basis']),('Solver status',result['status']),('Search time (s)',round(result['seconds'],1)),
        ('Required activity-periods',result['required_periods']),('Placed activity-periods',result['placed_periods']),('Unplaced activity-periods',result['unplaced_periods']),
        ('Unplaced sessions',len(result['unplaced'])),('Not scheduled by rule (MOOC / unresolved elective alternatives)',len(model.get('not_scheduled',[]))),
        ('Teaching week',', '.join(days)+f'; {n} periods, {hm(times[0][0])}-{hm(times[-1][1])}'),('Scope',meta.get('scope','')),
        ('Audit',result.get('audit',''))]
    r=table(ws,4,['Item','Value'],items,[48,110])
    by=defaultdict(lambda:[0,0,0]);
    for a in model['activities']:by[a['dept']][0]+=a['hours']
    for s in sessions:by[s['dept']][1]+=s['duration']
    for u in unplaced:by[u['dept']][2]+=u['periods']
    table(ws,r+1,['Department','Required periods','Placed','Unplaced'],[[d,*by[d]] for d in sorted(by)])
    ws.column_dimensions['C'].width=12;ws.column_dimensions['D'].width=12

    depts=sorted({a['dept'] for a in model['activities']})
    for d in depts:
        sh=wb.create_sheet(d[:31]);sh.sheet_view.showGridLines=False
        sh['A1']=f'{d} - master timetable grid';sh['A1'].font=Font(bold=True,size=14,color='203B60')
        req=sum(a['hours'] for a in model['activities'] if a['dept']==d);pl=sum(s['duration'] for s in sessions if s['dept']==d)
        sh['A2']=f'Provisional. Required {req} periods, placed {pl}, unplaced {req-pl}. Cell = course [code] / faculty / room. B1/B2 = lab batches. Colours: blue = theory, green = lab, purple = PSE/IP, orange = fixed training (with its time). TBC in a cell = placeholder faculty or room.'
        sh['A2'].font=Font(italic=True,color='4C607A')
        roots=sorted({root(g) for a in model['activities'] if a['dept']==d for g in a['groups']},key=sec_key)
        row=4;sh.column_dimensions['A'].width=7;sh.column_dimensions['B'].width=13
        for sem in sorted({sec_key(r)[1] for r in roots}):
            cols=[r for r in roots if sec_key(r)[1]==sem]
            sh.cell(row,1,f'Semester {sem}').font=Font(bold=True,size=12,color='203B60');row+=1
            for j,h in enumerate(['Day','Period']+[f'S{sem}-{sec_key(c)[2]}' for c in cols],1):
                c=sh.cell(row,j,h);c.font=Font(bold=True,color='FFFFFF');c.fill=HEAD;c.alignment=CENTER;c.border=BORDER
                if j>2:sh.column_dimensions[get_column_letter(j)].width=30
            sh.freeze_panes=None;row+=1
            grid=defaultdict(list);blocked={}
            for b in model.get('section_busy',[]):
                if b['section'] in cols:
                    for q in range(n):
                        if overlaps(times[q][0],times[q][1],b['start'],b['end']):blocked[b['section'],b['day'],q]=b['label']
            for s in sessions:
                for rt in s['roots']:
                    if rt in cols:
                        for q in range(s['period']-1,s['period']-1+s['duration']):grid[rt,s['day'],q].append(s)
            for day in days:
                for q in range(n):
                    if q and times[q-1][1]!=times[q][0]:
                        sh.cell(row,1,DAYS_SHORT.get(day,day));sh.cell(row,2,f'Break\n{hm(times[q-1][1])}-{hm(times[q][0])}')
                        for j in range(1,len(cols)+3):
                            c=sh.cell(row,j);c.fill=LUNCH;c.border=BORDER;c.alignment=CENTER
                            if j>2:c.value='LUNCH'
                        sh.row_dimensions[row].height=18;row+=1
                    sh.cell(row,1,DAYS_SHORT.get(day,day)).alignment=CENTER;sh.cell(row,2,f'P{q+1}\n{hm(times[q][0])}-{hm(times[q][1])}').alignment=CENTER
                    sh.cell(row,1).border=sh.cell(row,2).border=BORDER;sh.cell(row,1).font=Font(bold=True)
                    lines=0
                    for j,rt in enumerate(cols,3):
                        ss=sorted(grid[rt,day,q],key=lambda s:s['batches'])
                        c=sh.cell(row,j,'\n---\n'.join(cell_text(s) for s in ss) or None);c.alignment=WRAP;c.border=BORDER
                        if not ss and (rt,day,q) in blocked:c.value=blocked[rt,day,q];c.fill=FIXEDBLK;c.font=Font(italic=True,color='8A4B00')
                        if ss:
                            pinned=model.get('fixed_ids',set())
                            if any(s['activity_id'] in pinned for s in ss):
                                c.fill=FIXEDBLK
                            elif any(s['kind'] in ('IP','PSE') for s in ss):c.fill=PSE
                            elif any(s['kind']=='Lab' for s in ss):c.fill=LAB
                            else:c.fill=THEORY
                        lines=max(lines,str(c.value or '').count('\n')+1)
                    sh.row_dimensions[row].height=max(45,min(150,13*lines));row+=1
            row+=1
        du=[u for u in unplaced if u['dept']==d]
        sh.cell(row,1,f'Unplaced sessions - {d} ({sum(u["periods"] for u in du)} periods)').font=Font(bold=True,size=12,color='9B1C1C');row+=1
        if du:
            hdr=['Sem','Course code','Course','Component','Student groups','Faculty','Session','Periods','Reason']
            row=table(sh,row,hdr,[[u['semester'],u['code'],u['course'],u['kind'],u['groups'],u['faculty'],u['session'],u['periods'],u['reason']] for u in du])
        else:sh.cell(row,1,'None - every required period for this department is placed.')
        sh.freeze_panes='C4'

    s=wb.create_sheet('Unplaced')
    s['A1']='Unplaced sessions (all departments)';s['A1'].font=Font(bold=True,size=14,color='9B1C1C')
    s['A2']=f"{result['unplaced_periods']} activity-periods in {len(result['unplaced'])} sessions are not placed."
    table(s,4,['Department','Sem','Activity ID','Course code','Course','Component','Student groups','Faculty','Faculty IDs','Session','Periods','Reason'],
        [[u['dept'],u['semester'],u['activity_id'],u['code'],u['course'],u['kind'],u['groups'],u['faculty'],u['faculty_ids'],u['session'],u['periods'],u['reason']] for u in sorted(unplaced,key=lambda u:(u['dept'],u['semester'],u['activity_id']))],
        [12,6,24,12,34,10,26,30,18,8,8,60])
    s=wb.create_sheet('Not scheduled')
    s['A1']='Kept for records, not timetabled (MOOCs, unresolved elective alternatives, aggregates merged into section sessions)';s['A1'].font=Font(bold=True,size=13)
    ns=model.get('not_scheduled',[])
    table(s,3,['Status','Department','Sem','Ref','Course code','Course','Groups','Faculty','Periods','Reason'],
        [[x['status'],x['dept'],x['semester'],x['id'],x['code'],x['course'],x['groups'],x['faculty'],x['periods'],x['reason']] for x in ns],[26,10,6,26,12,36,30,34,8,70])
    cl=model.get('clarifications',[])
    if cl:
        s=wb.create_sheet('Clarifications');s['A1']='Points to confirm with departments';s['A1'].font=Font(bold=True,size=13)
        table(s,3,['Department','Year','Question'],cl,[22,14,110])
    s=wb.create_sheet('Sessions')
    table(s,1,['Department','Sem','Day','Period','Start','End','Duration','Course code','Course','Component','Student groups','Faculty','Faculty IDs','Room','Activity ID','Session'],
        [[x['dept'],x['semester'],x['day'],x['period'],hm(x['start']),hm(x['end']),x['duration'],x['code'],x['course'],x['kind'],', '.join(x['groups']),'; '.join(x['fac_names']),', '.join(x['faculty']),x['room'],x['activity_id'],x['session']]
         for x in sorted(sessions,key=lambda x:(x['dept'],x['semester'],model['days'].index(x['day']),x['period']))],[12,6,11,7,7,7,8,12,34,10,26,34,18,18,24,8])
    s=wb.create_sheet('Classes scheduled')
    s['A2']='Classes scheduled';s['A2'].font=Font(bold=True,size=13)
    s['A4']='The classes the solver placed, expanded from the workload workbook (one row per section, merged group, lab batch or elective subject). For checking only; edit the workload workbook instead.'
    table(s,8,INPUT_HEADERS,[[a['id'],a['dept'],a['semester'],a['code'],a['name'],a['kind'],'; '.join(a['groups']),'; '.join(a['faculty']),a.get('faculty_names',''),
        a['hours'],a['sessions'],a['duration'],'; '.join(a['rooms']),a.get('parallel',''),a.get('eslot',''),a.get('source',''),a.get('assumptions','')] for a in model['activities']],
        [24,10,6,12,34,10,26,22,34,8,8,8,28,30,60,16])
    s=wb.create_sheet('Assumptions')
    table(s,1,['Activity','Field','Value','Reason'],meta['notes'],[26,28,40,90])
    wb.save(path)

PAGE=r'''<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Timetable</title><style>
:root{--ink:#1d2b3a;--muted:#5d6f84;--line:#d5dde8;--bg:#f5f7fa;--card:#fff;--head:#203b60;--lab:#e6f4ea;--pse:#f1e8fa;--tbc:#fff4d6;--lunch:#eee}
*{box-sizing:border-box}body{margin:0;font:14px/1.4 system-ui,Segoe UI,Arial,sans-serif;color:var(--ink);background:var(--bg)}
header{background:var(--head);color:#fff;padding:18px 24px}h1{margin:0;font-size:22px}header p{margin:4px 0 0;color:#cfdaea;font-size:13px}
.wrap{padding:16px 24px;max-width:1800px;margin:auto}.bar{display:flex;flex-wrap:wrap;gap:10px;align-items:end;background:var(--card);border:1px solid var(--line);border-radius:8px;padding:12px;margin-bottom:14px}
.tabs button{border:1px solid var(--line);background:#fff;padding:8px 14px;border-radius:6px;cursor:pointer;font:inherit}.tabs button.on{background:var(--head);color:#fff;border-color:var(--head)}
label{display:grid;gap:4px;font-size:12px;color:var(--muted)}select,input{font:inherit;padding:7px 9px;border:1px solid #b9c6d6;border-radius:6px;min-width:150px}
.stats{display:flex;gap:10px;flex-wrap:wrap;margin-bottom:12px}.stat{background:#fff;border:1px solid var(--line);border-radius:8px;padding:8px 14px}.stat b{display:block;font-size:20px}
.gw{overflow:auto;border:1px solid var(--line);border-radius:8px;background:#fff;margin-bottom:18px}table.g{border-collapse:collapse;min-width:100%}
.g th,.g td{border:1px solid var(--line);padding:5px 6px;vertical-align:top;font-size:12px}.g th{background:var(--head);color:#fff;position:sticky;top:0;z-index:1}
.g td.t{white-space:nowrap;font-weight:600;background:#f0f3f8;position:sticky;left:0}.g tr.l td{background:var(--lunch);text-align:center;color:#777;font-size:11px}
.s{border-radius:5px;padding:4px 5px;margin:1px 0;background:#eef3fb}.s.lab{background:var(--lab)}.s.pse{background:var(--pse)}.s.tbc{background:var(--tbc)}
.s b{display:block}.s i{color:var(--muted);font-style:normal;display:block}.s small{color:#40566f}h2{font-size:17px;margin:14px 0 8px}
.ul{width:100%;border-collapse:collapse;background:#fff}.ul th,.ul td{border:1px solid var(--line);padding:5px 7px;font-size:12px;text-align:left}.ul th{background:#f0f3f8}
.legend span{display:inline-block;padding:2px 8px;border-radius:4px;margin-right:6px;font-size:12px}.muted{color:var(--muted)}
@media print{header,.bar{display:none}.gw{overflow:visible}}
</style></head><body><header><h1>__TITLE__</h1><p id="sub"></p></header>
<div class="wrap"><div class="stats" id="stats"></div>
<div class="bar"><div class="tabs" id="tabs"></div><label id="ldept">Department<select id="dept"></select></label><label id="lsem">Semester<select id="sem"></select></label>
<label id="lent">Show<select id="ent"></select></label><label id="lq">Find<input id="q" placeholder="name, ID or room"></label>
<div class="legend"><span style="background:#eef3fb">Theory</span><span style="background:var(--lab)">Lab</span><span style="background:var(--pse)">IP/PSE</span><span style="background:var(--tbc)">Placeholder</span></div></div>
<div id="view"></div></div>
<script>const D=__PAYLOAD__;
const $=id=>document.getElementById(id),esc=s=>String(s??'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const root=g=>g.split(':')[0],secOf=r=>{const m=r.match(/^(.+)-S(\d)-(.+)$/);return m?{d:m[1],s:+m[2],x:m[3]}:{d:r,s:0,x:r}};
const hm=m=>String(Math.floor(m/60)).padStart(2,'0')+':'+String(m%60).padStart(2,'0');
let mode='Master';const modes=['Master','Section','Faculty','Room','Unplaced'];
$('sub').textContent=D.meta.generated+' · '+D.meta.basis;
$('stats').innerHTML=[['Required periods',D.req],['Placed',D.placed],['Unplaced',D.unplaced],['Sessions',D.sessions.length],['Not scheduled (rule)',D.not_scheduled.length]].map(([a,b])=>`<div class=stat>${a}<b>${b}</b></div>`).join('');
$('tabs').innerHTML=modes.map(m=>`<button data-m="${m}">${m}</button>`).join('');
document.querySelectorAll('#tabs button').forEach(b=>b.onclick=()=>{mode=b.dataset.m;setup();});
const depts=[...new Set(D.acts.map(a=>a.dept))].sort();$('dept').innerHTML=depts.map(d=>`<option>${d}</option>`).join('');
function cls(s){if(s.faculty.some(f=>/^(TBC|DM\d|DGD\d|TP\d)/.test(f))||s.room.startsWith('TBC'))return 'tbc';if(s.kind==='IP'||s.kind==='PSE')return 'pse';if(s.kind==='Lab')return 'lab';return ''}
function box(s,show){const b=s.batches.length?s.batches.join('/')+': ':'';const mg='';
 return `<div class="s ${cls(s)}"><b>${esc(b+s.course)} <small>[${esc(s.code)}]</small></b>${show.includes('sec')?`<i>${esc(s.roots.map(r=>r.replace(/-S(\d)-/,' S$1-')).join(', '))}${esc(b?' '+s.batches.join('/'):'')}</i>`:''}${show.includes('fac')?`<i>${esc(s.fac_names.join('; '))}${mg}</i>`:''}${show.includes('room')?`<small>${esc(s.room)}</small>`:''}</div>`}
function grid(cols,colname,pick,show){const n=D.times.length;let h='<div class=gw><table class=g><tr><th>Day</th><th>Period</th>'+cols.map(c=>`<th>${esc(colname(c))}</th>`).join('')+'</tr>';
 for(const day of D.days){for(let q=0;q<n;q++){if(q&&D.times[q-1][1]!==D.times[q][0])h+=`<tr class=l><td class=t>${day.slice(0,3)}</td><td>Break ${hm(D.times[q-1][1])}-${hm(D.times[q][0])}</td>`+cols.map(()=>'<td>LUNCH</td>').join('')+'</tr>';
  h+=`<tr><td class=t>${day.slice(0,3)}</td><td class=t>P${q+1}<br><span class=muted>${hm(D.times[q][0])}-${hm(D.times[q][1])}</span></td>`;
  for(const c of cols){const ss=D.sessions.filter(s=>s.day===day&&q>=s.period-1&&q<s.period-1+s.duration&&pick(s,c));
   const busy=D.busyBy&&D.busyBy[c]&&D.busyBy[c].some(b=>b.day===day&&b.start<D.times[q][1]&&D.times[q][0]<b.end);
   h+='<td>'+ss.map(s=>box(s,show)).join('')+(busy&&!ss.length?'<div class="s" style="background:#e3e3e3"><i>Existing load</i></div>':'')+'</td>';}
  h+='</tr>'}}return h+'</table></div>'}
function week(key,pick,show){const n=D.times.length;let h='<div class=gw><table class=g><tr><th>Period</th>'+D.days.map(d=>`<th>${d}</th>`).join('')+'</tr>';
 for(let q=0;q<n;q++){if(q&&D.times[q-1][1]!==D.times[q][0])h+='<tr class=l><td class=t>Break</td>'+D.days.map(()=>`<td>${hm(D.times[q-1][1])}-${hm(D.times[q][0])}</td>`).join('')+'</tr>';
  h+=`<tr><td class=t>P${q+1}<br><span class=muted>${hm(D.times[q][0])}-${hm(D.times[q][1])}</span></td>`;
  for(const day of D.days){const ss=D.sessions.filter(s=>s.day===day&&q>=s.period-1&&q<s.period-1+s.duration&&pick(s,key));
   const busy=D.busyBy[key]&&D.busyBy[key].some(b=>b.day===day&&b.start<D.times[q][1]&&D.times[q][0]<b.end);
   h+='<td style="min-width:170px">'+ss.map(s=>box(s,show)).join('')+(busy&&!ss.length?'<div class="s" style="background:#e3e3e3"><i>Existing load</i></div>':'')+'</td>'}
  h+='</tr>'}return h+'</table></div>'}
function unplacedTable(list){if(!list.length)return '<p class=muted>No unplaced sessions.</p>';return '<table class=ul><tr><th>Dept</th><th>Sem</th><th>Course</th><th>Component</th><th>Groups</th><th>Faculty</th><th>Session</th><th>Periods</th><th>Reason</th></tr>'+list.map(u=>`<tr><td>${esc(u.dept)}</td><td>${u.semester}</td><td>${esc(u.course)} [${esc(u.code)}]</td><td>${esc(u.kind)}</td><td>${esc(u.groups)}</td><td>${esc(u.faculty)}</td><td>${u.session}</td><td>${u.periods}</td><td>${esc(u.reason)}</td></tr>`).join('')+'</table>'}
function setup(){document.querySelectorAll('#tabs button').forEach(b=>b.classList.toggle('on',b.dataset.m===mode));
 $('ldept').style.display=['Master','Section','Unplaced'].includes(mode)?'':'none';$('lsem').style.display=['Master','Section'].includes(mode)?'':'none';
 $('lent').style.display=['Section','Faculty','Room'].includes(mode)?'':'none';$('lq').style.display=['Faculty','Room'].includes(mode)?'':'none';fill();}
function sections(){const d=$('dept').value;return [...new Set(D.acts.filter(a=>a.dept===d).flatMap(a=>a.groups.map(root)))].sort()}
function fill(){const d=$('dept').value;const sems=[...new Set(sections().map(r=>secOf(r).s))].sort();const cur=$('sem').value;$('sem').innerHTML=sems.map(s=>`<option ${String(s)===cur?'selected':''}>${s}</option>`).join('');
 let opts=[];if(mode==='Section')opts=sections().filter(r=>String(secOf(r).s)===$('sem').value).map(r=>[r,r]);
 if(mode==='Faculty'){const q=$('q').value.toLowerCase();opts=Object.entries(D.fname).filter(([i,n])=>!q||(i+' '+n).toLowerCase().includes(q)).sort((a,b)=>a[1].localeCompare(b[1])).map(([i,n])=>[i,n+' ('+i+')'])}
 if(mode==='Room'){const q=$('q').value.toLowerCase();opts=[...new Set(D.sessions.map(s=>s.room))].filter(r=>!q||r.toLowerCase().includes(q)).sort().map(r=>[r,r])}
 const keep=$('ent').value;$('ent').innerHTML=opts.map(([v,t])=>`<option value="${esc(v)}" ${v===keep?'selected':''}>${esc(t)}</option>`).join('');render();}
function render(){const d=$('dept').value,sem=+$('sem').value,e=$('ent').value;let h='';
 if(mode==='Master'){const cols=sections().filter(r=>secOf(r).s===sem);h=`<h2>${esc(d)} — Semester ${sem} master grid</h2>`+grid(cols,c=>'S'+sem+'-'+secOf(c).x,(s,c)=>s.roots.includes(c),['fac','room'])+`<h2>Unplaced — ${esc(d)}</h2>`+unplacedTable(D.unplacedList.filter(u=>u.dept===d))}
 else if(mode==='Section'){h=`<h2>${esc(e)}</h2>`+week(e,(s,c)=>s.roots.includes(c),['fac','room'])+'<h2>Unplaced for this section</h2>'+unplacedTable(D.unplacedList.filter(u=>u.groups.split(', ').map(root).includes(e)))}
 else if(mode==='Faculty'){h=`<h2>${esc(D.fname[e]||e)} <span class=muted>(${esc(e)})</span></h2>`+week(e,(s,c)=>s.faculty.includes(c),['sec','room'])+'<h2>Unplaced sessions for this faculty</h2>'+unplacedTable(D.unplacedList.filter(u=>u.faculty_ids.split(', ').includes(e)))}
 else if(mode==='Room'){h=`<h2>Room ${esc(e)}</h2>`+week(e,(s,c)=>s.room===c,['sec','fac'])}
 else {h=`<h2>Unplaced sessions — ${esc(d)}</h2>`+unplacedTable(D.unplacedList.filter(u=>u.dept===d))+'<h2>All departments</h2>'+unplacedTable(D.unplacedList)+'<h2>Not scheduled by rule</h2><table class=ul><tr><th>Status</th><th>Dept</th><th>Course</th><th>Groups</th><th>Faculty</th><th>Reason</th></tr>'+D.not_scheduled.map(x=>`<tr><td>${esc(x.status)}</td><td>${esc(x.dept)}</td><td>${esc(x.course)} [${esc(x.code)}]</td><td>${esc(x.groups)}</td><td>${esc(x.faculty)}</td><td>${esc(x.reason)}</td></tr>`).join('')+'</table>'}
 $('view').innerHTML=h}
['dept','sem'].forEach(i=>$(i).onchange=fill);$('ent').onchange=render;$('q').oninput=fill;setup();
</script></body></html>'''

def write_html(model,result,acts,fname,sessions,unplaced,path,meta):
    busyBy=defaultdict(list)
    for b in model['busy']:
        for f in b['faculty']:busyBy[f].append({'day':b['day'],'start':b['start'],'end':b['end']})
    payload={'meta':meta,'days':model['days'],'times':model['times'],'acts':[{'dept':a['dept'],'groups':a['groups']} for a in model['activities']],
        'sessions':[{k:s[k] for k in ['day','period','duration','course','code','kind','room','faculty','fac_names','roots','batches','dept','semester']} for s in sessions],
        'unplacedList':unplaced,'not_scheduled':model.get('not_scheduled',[]),'fname':{f:n for f,n in fname.items()},'busyBy':busyBy,
        'req':result['required_periods'],'placed':result['placed_periods'],'unplaced':result['unplaced_periods']}
    path.write_text(PAGE.replace('__TITLE__',html.escape(meta['title'])).replace('__PAYLOAD__',json.dumps(payload,ensure_ascii=False).replace('<','\\u003c')),encoding='utf-8')


# ----------------------------------------------------------------- 5. rich HTML (Master / Section / Teacher / Room / Course)
# Same look and views as the Sem-1 rich_export.py, adapted to this solver: several
# classes can share one period (lab batches B1/B2, elective pools), faculty are keyed
# by "Name (ID)", and first-year commitments show as existing load.
RX_DEPT_COLORS=[("#7c3aed","#ede9fe","#4c1d95"),("#0369a1","#e0f2fe","#0c4a6e"),("#065f46","#d1fae5","#064e3b"),("#b45309","#fef3c7","#78350f"),
    ("#be123c","#ffe4e6","#881337"),("#0f766e","#ccfbf1","#134e4a"),("#4338ca","#e0e7ff","#312e81"),("#a16207","#fef9c3","#713f12"),
    ("#9d174d","#fce7f3","#831843"),("#166534","#dcfce7","#14532d"),("#1d4ed8","#dbeafe","#1e3a8a"),("#7e22ce","#f3e8ff","#581c87"),("#c2410c","#ffedd5","#7c2d12")]
RX_TYPE={"training":("#fde7d0","#8a4b00"),"theory":("#edf3ff","#1f2a60"),"lab":("#fff0f5","#a51d4d"),"project":("#eef9ee","#23752c")}

ABBR_OVERRIDES={'quantitative aptitude and logical reasoning':'QALR','verbal ability and critical reasoning':'VACR',
    'technical and business communication skills':'TBCS','problem solving in engineering':'PSE','integrated project':'IP','ip':'IP',
    'dimensions of engineering':'IP','pse':'PSE','web development with mern stack':'MERN','programming in python':'Python',
    'essentials of python programming':'Python','java programming':'Java','full stack java':'FSJ','embedded systems and iot':'ESIoT',
    'internet of things':'IoT','database management systems':'DBMS','database management system':'DBMS','database systems':'DBS',
    'advanced java programming':'Adv Java','essentials of java':'Java','big data':'Big Data','machine learning':'ML',
    'professional excellence through soft skills':'Soft Skills','english for professionals':'English','english through theatre arts':'English (Theatre)',
    'constitution of india':'COI','technical and business communication':'TBCS','cryptography':'Crypto','facts and hvdc':'FACTS-HVDC',
    'rtos concepts for automotive applications':'RTOS','steel detailing mastery using tekla structures':'Tekla','essentials of java programming':'Java',
    'project and risk management':'PRM','skill integrated language':'SIL','analog vlsi design':'AVLSI','vlsi design':'VLSI','power systems':'PS',
    'switchgear & protection':'SGP','switchgear and protection':'SGP','computer networking':'CN','introduction to natural lanuage processing':'NLP',
    'introduction to natural language processing':'NLP','ai for mechanical engineering':'AI-ME','r programming':'R Prog','data analytics with r':'DA-R','advanced reading comprehension skills':'ARCS','computer aided design':'CAD'}
ABBR_STOP={'and','of','in','for','the','using','with','to','through','its','on','a','an','&'}
def course_abbr(name):
    """Short course name for timetable cells (codes stay in the legend)."""
    raw=re.sub(r'\(.*?(\)|$)','',str(name or '')).strip()
    if sum(c.isupper() for c in raw)>0.6*sum(c.isalpha() for c in raw):raw=raw.title()
    lab=bool(re.search(r'\b(lab|laboratory)\b',raw,re.I))
    base=re.sub(r'\b(lab|laboratory)\b','',raw,flags=re.I).strip(' -–')
    roman=''
    m=re.search(r'[\s-]+(i{1,3}|iv|v)\s*$',base,re.I)
    if m:roman='-'+m.group(1).upper();base=base[:m.start()].strip(' -–')
    key=re.sub(r'\s+',' ',base.lower()).strip()
    if key in ABBR_OVERRIDES:ab=ABBR_OVERRIDES[key]
    else:
        words=[w for w in re.split(r'[\s,/&–-]+',base) if w and w.lower() not in ABBR_STOP]
        if len(words)==1:ab=words[0][:10]
        else:ab=''.join(w if (sum(ch.isupper() for ch in w)>=2 and len(w)<=5) else w[0].upper() for w in words)
    return ab+roman+(' Lab' if lab else '')
def rx_esc(v):return html.escape(str(v if v is not None else ''),quote=True)
def rx_slug(v):return ''.join(c if c.isalnum() else '_' for c in str(v))
def rx_short(t,n=30):
    t=str(t or '')
    return rx_esc(t) if len(t)<=n else f'<span title="{rx_esc(t)}">{rx_esc(t[:n-1])}&hellip;</span>'

def rx_result(model,res,data):
    """Shape the solved timetable the way the rich views expect."""
    days,times=model['days'],model['times'];n=len(times)
    acts={a['id']:a for a in model['activities']}
    fdept=data.get('facdept',{})
    fname={}
    for a in model['activities']:
        nm=[x.strip() for x in (a.get('faculty_names') or '').split(';')]
        for i,f in enumerate(a['faculty']):fname.setdefault(f,nm[i] if len(nm)==len(a['faculty']) else f)
    key=lambda f:'TBC (not yet allocated)' if f.startswith('TBC') else f'{fname.get(f,f)} ({f})' if fname.get(f,f)!=f else f
    secinfo={};tt=defaultdict(lambda:{d:[[] for _ in range(n)] for d in days});subjects={};rooms_used=defaultdict(int)
    for a in model['activities']:
        for g in a['groups']:
            r=g.split(':')[0];m=re.match(r'(.+)-S(\d+)-(.+)$',r)
            if m and r not in secinfo:secinfo[r]={'section':r,'dept':m[1],'sem':int(m[2]),'label':f'{m[1]} {["","I","II","III","IV"][min(4,(int(m[2])+1)//2)]}-{m[3]}'}
    for s in res['schedule']:
        a=acts[s['activity_id']];typ='training' if s['activity_id'] in model.get('fixed_ids',set()) else 'lab' if a['kind']=='Lab' else 'project' if a['kind']=='IP' else 'theory'
        subjects.setdefault(a['code'],{'code':a['code'],'name':a['name']+f" [{a['code']}]",'shortcut':course_abbr(a['name']),'dept':a['dept'],'type':typ})
        tags=sorted({g.split(':',1)[1] for g in a['groups'] if ':' in g})
        slot={'subject':a['code'],'abbr':course_abbr(a['name']),'tag':'/'.join(tags),'type':typ,'facs':[key(f) for f in a['faculty']],'room':s['room'],'merged':len({g.split(':')[0] for g in a['groups']})>1}
        rooms_used[s['room']]+=s['duration']
        for g in sorted({g.split(':')[0] for g in a['groups']}):
            for q in range(s['period']-1,s['period']-1+s['duration']):tt[g][s['day']][q].append(slot)
    for b in model.get('section_busy',[]):
        if b['section'] not in secinfo:continue
        blk={'subject':b.get('label') or 'Fixed commitment','tag':'','type':'training','facs':[],'room':'','merged':False}
        for q,(st,en) in enumerate(times):
            if overlaps(st,en,b['start'],b['end']) and not tt[b['section']][b['day']][q]:tt[b['section']][b['day']][q].append(blk)
    busy=defaultdict(lambda:defaultdict(list))
    for b in model['busy']:
        for q,(st,en) in enumerate(times):
            if overlaps(st,en,b['start'],b['end']):
                for f in b['faculty']:busy[key(f)][(b['day'],q)].append({'activity':b.get('note') or 'Existing load','period':q})
    homes={}
    for sec in secinfo:
        c=Counter(x['room'] for d in days for cell in tt[sec][d] for x in cell if x['type']=='theory' and not x['room'].startswith('TBC'))
        homes[sec]=c.most_common(1)[0][0] if c else ''
    lunch=next((q for q in range(1,n) if times[q-1][1]!=times[q][0]),0)
    faculty={}
    for a in model['activities']:
        for f in a['faculty']:
            k=key(f);faculty.setdefault(k,{'code':k,'dept':'Placeholder' if f.startswith('TBC') else fdept.get(f) or a['dept']})
    for k in busy:faculty.setdefault(k,{'code':k,'dept':'Other'})
    sections=sorted(secinfo.values(),key=lambda s:(s['dept'],s['sem'],s['section']))
    for s in sections:s['homeRoom']=homes.get(s['section'],'')
    return {'days':days,'nPeriods':n,'periodTimes':[f'{hm(a)}-{hm(b)}' for a,b in times],'lunchAfter':lunch,'sections':sections,'timetable':tt,
            'subjects':list(subjects.values()),'faculty':list(faculty.values()),'busy':busy,'roomsUsed':rooms_used,'roomPools':data.get('roompools',{})}

def rx_plabel(R,p):return f"P{p+1} ({R['periodTimes'][p]})"
def rx_lunch(n_cols,R):
    t=R['periodTimes'];b=R['lunchAfter'];lab='LUNCH'
    if 0<b<len(t):lab=f"LUNCH ({t[b-1].split('-')[1]}-{t[b].split('-')[0]})"
    return f'<tr class="lunch-row"><td colspan="{n_cols}">{rx_esc(lab)}</td></tr>'
RX_CODE=re.compile(r'[-_ ]?[A-Z]{2,4}[0-9X]{3,5}[A-Z]?$')
def rx_tag(tag):
    out=[]
    for t in (tag or '').split('/'):
        t=RX_CODE.sub('',t.strip())
        m=re.fullmatch(r'(PE|OE)(I{1,3}|IV|V|VI)',t)
        if m:t=f'{m[1]}-{m[2]}'
        if t and t not in out:out.append(t)
    return '/'.join(out)
def rx_join_labels(labels):
    # 'AIML III-A','AIML III-B' -> 'AIML III-A/B'
    grp={}
    for l in labels:
        pre,_,suf=l.rpartition('-')
        if not pre:pre,suf=l,''
        grp.setdefault(pre,[])
        if suf and suf not in grp[pre]:grp[pre].append(suf)
    return ', '.join(f"{pre}-{'/'.join(s)}" if s else pre for pre,s in grp.items())
def rx_group(pairs):
    # one cell per class: sections taught together are shown once
    order=[];by={}
    for s,x in pairs:
        k=(x.get('subject'),x.get('room'),tuple(x.get('facs') or ()),x.get('tag'))
        if k not in by:by[k]=(x,[]);order.append(k)
        if s['label'] not in by[k][1]:by[k][1].append(s['label'])
    return [(rx_join_labels(by[k][1]),by[k][0]) for k in order]
def rx_cell(slots,prefix=None):
    out=[]
    for x in slots:
        bg,fg=RX_TYPE.get(x['type'],RX_TYPE['theory'])
        tag=rx_tag(x['tag'])
        head=(f'{prefix}: ' if prefix else '')+x.get('abbr',x['subject'])+(f" · {tag}" if tag else '')
        fac=' / '.join(f.split(' (')[0] for f in x['facs'])
        out.append(f'<div class="cell" style="background:{bg};color:{fg}"><div class="ab">{rx_esc(head)}</div>'+(f'<div class="tc">{rx_short(fac,34)}</div>' if fac else '')+(f'<div class="rm">{rx_esc(x["room"])}</div>' if x['room'] else '')+'</div>')
    return '<hr>'.join(out)
def rx_existing(entries):
    return ''.join(f'<div class="existing-load"><span class="existing-x">×</span><div class="existing-title">{rx_esc(e["activity"])}</div><div class="existing-meta">P{e["period"]+1}</div></div>' for e in entries)

def rx_master(R,colors):
    days,n,secs=R['days'],R['nPeriods'],R['sections'];span=Counter(s['dept'] for s in secs);order=list(dict.fromkeys(s['dept'] for s in secs))
    h1=['<th class="lbl" rowspan="2">Day / Period</th>']+[f'<th style="background:{colors[d][0]};color:#fff" colspan="{span[d]}">{rx_esc(d)}</th>' for d in order]
    h2=[f'<th style="background:{colors[s["dept"]][1]};color:{colors[s["dept"]][2]}">{rx_esc(s["label"].split(" ",1)[1])}</th>' for s in secs]
    rows=[]
    for day in days:
        for p in range(n):
            cells=[f'<th class="lbl">{rx_esc(day[:3])} {rx_esc(rx_plabel(R,p))}</th>']+[f'<td>{rx_cell(R["timetable"][s["section"]][day][p])}</td>' for s in secs]
            rows.append(f'<tr class="{"day-end" if p==n-1 else ""}">{"".join(cells)}</tr>')
            if p==R['lunchAfter']-1:rows.append(rx_lunch(1+len(secs),R))
    return '<table><thead><tr>'+''.join(h1)+'</tr><tr>'+''.join(h2)+'</tr></thead><tbody>'+''.join(rows)+'</tbody></table>'

def rx_section(R,colors):
    days,n=R['days'],R['nPeriods'];by=defaultdict(list)
    for s in R['sections']:by[s['dept']].append(s)
    pick=[]
    for d,ss in by.items():
        c=colors[d][0];pick.append(f'<div class="picker-group"><span class="picker-dept" style="color:{c}">{rx_esc(d)}</span>')
        pick+= [f'<button class="btn" style="border-color:{c};color:{c}" data-sec="{rx_esc(s["section"])}" onclick="showSec(\'{rx_esc(s["section"])}\')">{rx_esc(s["label"].split(" ",1)[1])}</button>' for s in ss]
        pick.append('</div>')
    panels=[]
    for s in R['sections']:
        rows=[]
        for p in range(n):
            rows.append(f'<tr><th class="lbl">{rx_esc(rx_plabel(R,p))}</th>'+''.join(f'<td>{rx_cell(R["timetable"][s["section"]][d][p])}</td>' for d in days)+'</tr>')
            if p==R['lunchAfter']-1:rows.append(rx_lunch(1+len(days),R))
        panels.append(f'<div class="sec-panel" id="sec-{rx_slug(s["section"])}" style="display:none"><div class="panel-hdr">{rx_esc(s["label"])} <span>Semester {s["sem"]} · Home room {rx_esc(s["homeRoom"] or "—")}</span></div>'
            f'<div class="wrap"><table><tr><th class="lbl">Period</th>{"".join(f"<th>{rx_esc(d)}</th>" for d in days)}</tr>{"".join(rows)}</table></div></div>')
    return f'<div class="picker">{"".join(pick)}</div>'+''.join(panels)

def rx_teacher(R,colors):
    days,n=R['days'],R['nPeriods'];by=defaultdict(lambda:defaultdict(list));busy=R['busy']
    for s in R['sections']:
        for d in days:
            for p,cell in enumerate(R['timetable'][s['section']][d]):
                for x in cell:
                    for f in x['facs']:
                        if not any(sx is s and xx is x for sx,xx in by[f][(d,p)]):by[f][(d,p)].append((s,x))
    for f in busy:by[f]
    dept={f['code']:f['dept'] for f in R['faculty']}
    groups=defaultdict(list)
    for f in by:groups[dept.get(f,'Other')].append(f)
    for v in groups.values():v.sort()
    def tot(f):
        g=len(by[f]);e=len([k for k in busy.get(f,{}) if k[1]!=R['lunchAfter']-99]);return g,e,g+e
    summ=sorted(((f,dept.get(f,'Other'))+tot(f) for f in by if not f.startswith('TBC')),key=lambda r:-r[4])
    tr=[]
    for f,d,g,e,t in summ:
        c=colors.get(d,("#555",))[0];pct=round(100*t/(len(days)*n));bc="#dc2626" if t>30 else "#d97706" if t>24 else "#16a34a"
        tr.append(f'<tr><td style="text-align:left"><a href="#" onclick="showTab(\'teacher\');showTDept(\'{rx_esc(d)}\');showTeacher(\'{rx_esc(f)}\');return false" style="color:{c};font-weight:600;text-decoration:none">{rx_esc(f)}</a></td><td style="text-align:left;color:{c}">{rx_esc(d)}</td><td>{g}</td><td>{e}</td><td style="font-weight:700">{t}</td><td style="text-align:left"><div class="load-bar"><div style="width:{min(100,pct)}%;background:{bc}"></div></div></td></tr>')
    summary=(f'<div class="sc-legend collapsed" id="load-summary"><div class="sc-legend-hdr" onclick="toggleId(\'load-summary\')">&#128202; Full Faculty Load — all {len(summ)} (click to expand, sorted busiest first)</div>'
        '<div class="wrap"><table class="sc-table"><tr><th>Faculty</th><th>Dept</th><th>This sem</th><th>Existing load</th><th>Total/wk</th><th>Load</th></tr>'+''.join(tr)+'</table></div></div>')
    dcards=[];picks=[]
    for d in sorted(groups):
        c=colors.get(d,("#555",))[0]
        dcards.append(f'<button class="dept-card" style="border-color:{c};color:{c};--dc:{c}" data-tdept="{rx_esc(d)}" onclick="showTDept(\'{rx_esc(d)}\')">{rx_esc(d)}<span class="dept-count">{len(groups[d])} faculty</span></button>')
        picks.append(f'<div class="picker-group tdept-group" id="tdept-{rx_slug(d)}" style="display:none">')
        for f in groups[d]:
            g,e,t=tot(f)
            picks.append(f'<button class="tch-card" style="border-color:{c}" data-tch="{rx_esc(f)}" onclick="showTeacher(\'{rx_esc(f)}\')">{rx_esc(f)}<span class="tch-dept">Dept: {rx_esc(d)}</span><span class="tch-stats">{t} total/wk'+(f' ({g} this sem + {e} existing load)' if e else '')+'</span></button>')
        picks.append('</div>')
    panels=[]
    for f,sk in by.items():
        b=busy.get(f,{});g,e,t=tot(f);rows=[]
        for p in range(n):
            cells=[f'<th class="lbl">{rx_esc(rx_plabel(R,p))}</th>']
            for d in days:
                pairs=sk.get((d,p)) or [];ex=b.get((d,p),[]);parts=[rx_cell([x],lab) for lab,x in rx_group(pairs)]+([rx_existing(ex)] if ex else [])
                cells.append(f'<td{" class=existing-cell" if ex else ""}>{"<hr>".join(parts)}</td>')
            rows.append(f'<tr>{"".join(cells)}</tr>')
            if p==R['lunchAfter']-1:rows.append(rx_lunch(1+len(days),R))
        clash=' <span style="background:#dc2626;color:#fff;padding:2px 8px;border-radius:10px">placeholder</span>' if f.startswith('TBC') else ''
        panels.append(f'<div class="tch-panel" id="tch-{rx_slug(f)}" style="display:none"><div class="panel-hdr">{rx_esc(f)}{clash} <span class="teacher-dept">Dept: {rx_esc(dept.get(f,"Other"))}</span>'
            f'<div class="load-breakdown"><span>This sem: <b>{g}</b></span><span>Existing load: <b>{e}</b></span><span>Total: <b>{t}</b>/{len(days)*n} periods/wk</span></div></div>'
            f'<div class="wrap"><table><tr><th class="lbl">Period</th>{"".join(f"<th>{rx_esc(d)}</th>" for d in days)}</tr>{"".join(rows)}</table></div></div>')
    return summary+f'<div class="tch-grid dept-grid">{"".join(dcards)}</div><div class="tdept-wrap">{"".join(picks)}</div>'+''.join(panels)

def rx_room(R,colors):
    days,n=R['days'],R['nPeriods'];by=defaultdict(lambda:defaultdict(list))
    for s in R['sections']:
        for d in days:
            for p,cell in enumerate(R['timetable'][s['section']][d]):
                for x in cell:
                    if x['room'] and not any(xx is x for _,xx in by[x['room']][(d,p)]):by[x['room']][(d,p)].append((s,x))
    pools=R['roomPools'];placed=set();blocks=[]
    def card(title,rooms,col):
        rs=[r for r in rooms if r in by];placed.update(rs)
        if not rs:return ''
        return (f'<div class="pool"><div class="pool-hdr" style="background:{col}">{rx_esc(title)}<span>{len(rs)} room{"s" if len(rs)>1 else ""}</span></div><div class="pool-body">'
            +''.join(f'<button class="btn" data-room="{rx_esc(r)}" onclick="showRoom(\'{rx_esc(r)}\')">{rx_esc(r)}</button>' for r in rs)+'</div></div>')
    for kind,label,col in [('labs','Laboratory Room Pools','#1a237e'),('classrooms','Theory Room Pools','#b71c1c')]:
        cards=''.join(card(f'{d} {"Labs" if kind=="labs" else "Theory Rooms"}',pools[d][kind],colors.get(d,("#1a237e",))[0]) for d in sorted(pools) if pools[d][kind])
        if cards:blocks.append(f'<div class="pool-section"><div class="pool-section-hdr" style="background:{col}">{label}</div><div class="pool-grid">{cards}</div></div>')
    other=sorted(r for r in by if r not in placed and not r.startswith('TBC'));tbc=sorted(r for r in by if r.startswith('TBC'))
    if other:blocks.append(f'<div class="pool-section"><div class="pool-section-hdr" style="background:#374151">Other rooms used (not in any department list)</div><div class="pool-grid">{card("Other rooms",other,"#374151")}</div></div>')
    if tbc:blocks.append(f'<div class="pool-section"><div class="pool-section-hdr" style="background:#92400e">Provisional rooms (TBC): {len(tbc)} — need a real room</div><div class="pool-grid">{card("Provisional",tbc,"#92400e")}</div></div>')
    panels=[]
    for room,sk in by.items():
        rows=[]
        for p in range(n):
            rows.append(f'<tr><th class="lbl">{rx_esc(rx_plabel(R,p))}</th>'+''.join(f'<td>{"<hr>".join(rx_cell([x],lab) for lab,x in rx_group(sk.get((d,p)) or []))}</td>' for d in days)+'</tr>')
            if p==R['lunchAfter']-1:rows.append(rx_lunch(1+len(days),R))
        panels.append(f'<div class="room-panel" id="room-{rx_slug(room)}" style="display:none"><div class="panel-hdr">Room {rx_esc(room)} <span>{len(sk)} periods/wk occupied</span></div>'
            f'<div class="wrap"><table><tr><th class="lbl">Period</th>{"".join(f"<th>{rx_esc(d)}</th>" for d in days)}</tr>{"".join(rows)}</table></div></div>')
    return f'<div class="picker">{"".join(blocks)}</div>'+''.join(panels)

def rx_course(R):
    days,n=R['days'],R['nPeriods'];by=defaultdict(lambda:defaultdict(list));names={s['code']:s['name'] for s in R['subjects']}
    for s in R['sections']:
        for d in days:
            for p,cell in enumerate(R['timetable'][s['section']][d]):
                for x in cell:by[x['subject']][(d,p)].append((s,x))
    pick=['<div class="picker-group">']+[f'<button class="btn" data-subj="{rx_esc(c)}" onclick="showSubj(\'{rx_esc(c)}\')">{rx_esc(c)} — {rx_esc(names.get(c,c))}</button>' for c in sorted(by)]+['</div>']
    panels=[]
    for c,sk in by.items():
        rows=[]
        for p in range(n):
            cells=[f'<th class="lbl">{rx_esc(rx_plabel(R,p))}</th>']
            for d in days:
                parts=[f'<div class="cell" style="background:#edf3ff;color:#1f2a60"><div class="ab">{rx_esc(lab)}{(" · "+rx_esc(rx_tag(x["tag"]))) if rx_tag(x["tag"]) else ""}</div><div class="tc">{rx_short(" / ".join(f.split(" (")[0] for f in x["facs"]),30)}</div><div class="rm">{rx_esc(x["room"])}</div></div>' for lab,x in rx_group(sk.get((d,p)) or [])]
                cells.append(f'<td>{"<hr>".join(parts)}</td>')
            rows.append(f'<tr>{"".join(cells)}</tr>')
            if p==R['lunchAfter']-1:rows.append(rx_lunch(1+len(days),R))
        nsec=len({s['section'] for v in sk.values() for s,_ in v})
        panels.append(f'<div class="subj-panel" id="subj-{rx_slug(c)}" style="display:none"><div class="panel-hdr">{rx_esc(c)} — {rx_esc(names.get(c,c))} <span>{nsec} sections</span></div>'
            f'<div class="wrap"><table><tr><th class="lbl">Period</th>{"".join(f"<th>{rx_esc(d)}</th>" for d in days)}</tr>{"".join(rows)}</table></div></div>')
    return f'<div class="picker">{"".join(pick)}</div>'+''.join(panels)

def rx_legend(R):
    rows=''.join(f'<tr><td><span class="sc-chip" style="background:{RX_TYPE[s["type"]][0]};color:{RX_TYPE[s["type"]][1]}">{rx_esc(s["shortcut"])}</span></td><td>{rx_esc(s["name"])}</td><td>{rx_esc(s["dept"])}</td><td>{rx_esc(s["type"].capitalize())}</td></tr>'
        for s in sorted(R['subjects'],key=lambda s:(s['dept'],s['code'])))
    return ('<div class="sc-legend collapsed"><div class="sc-legend-hdr" onclick="toggleLegend(this)">&#128218; Course Shortcuts (click to expand)</div>'
        '<div class="wrap"><table class="sc-table"><tr><th>Shortcut</th><th>Course</th><th>Dept</th><th>Type</th></tr>'+rows+'</table></div></div>')

def rx_unplaced(unplaced):
    if not unplaced:return '<div class="meta" style="background:#dcfce7;color:#14532d">Every required period is placed.</div>'
    rows=''.join(f'<tr><td>{rx_esc(u["dept"])}</td><td>{u["semester"]}</td><td style="text-align:left">{rx_esc(u["course"])} [{rx_esc(u["code"])}]</td><td>{rx_esc(u["groups"])}</td><td style="text-align:left">{rx_esc(u["faculty"])}</td><td>{u["periods"]}</td><td style="text-align:left">{rx_esc(u["reason"])}</td></tr>' for u in unplaced)
    return (f'<div class="sc-legend" id="unplaced"><div class="sc-legend-hdr" onclick="toggleId(\'unplaced\')" style="color:#b91c1c">&#9888; Unplaced sessions: {len(unplaced)} ({sum(u["periods"] for u in unplaced)} periods)</div>'
        '<div class="wrap"><table class="sc-table"><tr><th>Dept</th><th>Sem</th><th>Course</th><th>Groups</th><th>Faculty</th><th>Periods</th><th>Reason</th></tr>'+rows+'</table></div></div>')

RX_CSS = """
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:Arial,sans-serif;font-size:11px;background:#f0f4f8;color:#20242d}
h1{text-align:center;padding:14px;font-size:18px;background:#1a237e;color:#fff}
.generated-time{display:block;margin-top:4px;font-size:10px;font-weight:normal;color:#c5cae9;letter-spacing:.2px}
.tabs{display:flex;background:#1a237e;padding:0 16px;gap:4px;flex-wrap:wrap}
.tab{padding:10px 20px;cursor:pointer;color:#c5cae9;font-size:12px;font-weight:bold;border-bottom:3px solid transparent;transition:.15s}
.tab.active{color:#fff;border-bottom-color:#facc15}.tab:hover{color:#fff}
.view{display:none}.view.active{display:block}
.wrap{overflow:auto;padding:10px}
table{border-collapse:collapse;white-space:nowrap;width:100%}
th,td{border:1px solid #dfe3ef;padding:3px 5px;text-align:center;min-width:70px;vertical-align:top}
.lbl{background:#1a237e;color:#fff;font-weight:bold;position:sticky;left:0;z-index:5;font-size:9.5px}
.day-end td,.day-end th{border-bottom:3px solid #374151!important}
.ab{font-weight:700;font-size:10.5px;line-height:1.2}.tc{font-size:8.5px;opacity:.85;line-height:1.2}.rm{font-size:7.5px;opacity:.65;line-height:1.1;margin-top:1px}
.cell{border-radius:4px;padding:3px 4px}
td hr{border:0;border-top:1px dashed #cbd2e1;margin:2px 0}
.picker{padding:10px 14px;background:#fff;border-bottom:1px solid #dfe3ef}
.picker-group{margin-bottom:8px;display:flex;flex-wrap:wrap;gap:6px;align-items:center}
.picker-dept{font-weight:700;font-size:10.5px;text-transform:uppercase;margin-right:6px;min-width:110px}
.btn{padding:5px 11px;border:1.5px solid #1a237e;border-radius:14px;cursor:pointer;font-size:10px;background:#fff;color:#1a237e;transition:.15s}
.btn.sel{color:#fff!important;background:#1a237e!important}
.tch-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(170px,1fr));gap:8px;padding:12px 14px;background:#fff;border-bottom:1px solid #dfe3ef}
.existing-cell{background:#fafbfc!important;color:#9aa3af;text-align:left!important}
.existing-load{position:relative;padding:3px 14px 3px 4px;border-left:2px solid #d7dce2;background:#f3f5f7;border-radius:3px;margin:2px 0;line-height:1.25}
.existing-x{position:absolute;right:4px;top:2px;font-size:11px;font-weight:600;color:#b9c1ca;line-height:1}
.existing-title{font-size:8px;font-weight:600;color:#7c8794}.existing-meta{font-size:7px;color:#9aa3af;margin-top:1px}
.lunch-row td{background:#fff8e1;color:#7a1f3d;font-weight:700;font-size:10px;text-align:center;letter-spacing:1.5px;padding:3px 0;border-top:1.5px solid #f0c987;border-bottom:1.5px solid #f0c987}
.load-breakdown{font-size:10px;color:#c5cae9;display:flex;gap:10px}.load-breakdown b{color:#fff}
.tch-card{padding:8px 10px;border-radius:6px;border:1.5px solid #374151;cursor:pointer;font-size:11px;font-weight:600;background:#fff;color:#374151;transition:.15s;text-align:center;line-height:1.3;display:block}
.tch-card:hover,.tch-card.sel{background:#1a237e;color:#fff;border-color:#1a237e}
.tch-dept{font-size:9px;color:#334155;font-weight:600;display:block;margin-top:3px}.tch-stats{font-size:9px;color:#6b7280;font-weight:normal;display:block;margin-top:2px}
.tch-card:hover .tch-dept,.tch-card.sel .tch-dept,.tch-card:hover .tch-stats,.tch-card.sel .tch-stats{color:#c5cae9}
.panel-hdr{background:#1a237e;color:#fff;padding:8px 12px;font-size:12.5px;font-weight:bold;display:flex;gap:16px;align-items:center;flex-wrap:wrap}
.panel-hdr span{font-size:10px;font-weight:normal;color:#c5cae9}
.panel-hdr .teacher-dept{font-size:11px;font-weight:700;color:#fff;background:#3949ab;padding:3px 8px;border-radius:10px}
.legend{display:flex;gap:14px;justify-content:center;padding:7px;flex-wrap:wrap;background:#fff;border-bottom:1px solid #dfe3ef}
.leg{display:flex;align-items:center;gap:5px;font-size:10px;padding:2px 8px;border-radius:10px;background:#f4f6fb}
.legbox{width:11px;height:11px;border-radius:50%;border:1px solid rgba(0,0,0,.15)}
.meta{text-align:center;padding:8px 6px;font-size:10.5px;background:#eef1f8;border-bottom:1px solid #dfe3ef;color:#374151;font-weight:600;letter-spacing:.2px}
.sc-legend{background:#eef1f8;border-bottom:2px solid #dfe3ef}
.sc-legend-hdr{text-align:center;font-size:11px;font-weight:bold;color:#1a237e;padding:6px 0;cursor:pointer;user-select:none}
.sc-legend.collapsed .wrap{display:none}
.sc-table{border-collapse:collapse;font-size:9.5px;margin:4px auto 8px;width:auto}
.sc-table th{background:#1a237e;color:#fff;padding:4px 8px;text-align:left;position:sticky;top:0}
.sc-table td{padding:3px 8px;border-bottom:1px solid #e5e7eb}
.sc-chip{display:inline-block;padding:2px 7px;border-radius:4px;font-weight:bold;font-size:9.5px;border:1px solid rgba(0,0,0,.08)}
.dept-grid{grid-template-columns:repeat(auto-fill,minmax(140px,1fr))}
.dept-card{padding:10px 12px;border-radius:8px;border:2px solid #374151;cursor:pointer;font-size:12px;font-weight:700;background:#fff;transition:.15s;text-align:center}
.dept-card:hover,.dept-card.sel{background:var(--dc,#1a237e);color:#fff!important}
.dept-card.sel .dept-count{color:#fff}.dept-count{display:block;font-size:9.5px;font-weight:normal;color:#6b7280;margin-top:3px}
.tdept-wrap{padding:12px 14px;background:#fff;border-bottom:1px solid #dfe3ef;max-height:260px;overflow-y:auto}
.tdept-group{display:grid;grid-template-columns:repeat(auto-fill,minmax(190px,1fr));gap:8px}
.load-bar{width:90px;height:8px;background:#e5e7eb;border-radius:4px;overflow:hidden}.load-bar div{height:100%;border-radius:4px}
#load-summary .wrap,#unplaced .wrap{max-height:420px;overflow-y:auto}
.pool-section{border:2px solid #1a237e;border-radius:8px;margin-bottom:12px;overflow:hidden;background:#eef1f8}
.pool-section-hdr{color:#fff;font-weight:700;font-size:13px;padding:8px 14px}
.pool-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(250px,1fr));gap:10px;padding:10px}
.pool{background:#fff;border:1px solid #dfe3ef;border-radius:6px;overflow:hidden}
.pool-hdr{color:#fff;font-weight:700;font-size:11px;padding:6px 10px;display:flex;justify-content:space-between}
.pool-hdr span{font-weight:normal;font-size:9.5px;opacity:.85}
.pool-body{display:flex;flex-wrap:wrap;gap:6px;padding:10px}
"""
RX_JS = """
function showTab(id){document.querySelectorAll('.tab').forEach(t=>t.classList.remove('active'));document.querySelectorAll('.view').forEach(v=>v.classList.remove('active'));
 document.getElementById('tab-'+id).classList.add('active');document.getElementById('view-'+id).classList.add('active');window.scrollTo(0,0);}
function pick(attr,val,cls,prefix){document.querySelectorAll('['+attr+']').forEach(b=>b.classList.remove('sel'));document.querySelectorAll('['+attr+'="'+val+'"]').forEach(b=>b.classList.add('sel'));
 document.querySelectorAll('.'+cls).forEach(p=>p.style.display='none');var el=document.getElementById(prefix+val.replace(/[^a-zA-Z0-9]/g,'_'));if(el){el.style.display='block';el.scrollIntoView({behavior:'smooth',block:'start'});}}
function showSec(s){pick('data-sec',s,'sec-panel','sec-')}
function showRoom(r){pick('data-room',r,'room-panel','room-')}
function showSubj(s){pick('data-subj',s,'subj-panel','subj-')}
function showTeacher(t){pick('data-tch',t,'tch-panel','tch-')}
function showTDept(d){document.querySelectorAll('.dept-card').forEach(b=>b.classList.remove('sel'));document.querySelectorAll('.dept-card[data-tdept="'+d+'"]').forEach(b=>b.classList.add('sel'));
 document.querySelectorAll('.tdept-group').forEach(g=>g.style.display='none');document.querySelectorAll('.tch-panel').forEach(p=>p.style.display='none');
 var el=document.getElementById('tdept-'+d.replace(/[^a-zA-Z0-9]/g,'_'));if(el){el.style.display='grid';}}
function toggleLegend(el){el.parentElement.classList.toggle('collapsed')}
function toggleId(id){document.getElementById(id).classList.toggle('collapsed')}
window.onload=function(){showTab('master');var f=document.querySelector('[data-sec]');if(f)showSec(f.getAttribute('data-sec'));window.scrollTo(0,0);};
"""

def write_rich_html(model,res,data,unplaced,path,meta):
    R=rx_result(model,res,data)
    depts=list(dict.fromkeys([s['dept'] for s in R['sections']]+sorted({f['dept'] for f in R['faculty']})))
    colors={d:RX_DEPT_COLORS[i%len(RX_DEPT_COLORS)] for i,d in enumerate(depts)}
    nfac=len([f for f in R['faculty'] if not f['code'].startswith('TBC')])
    page=f"""<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>{rx_esc(meta['title'])} — {rx_esc(meta['generated'])}</title><style>{RX_CSS}</style></head><body>
<h1>&#128197; {rx_esc(meta['title'])}<span class="generated-time">Generated: {rx_esc(meta['generated'])} · {rx_esc(res['status'])} · {res['placed_periods']} of {res['required_periods']} periods placed</span></h1>
<div class="meta">{len(R['sections'])} sections &middot; {nfac} faculty &middot; {len(R['days'])}-day week &middot; {len(R['roomsUsed'])} rooms used</div>
<div class="legend"><div class="leg"><span class="legbox" style="background:#edf3ff"></span> Theory</div><div class="leg"><span class="legbox" style="background:#fff0f5"></span> Lab</div><div class="leg"><span class="legbox" style="background:#eef9ee"></span> IP / PSE</div><div class="leg">B1 / B2 = lab batch · pooled electives share a cell</div></div>
{rx_unplaced(unplaced)}{rx_legend(R)}
<div class="tabs"><div class="tab" id="tab-master" onclick="showTab('master')">&#128203; Master View</div><div class="tab" id="tab-section" onclick="showTab('section')">&#127979; Section View</div>
<div class="tab" id="tab-teacher" onclick="showTab('teacher')">&#128100; Teacher View</div><div class="tab" id="tab-room" onclick="showTab('room')">&#127970; Room View</div><div class="tab" id="tab-course" onclick="showTab('course')">&#128218; Course View</div></div>
<div class="view" id="view-master"><div class="wrap">{rx_master(R,colors)}</div></div>
<div class="view" id="view-section">{rx_section(R,colors)}</div>
<div class="view" id="view-teacher">{rx_teacher(R,colors)}</div>
<div class="view" id="view-room">{rx_room(R,colors)}</div>
<div class="view" id="view-course">{rx_course(R)}</div>
<script>{RX_JS}</script></body></html>"""
    path.write_text(page,encoding='utf-8')

# ----------------------------------------------------------------- 6. main
def main():
    p=argparse.ArgumentParser(description='Solve the timetable from one input workbook and export Excel + HTML.')
    p.add_argument('--input',type=Path,default=HERE.parent/'output'/'workload_all_departments_R24_nov20_v54.xlsx')
    p.add_argument('--seconds',type=int,default=None,help='search time (default from Settings, else 600)')
    p.add_argument('--workers',type=int,default=None)
    p.add_argument('--out',type=Path,default=None,help='output folder (default: next to the input)')
    p.add_argument('--name',default='timetable_nov20_v54',help='output file name without extension')
    p.add_argument('--hint',type=Path,default=None,help='previous output .xlsx for warm start (default: <out>/<name>.xlsx if present)')
    p.add_argument('--free-depts',default=None,help='comma list (e.g. CS,IT,CIVIL): only these departments are re-solved; all others are fixed to the --hint timetable')
    p.add_argument('--free-faculty',default=None,help='comma list of teacher IDs whose classes in frozen departments may also move (needs --free-depts)')
    p.add_argument('--stop-at-first',action='store_true',help='stop at the first complete timetable (feasibility test)')
    p.add_argument('--check-only',action='store_true',help='validate the input and stop')
    p.add_argument('--soft-rooms',type=int,default=0,metavar='WEIGHT',help='allow a TBC room even where rooms are listed, at this penalty per class (0 = room lists are hard)')
    p.add_argument('--redraw',type=Path,default=None,metavar='XLSX',help='no solving: rebuild the outputs from a generated timetable Excel')
    g=p.add_mutually_exclusive_group()
    g.add_argument('--excel',action='store_true',help='solve, write the Excel output only')
    g.add_argument('--html',action='store_true',help='solve, write the HTML output only')
    g.add_argument('--html-from-excel',nargs='?',const='',type=str,metavar='XLSX',help='no solving: build the HTML from a generated Excel (default <out>/<name>.xlsx)')
    a=p.parse_args()
    global SOFT_ROOM_WEIGHT
    SOFT_ROOM_WEIGHT=a.soft_rooms
    data=read_input(a.input);st=data['settings']
    print(f'Input OK: {len(data["activities"])} activities, {sum(x["hours"] for x in data["activities"])} periods, {len(data["busy"])} busy blocks')
    print(f'Constraints: {sum(rule["enabled"] for rule in data["constraints"])} enabled; existing unknown components count as theory for the daily cap')
    warnings = existing_cap_warnings(data)
    for warning in warnings:print('EXISTING LOAD:', warning + ('; no additional teaching allowed that day' if 'PER_DAY' in warning else ''))
    blocking_warnings = [warning for warning in warnings if 'MAX_FACULTY_PERIODS_PER_WEEK' in warning]
    if a.check_only:
        model, _ = build_cp_model(data)
        error = model.validate()
        if error:raise InputError('Constraint model is invalid: ' + error)
        print('Constraint settings and CP-SAT model validated; no search run.')
        if blocking_warnings:raise InputError('Existing commitments exceed the configured hard weekly cap.')
        return
    if blocking_warnings and a.html_from_excel is None:raise InputError('Existing commitments exceed the configured hard weekly cap. Run --check-only for details.')
    out=a.out or a.input.parent;out.mkdir(parents=True,exist_ok=True)
    hint=a.hint or next((h for h in [out/(a.name+'.xlsx'),out/'timetable_output.xlsx'] if h.exists()),out/(a.name+'.xlsx'))
    hints=read_hint(hint) if hint.exists() else []
    data['stop_at_first']=a.stop_at_first
    if a.free_depts:
        if not hints:raise InputError('--free-depts needs --hint with a previous timetable to freeze the other departments')
        data['free_depts']={x.strip() for x in a.free_depts.split(',') if x.strip()}
        known={x['dept'] for x in data['activities']};bad=data['free_depts']-known
        if bad:raise InputError(f'Unknown departments in --free-depts: {sorted(bad)}; known: {sorted(known)}')
        data['free_faculty']={x.strip() for x in (a.free_faculty or '').split(',') if x.strip()}
    secs=a.seconds or int(st.get('Seconds') or 600);workers=a.workers or int(st.get('Workers') or 4)
    if a.html_from_excel is not None:
        src=Path(a.html_from_excel) if a.html_from_excel else out/(a.name+'.xlsx')
        hints=read_hint(src) if src.exists() else []
        if not hints:raise InputError(f'--html-from-excel needs a generated timetable Excel with a Sessions sheet: {src}')
        res=from_hints(data,hints)
    elif a.redraw is not None:
        hints=read_hint(a.redraw) if a.redraw.exists() else []
        if not hints:raise InputError(f'--redraw needs a generated timetable Excel with a Sessions sheet: {a.redraw}')
        res=from_hints(data,hints)
    model={'days':data['days'],'times':data['times'],'activities':data['activities'],'busy':data['busy'],
           'not_scheduled':data['not_scheduled'],'clarifications':data['clarifications'],'section_busy':data.get('section_busy',[]),'fixed_ids':{x['activity_id'] for x in data.get('fixed_sessions',[])}}
    def export(res,name,quiet=False):
        meta={'title':str(st.get('Title') or 'Timetable'),'generated':datetime.datetime.now().strftime('%d %b %Y %H:%M'),
              'basis':str(st.get('Basis') or a.input.name),'scope':str(st.get('Scope') or ''),'notes':data['notes']}
        acts,fname,sessions,unplaced=build(model,res)
        tmp=out/(name+'.xlsx');saved=[]
        if not (a.html or a.html_from_excel is not None):
            try:write_excel(model,res,acts,fname,sessions,unplaced,tmp,meta)
            except PermissionError:
                tmp=out/(name+'_'+datetime.datetime.now().strftime('%H%M')+'.xlsx');print('Output open in Excel; saving as',tmp.name)
                write_excel(model,res,acts,fname,sessions,unplaced,tmp,meta)
            saved.append(tmp)
        if not a.excel:
            write_rich_html(model,res,data,unplaced,tmp.with_suffix('.html'),meta);saved.append(tmp.with_suffix('.html'))
        if not quiet:print('Saved:',' and '.join(map(str,saved)),flush=True)
        return unplaced
    saved_during=[]
    def on_save(res,first):
        saved_during.append(1)
        if first:export(res,a.name+'_first',quiet=True)
        export(res,a.name,quiet=True)
    if a.html_from_excel is None and a.redraw is None:
        res=solve(data,secs,workers,hints,on_save=on_save)
        if res['status'] not in ('FEASIBLE','OPTIMAL') and saved_during:
            print(json.dumps({k:res[k] for k in ['status','required_periods','placed_periods','unplaced_periods']}),f'{res["seconds"]:.0f}s',flush=True)
            print('Search ended without a final result; keeping the timetable saved during the search:',out/(a.name+'.xlsx'))
            return
    print(json.dumps({k:res[k] for k in ['status','required_periods','placed_periods','unplaced_periods']}),f'{res["seconds"]:.0f}s',flush=True)
    export(res,a.name)
    if res['unplaced']:print(f"  {len(res['unplaced'])} unplaced sessions (listed in the Unplaced sheet of the output)")

if __name__=='__main__':
    try:main()
    except InputError as e:print(e);sys.exit(2)
