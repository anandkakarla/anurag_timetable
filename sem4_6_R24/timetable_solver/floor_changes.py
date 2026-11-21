import sys,openpyxl,re
from collections import defaultdict
def floor(room):
    m=re.match(r'^([A-Z])-(G|\d)',str(room or '').upper())
    if not m: return None
    num=re.match(r'^[A-Z]-(\w+)',str(room).upper()).group(1)
    return m.group(1)+':'+('0' if m.group(2)=='G' or len(num)<3 else m.group(2))
for f in sys.argv[1:]:
    ws=openpyxl.load_workbook(f,read_only=True)['Sessions']; rows=list(ws.iter_rows(values_only=True)); h=rows[0]; I={k:i for i,k in enumerate(h)}
    occ={'fac':defaultdict(dict),'sec':defaultdict(dict)}
    for r in rows[1:]:
        for k in range(r[I['Duration']]):
            p=r[I['Period']]+k; key=(r[I['Day']],p)
            for fid in str(r[I['Faculty IDs']] or '').split(', '):
                if fid and not fid.startswith('TBC'): occ['fac'][fid][key]=r[I['Room']]
            for g in str(r[I['Student groups']]).split(', '): occ['sec'][g][key]=r[I['Room']]
    out=[]
    for kind in ('fac','sec'):
        n=0
        for who,d in occ[kind].items():
            for (day,p),room in d.items():
                if p in (4,7): continue  # P4->P5 is lunch; P7 last
                nxt=d.get((day,p+1))
                if nxt is None: continue
                a,b=floor(room),floor(nxt)
                if a and b and a[0]==b[0] and a!=b: n+=1
        out.append(f'{kind} floor changes {n}')
    print(f.split('/')[-1],*out)
