from pathlib import Path
import csv,re
p=Path('C:/Users/dora9/outputs/ASC_ThreeASC_Project');cfg=list(csv.DictReader((p/'Sample_Monitor_Map_64ea/Config_TNO.csv').open(encoding='utf-8-sig')))
func=[r for r in cfg if r['Group']=='Function'];volts={r['TNO']:float(r['Test Voltage']) for r in cfg if r['Group']=='Monitor'}
raw={}
for wt in range(1,4):
 for line in (p/f'Sample_Monitor_Map_64ea/SIMMAP_WT{wt}.asc').read_text().splitlines():
  m=re.search(r'Tno=\s*(\d+).*?X=\s*(-?\d+).*?Y=\s*(-?\d+).*?FBC BY BLK=\s*(.*)',line,re.I)
  if m:
   values=tuple(map(int,m[4].split()));assert len(values)==256
   raw[(f'WT{wt}',m[1],int(m[2]),int(m[3]))]=values
patterns={};hard=[];previous=0
for i,r in enumerate(func):
 if r['Stop Sign']:
  if r['Stop Sign']=='Hard Fail':hard.append(r['TNO'])
  else:patterns[r['Stop Sign']]=(r['TNO'],func[previous:i+1],list(patterns))
  previous=i+1
obs={}
for r in csv.DictReader((p/'Sample_Monitor_Map_64ea/Expected_Monitor_Details.csv').open(encoding='utf-8-sig')):
 key=(r['WT'],r['Pattern'],int(r['X']),int(r['Y']),int(r['BLK']));a=obs.setdefault(key,[0,0,0,0,0])
 w=[volts[t] for t in r['Write Monitor TNOs'].split(',') if t];rd=[volts[t] for t in r['Read Monitor TNOs'].split(',') if t]
 c=3 if not w or not rd else (2 if abs(max(w)-max(rd))<1e-7 else (0 if max(w)>max(rd) else 1));a[c]+=1;a[4]+=1
contexts=[('WT1','Scan',10,7),('WT1','March',10,7),('WT2','Scan',10,7),('WT3','March',10,7),('WT1','Scan',11,8),('WT2','Scan',12,8),('WT1','March',13,8),('WT1','Scan',14,8),('WT1','Scan',9,8)]
def state(wt,pat,x,y,b):
 end,scope,earlier=patterns[pat]
 if (wt,end,x,y) not in raw:return 'No Data'
 for ph in range(1,int(wt[-1])+1):
  for t in hard:
   k=(f'WT{ph}',t,x,y)
   if k not in raw:return 'No Data'
   if raw[k][b]:return f'Excluded: Hard Fail WT{ph}'
  if ph<int(wt[-1]):
   for r in func:
    k=(f'WT{ph}',r['TNO'],x,y)
    if k not in raw:return 'No Data'
    if raw[k][b]:return f"Excluded: WT{ph} TNO{r['TNO']}"
 for ep in earlier:
  t=patterns[ep][0];k=(wt,t,x,y)
  if k not in raw:return 'No Data'
  if raw[k][b]:return f'Excluded: earlier Pattern TNO{t}'
 return 'Fail' if raw[(wt,end,x,y)][b] else 'Pass'
rows=[]
for wt,pat,x,y in contexts:
 for b in range(256):
  st=state(wt,pat,x,y,b);tn='';v='';missing=False
  if st=='Fail':
   for r in patterns[pat][1]:
    k=(wt,r['TNO'],x,y)
    if k not in raw:missing=True
    elif raw[k][b]:
     tn=r['TNO'];v='' if missing else r['Test Voltage'];break
  a=obs.get((wt,pat,x,y,b),[0]*5);ks=[i for i in range(3) if a[i]];tend='Insufficient' if not ks else ('Mixed' if len(ks)>1 else ['Write worse','Read worse','Equal'][ks[0]])
  rows.append([wt,pat,x,y,b,st,tn,v,tend,*a])
with (p/'wafer128/Expected_BLK_Voltage_Grid.csv').open('w',newline='') as f:
 w=csv.writer(f);w.writerow(['WT','Pattern','X','Y','BLK','State','FirstTNO','FirstVoltage','Tendency','W','R','Equal','Unknown','Captured']);w.writerows(rows)
print(f'Independent ASC/config expectations: {len(raw)} FBC records;{len(rows)} BLK states/voltages/WR counts across{len(contexts)} Chip contexts.')
