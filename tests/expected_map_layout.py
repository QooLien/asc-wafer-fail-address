from pathlib import Path
import csv,re,collections
p=Path('C:/Users/dora9/outputs/ASC_ThreeASC_Project');d=p/'Sample_Monitor_Map_64ea';fbc={};cats={}
for wt in range(1,4):
 for line in (d/f'SIMMAP_WT{wt}.asc').read_text().splitlines():
  if wt==1:
   m=re.search(r'TNO=\s*1000\b.*?X=\s*(-?\d+).*?Y=\s*(-?\d+).*?CATLST:\s*([01]+)',line,re.I)
   if m:cats[(int(m[1]),int(m[2]))]=m[3].find('1')+1
  m=re.search(r'Tno=\s*(\d+).*?X=\s*(-?\d+).*?Y=\s*(-?\d+).*?FBC BY BLK=\s*(.*)',line,re.I)
  if m:
   values=list(map(int,m[4].split()));assert len(values)==256;fbc[(f'WT{wt}',m[1],int(m[2]),int(m[3]))]=values
observed=collections.Counter()
for r in csv.DictReader((d/'Expected_Monitor_Details.csv').open(encoding='utf-8-sig')):observed[(r['WT'],r['Pattern'],int(r['X']),int(r['Y']))]+=1
stops={r['Stop Sign']:r['TNO'] for r in csv.DictReader((d/'Config_TNO.csv').open(encoding='utf-8-sig')) if r['Group']=='Function' and r['Stop Sign'] and r['Stop Sign']!='Hard Fail'}
rows=[]
for r in csv.DictReader((d/'Expected_Fail_BLK_Map.csv').open(encoding='utf-8-sig')):
 key=(r['WT'],r['Pattern'],int(r['X']),int(r['Y']));a=fbc.get((r['WT'],stops[r['Pattern']],int(r['X']),int(r['Y'])));cat=cats.get((int(r['X']),int(r['Y'])),0)
 blue=1<=cat<=4 and a is not None and all(v==0 for v in a) and observed[key]==0
 rows.append([*key,cat,int(blue)])
with (p/'wafer128/Expected_Map_Layout_Blue.csv').open('w',newline='') as f:
 w=csv.writer(f);w.writerow(['WT','Pattern','X','Y','Cat','Blue']);w.writerows(rows)
print('Independent Raw ASC/CATLST expectation:',len(rows),'Chip/WT/Pattern states;',sum(r[-1] for r in rows),'blue.')
