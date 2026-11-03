from pathlib import Path
import json
root=Path(__file__).resolve().parents[1]
folder=root/'examples/default_tnos';folder.mkdir(exist_ok=True)
voltages=[1.20,1.035,.99,.90,.81,.765,.72,.68,.64,.60,.56]
groups={'W-Low pBL':list(range(350,361)),'R-Low pBL':list(range(370,381))}
expected={}
for phase,step in [(1,2),(2,30),(3,100)]:
    raw=f'LOT=FAX7680T0 WNO=07 STEP=WTT{step} Program=garbledVER_FA\n'
    data={}
    for name,tnos in groups.items():
        for i,tno in enumerate(tnos):
            values=[0]*256
            values[0]=[0]*8+[1,2,2]
            values[0]=values[0][i]
            values[1]=int(i==(phase-1 if name.startswith('W-') else phase))
            values[255]=phase;data[tno]=values
            raw+=f'Tno={tno} X=8 Y=12 FBC BY BLK='+' '.join(map(str,values))+'\n'
            for block,base in [(0,100),(1,200),(255,0)]:
                for j in range(values[block]):raw+=f'[2,{tno},8,12,{block},{base+j},0,0,]\n'
    (folder/f'WT{phase}_STEP{step}.asc').write_text(raw,encoding='ascii')
    expected[f'WT{phase}']={}
    for name,tnos in groups.items():
        result={}
        for threshold in (0,5):
            row=[]
            for block in range(256):
                total=0;value='Pass'
                for i,tno in enumerate(tnos):
                    fbc=data[tno][block];total+=fbc
                    if (fbc!=0 if threshold==0 else total>=threshold):value='Fail' if i==0 else voltages[i-1];break
                row.append(value)
            result['Legacy' if threshold==0 else 'Threshold5']=row
        expected[f'WT{phase}'][name]=result
(folder/'expected.json').write_text(json.dumps(expected,indent=2),encoding='utf-8')
print('Generated three synthetic ASC fixtures for the 22 configured TNOs and independent Vmin expectations.')
