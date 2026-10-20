#!/usr/bin/env python3
"""Portal Uploads Report de un mes con CORTE a fin de mes (English).
Mismas métricas que el Tesorio report; breakdown por portal en vez de prioridad.
Uso: python3 generar_portal_corte.py <workflow.xlsx> <month YYYY-MM> [output.xlsx]
"""
import sys
from datetime import datetime
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference
from openpyxl.workbook.properties import CalcProperties

SRC=sys.argv[1]; TARGET=sys.argv[2]
OUT=sys.argv[3] if len(sys.argv)>3 else "/home/user/Finanzas/Portal_Uploads_Report_Sep_2026.xlsx"
y,m=map(int,TARGET.split('-'))
CUT=(pd.Timestamp(year=y,month=m,day=1)+pd.offsets.MonthEnd(1)).normalize()+pd.Timedelta(hours=23,minutes=59,seconds=59)
DCUT=CUT+pd.Timedelta(days=2)  # ventana SLA 2 dias para resolver
def mlabel(s): return datetime.strptime(s,'%Y-%m').strftime('%b - %y')
last3=[(pd.Timestamp(year=y,month=m,day=1)-pd.offsets.MonthBegin(k)).strftime('%Y-%m') for k in (2,1,0)]

def norm(x):
    s=str(x).strip().lower()
    for k,v in [('ariba','Ariba'),('coupa','Coupa'),('chorus','Chorus Pro'),('msft','Microsoft'),('microsoft','Microsoft'),
                ('taulia','Taulia'),('transcepta','Transcepta'),('google','Google'),('oracle','Oracle'),('ivalua','Ivalua'),
                ('vendors@gov','Vendors@Gov'),('tungsten','Tungsten')]:
        if k in s: return v
    if s.startswith('http'): return 'URL'
    if s in ('nan','','test','testing'): return 'Other'
    return str(x).strip().title()[:16]

df=pd.read_excel(SRC,sheet_name='Form Responses')
df['Timestamp']=pd.to_datetime(df['Timestamp'],errors='coerce')
df['Message Timestamp']=pd.to_datetime(df['Message Timestamp'],errors='coerce')
df=df[df['Timestamp'].notna() & (df['Timestamp']<=CUT)].copy()
df['portal']=df['Portal Name or Link'].apply(norm)
df['hasres']=df['Message Timestamp'].notna() & (df['Message Timestamp']<=DCUT)
df['st']=df['hasres'].map({True:'RESOLVED',False:'OPEN'})
df['mes']=df['Timestamp'].dt.strftime('%Y-%m')
# portales a nombrar en la tabla (>=2 en el mes objetivo)
vc=df[df['mes']==TARGET]['portal'].value_counts()
NAMED=[p for p,n in vc.items() if n>=2][:8]
sepall=df[df['mes']==TARGET]
otros=sepall[~sepall['portal'].isin(NAMED)]
otros_res=otros[otros['hasres']]
otros_d=(otros_res['Message Timestamp']-otros_res['Timestamp']).dt.total_seconds()/86400

hfill=PatternFill("solid",fgColor="1F3864"); sfill=PatternFill("solid",fgColor="2E5496")
hfont=Font(name="Arial",bold=True,color="FFFFFF",size=11); sfont=hfont
cfont=Font(name="Arial",size=10); bfont=Font(name="Arial",bold=True,size=10); kfont=Font(name="Arial",bold=True,size=14,color="1F3864")
center=Alignment(horizontal="center",vertical="center"); left=Alignment(horizontal="left",vertical="center")
thin=Side(style="thin",color="D9D9D9"); border=Border(left=thin,right=thin,top=thin,bottom=thin)
red=PatternFill("solid",fgColor="F4CCCC"); yel=PatternFill("solid",fgColor="FFF2CC"); DFMT='dd-mmm-yy'

wb=Workbook()
dz=wb.active; dz.title="Data (calc)"
dcols=["Invoice Number","Status","Portal (norm.)","Submitted By","Collector assigned","Submit Timestamp","Resolved Timestamp","Resolution (hrs)","Resolution (days)","Month","Days open"]
for c,h in enumerate(dcols,1):
    cell=dz.cell(1,c,h); cell.fill=hfill; cell.font=hfont; cell.alignment=center; cell.border=border
r=2
for _,row in df.iterrows():
    inv=str(row['Invoice Number']).replace('\n',' / ') if pd.notna(row['Invoice Number']) else ""
    dz.cell(r,1,inv); dz.cell(r,2,row['st']); dz.cell(r,3,row['portal'])
    dz.cell(r,4,row['Submitted By'] if pd.notna(row['Submitted By']) else "")
    dz.cell(r,5,row['Collector assigned'] if pd.notna(row['Collector assigned']) else "")
    ts=dz.cell(r,6,row['Timestamp'].to_pydatetime()); ts.number_format=DFMT
    if row['hasres']:
        me=dz.cell(r,7,row['Message Timestamp'].to_pydatetime()); me.number_format=DFMT
        h=(row['Message Timestamp']-row['Timestamp']).total_seconds()/3600; dz.cell(r,8,round(h,1)); dz.cell(r,9,round(h/24,1))
    dz.cell(r,8).number_format='0.0'; dz.cell(r,9).number_format='0.0'
    dz.cell(r,10,row['mes'])
    if not row['hasres']:
        dz.cell(r,11, max((DCUT.date()-row['Timestamp'].date()).days,0)); dz.cell(r,11).number_format='0'
    for c in range(1,12):
        cc=dz.cell(r,c); cc.font=cfont; cc.border=border; cc.alignment=left if c in (1,2,3,4,5) else center
    r+=1
LAST=r-1
for i,w in enumerate([20,10,16,18,16,16,16,14,14,10,10],1): dz.column_dimensions[get_column_letter(i)].width=w
dz.freeze_panes="A2"
D="'Data (calc)'"; ST=f"{D}!$B$2:$B${LAST}"; PRT=f"{D}!$C$2:$C${LAST}"
HRS=f"{D}!$H$2:$H${LAST}"; DAYS=f"{D}!$I$2:$I${LAST}"; MES=f"{D}!$J$2:$J${LAST}"

ws=wb.create_sheet("Portal Uploads Report",0); ws.sheet_view.showGridLines=False
TL=mlabel(TARGET)
def title(cell,txt): ws[cell]=txt; ws[cell].font=sfont; ws[cell].fill=sfill; ws[cell].alignment=left
def hdr(row,cols,sc=1):
    for i,h in enumerate(cols):
        cell=ws.cell(row,sc+i,h); cell.fill=hfill; cell.font=hfont; cell.alignment=center; cell.border=border
def box(row,col,val,fmt='General',font=cfont,al=center):
    cell=ws.cell(row,col,val); cell.font=font; cell.alignment=al; cell.number_format=fmt; cell.border=border; return cell
ws['A1']=f"PORTAL UPLOADS REPORT — {TL}"; ws['A1'].font=Font(name="Arial",bold=True,size=16,color="1F3864")
title('A4',f"  {TL} SUMMARY")
kpis=[("Uploads in month",f'=COUNTIF({MES},"{TARGET}")','0'),
 ("Resolved",f'=COUNTIFS({MES},"{TARGET}",{HRS},">=0")','0'),
 ("Avg. resolution (days)",f'=IFERROR(ROUND(AVERAGEIF({MES},"{TARGET}",{DAYS}),1),0)','0.0'),
 ("% same day",f'=IFERROR(COUNTIFS({MES},"{TARGET}",{HRS},"<24")/COUNTIFS({MES},"{TARGET}",{HRS},">=0"),0)','0.0%'),
 ("Open (not resolved) at month-end",f'=COUNTIFS({ST},"OPEN",{MES},"{TARGET}")','0'),
 ("Portals used (month)", int(sepall['portal'].nunique()),'0')]
r=5
for lbl,frm,fmt in kpis:
    ws.cell(r,1,lbl).font=cfont; ws.cell(r,1).alignment=left; ws.cell(r,1).border=border
    box(r,2,frm,fmt,kfont); r+=1
ws.column_dimensions['A'].width=32; ws.column_dimensions['B'].width=14
title('D4',"  RESOLUTION TIME BY MONTH (last 3 months)")
hdr(5,["Month","# Uploads","Resolved","Avg. hours","Avg. days","% same day"],4)
r=6
for mm in last3:
    box(r,4,mlabel(mm),'General',bfont,left)
    box(r,5,f'=COUNTIF({MES},"{mm}")','0'); box(r,6,f'=COUNTIFS({MES},"{mm}",{HRS},">=0")','0')
    box(r,7,f'=IFERROR(ROUND(AVERAGEIF({MES},"{mm}",{HRS}),1),0)','0.0')
    box(r,8,f'=IFERROR(ROUND(AVERAGEIF({MES},"{mm}",{DAYS}),1),0)','0.0')
    box(r,9,f'=IFERROR(COUNTIFS({MES},"{mm}",{HRS},"<24")/COUNTIFS({MES},"{mm}",{HRS},">=0"),0)','0.0%'); r+=1
box(r,4,"TOTAL",'General',bfont,left); box(r,5,f'=SUM(E6:E{r-1})','0',bfont); box(r,6,f'=SUM(F6:F{r-1})','0',bfont)
box(r,7,f'=IFERROR(ROUND(AVERAGEIFS({HRS},{MES},">={last3[0]}"),1),0)','0.0',bfont)
box(r,8,f'=IFERROR(ROUND(AVERAGEIFS({DAYS},{MES},">={last3[0]}"),1),0)','0.0',bfont)
box(r,9,f'=IFERROR(COUNTIFS({MES},">={last3[0]}",{HRS},"<24")/COUNTIFS({MES},">={last3[0]}",{HRS},">=0"),0)','0.0%',bfont)
totrow=r
for col,w in zip(['D','E','F','G','H','I'],[11,11,11,12,11,12]): ws.column_dimensions[col].width=w
title('A15',f"  RESOLUTION TIME DISTRIBUTION — {TL}")
hdr(16,["Range","# Uploads","% of total"])
buckets=[("Same day (< 24 h)",'"<24"'),("1 – 2 days",f'">=24",{HRS},"<48"'),("2 – 4 days",f'">=48",{HRS},"<96"'),("4 – 7 days",f'">=96",{HRS},"<168"'),("More than 7 days",'">=168"')]
r=17; first=r
for lbl,cond in buckets:
    box(r,1,lbl,'General',cfont,left); box(r,2,f'=COUNTIFS({MES},"{TARGET}",{HRS},{cond})','0')
    box(r,3,f'=IFERROR(B{r}/COUNTIFS({MES},"{TARGET}",{HRS},">=0"),0)','0.0%'); r+=1
box(r,1,"TOTAL",'General',bfont,left); box(r,2,f'=SUM(B{first}:B{r-1})','0',bfont)
box(r,3,f'=IFERROR(B{r}/COUNTIFS({MES},"{TARGET}",{HRS},">=0"),0)','0.0%',bfont)
title('D15',f"  UPLOADS BY PORTAL — {TL}")
hdr(16,["Portal","# Uploads","% of total","Avg. days","% same day"],4)
TOTU=f'COUNTIF({MES},"{TARGET}")'; r=17
for p in NAMED:
    box(r,4,p,'General',cfont,left)
    box(r,5,f'=COUNTIFS({PRT},"{p}",{MES},"{TARGET}")','0')
    box(r,6,f'=IFERROR(COUNTIFS({PRT},"{p}",{MES},"{TARGET}")/{TOTU},0)','0.0%')
    box(r,7,f'=IFERROR(ROUND(AVERAGEIFS({DAYS},{PRT},"{p}",{MES},"{TARGET}"),1),0)','0.0')
    box(r,8,f'=IFERROR(COUNTIFS({PRT},"{p}",{MES},"{TARGET}",{HRS},"<24")/COUNTIFS({PRT},"{p}",{MES},"{TARGET}",{HRS},">=0"),0)','0.0%'); r+=1
box(r,4,"Other",'General',cfont,left); box(r,5,len(otros),'0'); box(r,6,f'=IFERROR(E{r}/{TOTU},0)','0.0%')
box(r,7,round(otros_d.mean(),1) if len(otros_d) else 0,'0.0'); box(r,8,round((otros_d<1).mean(),3) if len(otros_d) else 0,'0.0%'); r+=1
box(r,4,"TOTAL",'General',bfont,left); box(r,5,f'={TOTU}','0',bfont); box(r,6,f'=IFERROR(E{r}/{TOTU},0)','0.0%',bfont)
box(r,7,f'=IFERROR(ROUND(AVERAGEIF({MES},"{TARGET}",{DAYS}),1),0)','0.0',bfont)
box(r,8,f'=IFERROR(COUNTIFS({MES},"{TARGET}",{HRS},"<24")/COUNTIFS({MES},"{TARGET}",{HRS},">=0"),0)','0.0%',bfont)
chart=BarChart(); chart.type="col"; chart.title="Avg. resolution days by month"; chart.height=7; chart.width=13; chart.legend=None
chart.add_data(Reference(ws,min_col=8,min_row=5,max_row=totrow-1),titles_from_data=True)
chart.set_categories(Reference(ws,min_col=4,min_row=6,max_row=totrow-1)); ws.add_chart(chart,"A24")
ws.freeze_panes="A3"

op=df[(df['st']=='OPEN') & (df['mes']==TARGET)].copy()
op['ab']=op['Timestamp'].apply(lambda x:max((DCUT.date()-x.date()).days,0)); op=op.sort_values('ab',ascending=False)
pz=wb.create_sheet("Open & In Progress"); pz.sheet_view.showGridLines=False
pz['A1']=f"OPEN AT {TL} MONTH-END — not resolved by month-end"; pz['A1'].font=Font(name="Arial",bold=True,size=14,color="1F3864")
cols=["Invoice Number","Portal","Submitted By","Collector assigned","Submit Timestamp","Days open"]
for c,h in enumerate(cols,1):
    cell=pz.cell(3,c,h); cell.fill=hfill; cell.font=hfont; cell.alignment=center; cell.border=border
r=4
for _,row in op.iterrows():
    pz.cell(r,1,str(row['Invoice Number']).replace('\n',' / ') if pd.notna(row['Invoice Number']) else "")
    pz.cell(r,2,row['portal']); pz.cell(r,3,row['Submitted By'] if pd.notna(row['Submitted By']) else "")
    pz.cell(r,4,row['Collector assigned'] if pd.notna(row['Collector assigned']) else "")
    t=pz.cell(r,5,row['Timestamp'].to_pydatetime()); t.number_format=DFMT
    dd=pz.cell(r,6,int(row['ab'])); dd.font=bfont
    if row['ab']>=7: dd.fill=red
    elif row['ab']>=2: dd.fill=yel
    for c in range(1,7):
        cc=pz.cell(r,c); cc.border=border
        if c!=6: cc.font=cfont
        cc.alignment=left if c in (1,2,3,4) else center
    r+=1
for i,w in enumerate([20,16,18,16,16,10],1): pz.column_dimensions[get_column_letter(i)].width=w
pz.freeze_panes="A4"

rr=df[df['hasres']].copy(); rr['dias']=(rr['Message Timestamp']-rr['Timestamp']).dt.total_seconds()/86400
slow=rr[(rr['mes']==TARGET)&(rr['dias']>2)].sort_values('dias',ascending=False)
sz=wb.create_sheet(f"Slow Uploads {TL} (>2d)"); sz.sheet_view.showGridLines=False
sz['A1']=f"{TL} UPLOADS RESOLVED IN MORE THAN 2 DAYS"; sz['A1'].font=Font(name="Arial",bold=True,size=14,color="1F3864")
cols=["Invoice Number","Portal","Submitted By","Collector assigned","Submit Timestamp","Resolved Timestamp","Resolution days"]
for c,h in enumerate(cols,1):
    cell=sz.cell(3,c,h); cell.fill=hfill; cell.font=hfont; cell.alignment=center; cell.border=border
r=4
for _,row in slow.iterrows():
    sz.cell(r,1,str(row['Invoice Number']).replace('\n',' / ') if pd.notna(row['Invoice Number']) else "")
    sz.cell(r,2,row['portal']); sz.cell(r,3,row['Submitted By'] if pd.notna(row['Submitted By']) else ""); sz.cell(r,4,row['Collector assigned'] if pd.notna(row['Collector assigned']) else "")
    t=sz.cell(r,5,row['Timestamp'].to_pydatetime()); t.number_format=DFMT
    me=sz.cell(r,6,row['Message Timestamp'].to_pydatetime()); me.number_format=DFMT
    d=sz.cell(r,7); d.value=f'=(F{r}-E{r})'; d.number_format='0.0'; d.fill=red; d.font=bfont
    for c in range(1,8):
        cc=sz.cell(r,c); cc.border=border
        if c!=7: cc.font=cfont
        cc.alignment=left if c in (1,2,3,4) else center
    r+=1
for i,w in enumerate([20,16,18,16,16,16,15],1): sz.column_dimensions[get_column_letter(i)].width=w
sz.freeze_panes="A4"
wb.calculation=CalcProperties(calcId=124519, fullCalcOnLoad=True)
wb.save(OUT)
print("OK ->",OUT,"| month:",TL,"| named portals:",NAMED,"| open:",len(op),"| slow:",len(slow),"| rows:",LAST-1)
