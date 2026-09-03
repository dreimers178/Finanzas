#!/usr/bin/env python3
"""Tesorio Tasks Report focused on one month (English).
Monthly table = last 3 months; distribution/priority/slow = target month.

Usage: python3 generar_tesorio_tasks_mes.py <export.csv> <month YYYY-MM> [output.xlsx] [today YYYY-MM-DD]
"""
import sys
from datetime import datetime
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference
from openpyxl.workbook.properties import CalcProperties

SRC=sys.argv[1]
TARGET=sys.argv[2] if len(sys.argv)>2 else None
OUT=sys.argv[3] if len(sys.argv)>3 else "/home/user/Finanzas/Tesorio_Tasks_Report_Aug_2026.xlsx"
TODAY=pd.Timestamp(sys.argv[4], tz='UTC') if len(sys.argv)>4 else pd.Timestamp('2026-09-02', tz='UTC')

def mlabel(m): return datetime.strptime(m,'%Y-%m').strftime('%b - %y')  # 2026-08 -> "Aug - 26"

df=pd.read_csv(SRC, dtype=str)
df['c']=pd.to_datetime(df['Created At'],utc=True,errors='coerce')
df['d']=pd.to_datetime(df['Completed At'],utc=True,errors='coerce')
df=df[df['c'].notna()].reset_index(drop=True)
df['Status']=df['Status'].astype(str).str.replace('_',' ',regex=False)
df['Priority']=df['Priority'].fillna('').astype(str).str.replace('_',' ',regex=False)
done=df[(df['Status']=='DONE') & df['d'].notna()].copy()
done['h']=(done['d']-done['c']).dt.total_seconds()/3600; done['dias']=done['h']/24
done['mes']=done['c'].dt.strftime('%Y-%m')
if TARGET is None: TARGET=sorted(done['mes'].unique())[-1]
last3=sorted(done['mes'].unique())[-3:]
tgt=done[done['mes']==TARGET]
TL=mlabel(TARGET)

NAVY="1F3864"
hfill=PatternFill("solid",fgColor=NAVY); sfill=PatternFill("solid",fgColor="2E5496")
hfont=Font(name="Arial",bold=True,color="FFFFFF",size=11); sfont=hfont
cfont=Font(name="Arial",size=10); bfont=Font(name="Arial",bold=True,size=10)
kfont=Font(name="Arial",bold=True,size=14,color=NAVY)
center=Alignment(horizontal="center",vertical="center"); left=Alignment(horizontal="left",vertical="center")
thin=Side(style="thin",color="D9D9D9"); border=Border(left=thin,right=thin,top=thin,bottom=thin)
red=PatternFill("solid",fgColor="F4CCCC"); yel=PatternFill("solid",fgColor="FFF2CC")
DFMT='dd-mmm-yy'

wb=Workbook()

# Data (calc)
keep=df[(df['c'].dt.strftime('%Y-%m').isin(last3)) | (df['Status'].isin(['TO DO','WORKING']))].copy()
dz=wb.active; dz.title="Data (calc)"
dcols=["Invoice Number","Status","Priority","Assigned by","Created At","Completed At",
       "Resolution (hrs)","Resolution (days)","Month (created)","Days open"]
for c,h in enumerate(dcols,1):
    cell=dz.cell(1,c,h); cell.fill=hfill; cell.font=hfont; cell.alignment=center; cell.border=border
r=2
for _,row in keep.iterrows():
    dz.cell(r,1,row['Invoice Number'] if pd.notna(row['Invoice Number']) else "")
    dz.cell(r,2,row['Status']); dz.cell(r,3,row['Priority'])
    dz.cell(r,4,row['Assigned by'] if pd.notna(row['Assigned by']) else "")
    ts=dz.cell(r,5,row['c'].tz_localize(None).to_pydatetime()); ts.number_format=DFMT
    if pd.notna(row['d']):
        me=dz.cell(r,6,row['d'].tz_localize(None).to_pydatetime()); me.number_format=DFMT
    if row['Status']=='DONE' and pd.notna(row['d']):
        h=(row['d']-row['c']).total_seconds()/3600; dz.cell(r,7,round(h,1)); dz.cell(r,8,round(h/24,1))
    dz.cell(r,7).number_format='0.0'; dz.cell(r,8).number_format='0.0'
    dz.cell(r,9,row['c'].strftime('%Y-%m'))
    if row['Status'] in ('TO DO','WORKING'):
        dz.cell(r,10, max((TODAY.date()-row['c'].date()).days,0)); dz.cell(r,10).number_format='0'
    for c in range(1,11):
        cc=dz.cell(r,c); cc.font=cfont; cc.border=border; cc.alignment=left if c in (1,2,3,4) else center
    r+=1
LAST=r-1
for i,w in enumerate([16,10,15,16,14,14,15,15,13,12],1): dz.column_dimensions[get_column_letter(i)].width=w
dz.freeze_panes="A2"
D="'Data (calc)'"; ST=f"{D}!$B$2:$B${LAST}"; PRI=f"{D}!$C$2:$C${LAST}"
HRS=f"{D}!$G$2:$G${LAST}"; DAYS=f"{D}!$H$2:$H${LAST}"; MES=f"{D}!$I$2:$I${LAST}"

ws=wb.create_sheet("Tesorio Tasks Report", 0); ws.sheet_view.showGridLines=False
def title(cell,txt): ws[cell]=txt; ws[cell].font=sfont; ws[cell].fill=sfill; ws[cell].alignment=left
def hdr(row,cols,sc=1):
    for i,h in enumerate(cols):
        cell=ws.cell(row,sc+i,h); cell.fill=hfill; cell.font=hfont; cell.alignment=center; cell.border=border
def box(row,col,val,fmt='General',font=cfont,al=center):
    cell=ws.cell(row,col,val); cell.font=font; cell.alignment=al; cell.number_format=fmt; cell.border=border; return cell

ws['A1']=f"TESORIO TASKS REPORT — {TL}"; ws['A1'].font=Font(name="Arial",bold=True,size=16,color=NAVY)

# Summary KPIs
title('A4',f"  {TL} SUMMARY")
kpis=[("Created in month",f'=COUNTIF({MES},"{TARGET}")','0'),
 ("Completed (DONE)",f'=COUNTIFS({MES},"{TARGET}",{HRS},">=0")','0'),
 ("Avg. resolution (days)",f'=IFERROR(ROUND(AVERAGEIF({MES},"{TARGET}",{DAYS}),1),0)','0.0'),
 ("% same day",f'=IFERROR(COUNTIFS({MES},"{TARGET}",{HRS},"<24")/COUNTIFS({MES},"{TARGET}",{HRS},">=0"),0)','0.0%'),
 ("Pending (TO DO) today",f'=COUNTIF({ST},"TO DO")','0'),
 ("In progress (WORKING) today",f'=COUNTIF({ST},"WORKING")','0')]
r=5
for lbl,frm,fmt in kpis:
    ws.cell(r,1,lbl).font=cfont; ws.cell(r,1).alignment=left; ws.cell(r,1).border=border
    box(r,2,frm,fmt,kfont); r+=1
ws.column_dimensions['A'].width=28; ws.column_dimensions['B'].width=14

# Table 1: by month (last 3), no median
title('D4',"  RESOLUTION TIME BY MONTH (last 3 months)")
hdr(5,["Month","# Created","Completed","Avg. hours","Avg. days","% same day"],4)
r=6
for m in last3:
    box(r,4,mlabel(m),'General',bfont,left)
    box(r,5,f'=COUNTIF({MES},"{m}")','0')
    box(r,6,f'=COUNTIFS({MES},"{m}",{HRS},">=0")','0')
    box(r,7,f'=IFERROR(ROUND(AVERAGEIF({MES},"{m}",{HRS}),1),0)','0.0')
    box(r,8,f'=IFERROR(ROUND(AVERAGEIF({MES},"{m}",{DAYS}),1),0)','0.0')
    box(r,9,f'=IFERROR(COUNTIFS({MES},"{m}",{HRS},"<24")/COUNTIFS({MES},"{m}",{HRS},">=0"),0)','0.0%')
    r+=1
box(r,4,"TOTAL",'General',bfont,left); box(r,5,f'=SUM(E6:E{r-1})','0',bfont); box(r,6,f'=SUM(F6:F{r-1})','0',bfont)
box(r,7,f'=IFERROR(ROUND(AVERAGEIFS({HRS},{MES},">={last3[0]}"),1),0)','0.0',bfont)
box(r,8,f'=IFERROR(ROUND(AVERAGEIFS({DAYS},{MES},">={last3[0]}"),1),0)','0.0',bfont)
box(r,9,f'=IFERROR(COUNTIFS({MES},">={last3[0]}",{HRS},"<24")/COUNTIFS({MES},">={last3[0]}",{HRS},">=0"),0)','0.0%',bfont)
totrow=r
for col,w in zip(['D','E','F','G','H','I'],[11,11,12,12,11,12]): ws.column_dimensions[col].width=w

# Table 2: distribution (month)
title('A15',f"  RESOLUTION TIME DISTRIBUTION — {TL}")
hdr(16,["Range","# Tasks","% of total"])
buckets=[("Same day (< 24 h)",'"<24"'),("1 – 2 days",f'">=24",{HRS},"<48"'),("2 – 4 days",f'">=48",{HRS},"<96"'),
 ("4 – 7 days",f'">=96",{HRS},"<168"'),("More than 7 days",'">=168"')]
r=17; first=r
for lbl,cond in buckets:
    box(r,1,lbl,'General',cfont,left); box(r,2,f'=COUNTIFS({MES},"{TARGET}",{HRS},{cond})','0')
    box(r,3,f'=IFERROR(B{r}/COUNTIFS({MES},"{TARGET}",{HRS},">=0"),0)','0.0%'); r+=1
box(r,1,"TOTAL",'General',bfont,left); box(r,2,f'=SUM(B{first}:B{r-1})','0',bfont)
box(r,3,f'=IFERROR(B{r}/COUNTIFS({MES},"{TARGET}",{HRS},">=0"),0)','0.0%',bfont)

# Table 3: priority (month), no median
title('D15',f"  RESOLUTION TIME BY PRIORITY — {TL}")
hdr(16,["Priority","# Tasks","% of total","Avg. days","% same day"],4)
prios=["URGENT","HIGH PRIORITY","NORMAL"]
TOT=f'COUNTIFS({MES},"{TARGET}",{HRS},">=0")'; r=17
for p in prios:
    box(r,4,p,'General',cfont,left)
    box(r,5,f'=COUNTIFS({PRI},"{p}",{MES},"{TARGET}",{HRS},">=0")','0')
    box(r,6,f'=IFERROR(COUNTIFS({PRI},"{p}",{MES},"{TARGET}",{HRS},">=0")/{TOT},0)','0.0%')
    box(r,7,f'=IFERROR(ROUND(AVERAGEIFS({DAYS},{PRI},"{p}",{MES},"{TARGET}"),1),0)','0.0')
    box(r,8,f'=IFERROR(COUNTIFS({PRI},"{p}",{MES},"{TARGET}",{HRS},"<24")/COUNTIFS({PRI},"{p}",{MES},"{TARGET}",{HRS},">=0"),0)','0.0%')
    r+=1
box(r,4,"TOTAL",'General',bfont,left); box(r,5,f'={TOT}','0',bfont); box(r,6,'100%','0%',bfont)
box(r,7,f'=IFERROR(ROUND(AVERAGEIF({MES},"{TARGET}",{DAYS}),1),0)','0.0',bfont)
box(r,8,f'=IFERROR(COUNTIFS({MES},"{TARGET}",{HRS},"<24")/{TOT},0)','0.0%',bfont)

# Simple chart: avg days by month
chart=BarChart(); chart.title="Avg. resolution days by month"; chart.height=7; chart.width=13
chart.y_axis.title="Days"; chart.legend=None
chart.add_data(Reference(ws,min_col=8,min_row=5,max_row=totrow-1),titles_from_data=True)
chart.set_categories(Reference(ws,min_col=4,min_row=6,max_row=totrow-1))
ws.add_chart(chart,"A24")
ws.freeze_panes="A3"

# Open & In Progress
op=df[df['Status'].isin(['TO DO','WORKING'])].copy()
op['ab']=op['c'].apply(lambda x:max((TODAY.date()-x.date()).days,0)); op=op.sort_values('ab',ascending=False)
pz=wb.create_sheet("Open & In Progress"); pz.sheet_view.showGridLines=False
pz['A1']="OPEN TASKS — Pending (TO DO) and In Progress (WORKING)"; pz['A1'].font=Font(name="Arial",bold=True,size=14,color=NAVY)
cols=["Invoice Number","Status","Priority","Assigned by","Created At","Days open"]
for c,h in enumerate(cols,1):
    cell=pz.cell(3,c,h); cell.fill=hfill; cell.font=hfont; cell.alignment=center; cell.border=border
r=4
for _,row in op.iterrows():
    pz.cell(r,1,row['Invoice Number'] if pd.notna(row['Invoice Number']) else "")
    pz.cell(r,2,row['Status']); pz.cell(r,3,row['Priority'])
    pz.cell(r,4,row['Assigned by'] if pd.notna(row['Assigned by']) else "")
    t=pz.cell(r,5,row['c'].tz_localize(None).to_pydatetime()); t.number_format=DFMT
    dd=pz.cell(r,6,int(row['ab'])); dd.font=bfont
    if row['ab']>=7: dd.fill=red
    elif row['ab']>=2: dd.fill=yel
    for c in range(1,7):
        cc=pz.cell(r,c); cc.border=border
        if c!=6: cc.font=cfont
        cc.alignment=left if c in (1,3,4) else center
    r+=1
for i,w in enumerate([16,12,15,16,14,12],1): pz.column_dimensions[get_column_letter(i)].width=w
pz.freeze_panes="A4"

# Slow tasks (month)
slow=tgt[tgt['dias']>=2].sort_values('dias',ascending=False)
sz=wb.create_sheet(f"Slow Tasks {TL} (>2d)"); sz.sheet_view.showGridLines=False
sz['A1']=f"{TL} TASKS RESOLVED IN MORE THAN 2 DAYS"; sz['A1'].font=Font(name="Arial",bold=True,size=14,color=NAVY)
cols=["Invoice Number","Priority","Assigned by","Created At","Completed At","Resolution days"]
for c,h in enumerate(cols,1):
    cell=sz.cell(3,c,h); cell.fill=hfill; cell.font=hfont; cell.alignment=center; cell.border=border
r=4
for _,row in slow.iterrows():
    sz.cell(r,1,row['Invoice Number'] if pd.notna(row['Invoice Number']) else "")
    sz.cell(r,2,row['Priority']); sz.cell(r,3,row['Assigned by'] if pd.notna(row['Assigned by']) else "")
    t=sz.cell(r,4,row['c'].tz_localize(None).to_pydatetime()); t.number_format=DFMT
    m=sz.cell(r,5,row['d'].tz_localize(None).to_pydatetime()); m.number_format=DFMT
    d=sz.cell(r,6); d.value=f'=(E{r}-D{r})'; d.number_format='0.0'; d.fill=red; d.font=bfont
    for c in range(1,7):
        cc=sz.cell(r,c); cc.border=border
        if c!=6: cc.font=cfont
        cc.alignment=left if c in (1,3) else center
    r+=1
for i,w in enumerate([16,15,16,14,14,15],1): sz.column_dimensions[get_column_letter(i)].width=w
sz.freeze_panes="A4"

wb.calculation=CalcProperties(calcId=124519, fullCalcOnLoad=True)
wb.save(OUT)
print("OK ->",OUT,"| month:",TL,"| last3:",[mlabel(m) for m in last3],"| slow:",len(slow),"| open:",len(op))
