const pptxgen = require("pptxgenjs");
const p = new pptxgen();
p.layout = "LAYOUT_16x9"; // 10 x 5.625

const GREEN="158158", LIGHT="F3F3F3", WHITE="FFFFFF", DARK="1A2E2A", MUTED="6B7B77";
const BLUE="058DC7", LIME="50B432", ORANGE="ED561B", CYAN="24CBE5";
const FH="Calibri", FB="Calibri";
const W=10, H=5.625;

function kpiCard(s,x,y,w,val,label,color){
  s.addShape(p.ShapeType.roundRect,{x,y,w,h:1.15,fill:{color:LIGHT},line:{type:"none"},rectRadius:0.08});
  s.addText(val,{x,y:y+0.12,w,h:0.6,align:"center",fontFace:FH,fontSize:30,bold:true,color,isTextBox:true,margin:0});
  s.addText(label,{x,y:y+0.72,w,h:0.35,align:"center",fontFace:FB,fontSize:11,color:MUTED,isTextBox:true,margin:0});
}
function sectionTitle(s,txt){
  s.addText(txt,{x:0.5,y:0.35,w:9,h:0.6,fontFace:FH,fontSize:26,bold:true,color:GREEN,isTextBox:true,margin:0});
}
function chartFrame(){return{showLegend:false,showTitle:true,titleColor:GREEN,titleFontFace:FH,titleFontSize:13,
  catAxisLabelColor:MUTED,catAxisLabelFontFace:FB,catAxisLabelFontSize:10,
  valAxisLabelColor:MUTED,valAxisLabelFontFace:FB,valAxisLabelFontSize:9,
  valGridLine:{color:"E5E5E5",size:1},catGridLine:{style:"none"},
  showValue:true,dataLabelColor:DARK,dataLabelFontFace:FB,dataLabelFontSize:10,dataLabelPosition:"outEnd"};}

// ---------- Slide 1: Title ----------
let s=p.addSlide(); s.background={color:GREEN};
s.addText("Portal Metrics",{x:0.6,y:1.6,w:8.8,h:1.0,fontFace:FH,fontSize:44,bold:true,color:WHITE,isTextBox:true,margin:0});
s.addText("Monthly Meeting  ·  Aug - 26",{x:0.62,y:2.7,w:8.8,h:0.5,fontFace:FB,fontSize:20,color:CYAN,isTextBox:true,margin:0});
s.addText("Portal Uploads  &  Tesorio Tasks",{x:0.62,y:3.25,w:8.8,h:0.4,fontFace:FB,fontSize:14,color:LIGHT,isTextBox:true,margin:0});

// ---------- Slide 2: Portal Uploads overview ----------
s=p.addSlide(); s.background={color:WHITE};
sectionTitle(s,"Portal Uploads  —  Aug - 26");
kpiCard(s,0.5,1.05,2.15,"46","Uploads",GREEN);
kpiCard(s,2.78,1.05,2.15,"42","Resolved",BLUE);
kpiCard(s,5.06,1.05,2.15,"1.3 d","Avg. resolution",ORANGE);
kpiCard(s,7.34,1.05,2.15,"83%","Same-day",LIME);
s.addText("Resolution time by month",{x:0.5,y:2.5,w:4.6,h:0.35,fontFace:FH,fontSize:13,bold:true,color:GREEN,isTextBox:true,margin:0});
const pRows=[
 [{text:"Month"},{text:"Uploads"},{text:"Resolved"},{text:"Avg. days"},{text:"% same day"}],
 [{text:"Jun - 26"},{text:"58"},{text:"58"},{text:"2.8"},{text:"69%"}],
 [{text:"Jul - 26"},{text:"62"},{text:"61"},{text:"2.9"},{text:"77%"}],
 [{text:"Aug - 26"},{text:"46"},{text:"42"},{text:"1.3"},{text:"83%"}]];
s.addTable(pRows,{x:0.5,y:2.9,w:4.6,colW:[1.1,0.9,0.95,0.85,1.0],rowH:0.4,
  fontFace:FB,fontSize:11,color:DARK,valign:"middle",align:"center",border:{type:"solid",color:"DDDDDD",pt:0.5},
  fill:{color:WHITE}});
// header + Aug highlight
p.tableHeaderFix=true;
s.addChart(p.ChartType.bar,[{name:"Avg. days",labels:["Jun - 26","Jul - 26","Aug - 26"],values:[2.8,2.9,1.3]}],
  {x:5.4,y:2.5,w:4.2,h:2.75,barDir:"col",chartColors:[BLUE],...chartFrame(),title:"Avg. resolution days by month"});

// ---------- Slide 3: Portal breakdown ----------
s=p.addSlide(); s.background={color:WHITE};
sectionTitle(s,"Portal Uploads  —  August Breakdown");
s.addChart(p.ChartType.doughnut,[{name:"Distribution",labels:["Same day","1-2 d","2-4 d","4-7 d",">7 d"],values:[35,2,1,1,3]}],
  {x:0.5,y:1.2,w:4.4,h:3.9,chartColors:[LIME,CYAN,BLUE,ORANGE,"C0392B"],showLegend:true,legendPos:"b",legendColor:DARK,legendFontFace:FB,legendFontSize:10,
   showTitle:true,title:"Resolution time distribution",titleColor:GREEN,titleFontFace:FH,titleFontSize:13,
   showValue:true,dataLabelColor:WHITE,dataLabelFontFace:FB,dataLabelFontSize:10,holeSize:55});
s.addChart(p.ChartType.bar,[{name:"Uploads",labels:["Coupa","Ariba","Oracle","Taulia","Other","URL"],values:[19,12,3,3,3,2]}],
  {x:5.2,y:1.2,w:4.4,h:3.9,barDir:"bar",chartColors:[GREEN],showLegend:false,
   showTitle:true,title:"Uploads by portal",titleColor:GREEN,titleFontFace:FH,titleFontSize:13,
   catAxisLabelColor:MUTED,catAxisLabelFontFace:FB,catAxisLabelFontSize:10,valAxisHidden:true,
   valGridLine:{style:"none"},catGridLine:{style:"none"},showValue:true,dataLabelColor:DARK,dataLabelFontFace:FB,dataLabelFontSize:10,dataLabelPosition:"outEnd"});

// ---------- Slide 4: Tesorio overview ----------
s=p.addSlide(); s.background={color:WHITE};
sectionTitle(s,"Tesorio Tasks  —  Aug - 26");
kpiCard(s,0.5,1.05,2.15,"51","Completed",GREEN);
kpiCard(s,2.78,1.05,2.15,"1.1 d","Avg. resolution",ORANGE);
kpiCard(s,5.06,1.05,2.15,"75%","Same-day",LIME);
kpiCard(s,7.34,1.05,2.15,"9","Open today",BLUE);
s.addText("Resolution time by month",{x:0.5,y:2.5,w:4.6,h:0.35,fontFace:FH,fontSize:13,bold:true,color:GREEN,isTextBox:true,margin:0});
const tRows=[
 [{text:"Month"},{text:"Created"},{text:"Completed"},{text:"Avg. days"},{text:"% same day"}],
 [{text:"Jun - 26"},{text:"30"},{text:"30"},{text:"5.2"},{text:"37%"}],
 [{text:"Jul - 26"},{text:"78"},{text:"78"},{text:"2.0"},{text:"76%"}],
 [{text:"Aug - 26"},{text:"53"},{text:"51"},{text:"1.1"},{text:"75%"}]];
s.addTable(tRows,{x:0.5,y:2.9,w:4.6,colW:[1.1,0.9,1.0,0.85,0.95],rowH:0.4,
  fontFace:FB,fontSize:11,color:DARK,valign:"middle",align:"center",border:{type:"solid",color:"DDDDDD",pt:0.5},fill:{color:WHITE}});
s.addChart(p.ChartType.bar,[{name:"Avg. days",labels:["Jun - 26","Jul - 26","Aug - 26"],values:[5.2,2.0,1.1]}],
  {x:5.4,y:2.5,w:4.2,h:2.75,barDir:"col",chartColors:[GREEN],...chartFrame(),title:"Avg. resolution days by month"});

// ---------- Slide 5: Tesorio breakdown ----------
s=p.addSlide(); s.background={color:WHITE};
sectionTitle(s,"Tesorio Tasks  —  August Breakdown");
s.addChart(p.ChartType.doughnut,[{name:"Distribution",labels:["Same day","1-2 d","2-4 d","4-7 d",">7 d"],values:[38,4,6,1,2]}],
  {x:0.5,y:1.2,w:4.4,h:3.9,chartColors:[LIME,CYAN,BLUE,ORANGE,"C0392B"],showLegend:true,legendPos:"b",legendColor:DARK,legendFontFace:FB,legendFontSize:10,
   showTitle:true,title:"Resolution time distribution",titleColor:GREEN,titleFontFace:FH,titleFontSize:13,
   showValue:true,dataLabelColor:WHITE,dataLabelFontFace:FB,dataLabelFontSize:10,holeSize:55});
s.addChart(p.ChartType.bar,[{name:"Avg. days",labels:["Urgent","High","Normal"],values:[0.9,5.9,0.8]}],
  {x:5.2,y:1.2,w:4.4,h:3.15,barDir:"col",chartColors:[ORANGE],...chartFrame(),title:"Avg. resolution days by priority"});
s.addText("9 open tasks today  ·  6 To Do + 3 Working  ·  oldest 43 days",
  {x:5.2,y:4.55,w:4.4,h:0.5,fontFace:FB,fontSize:11,italic:true,color:MUTED,align:"center",isTextBox:true,margin:0});

// ---------- Slide 6: Takeaways ----------
s=p.addSlide(); s.background={color:GREEN};
s.addText("Key Takeaways — Aug - 26",{x:0.6,y:0.5,w:8.8,h:0.7,fontFace:FH,fontSize:30,bold:true,color:WHITE,isTextBox:true,margin:0});
const pts=[
 "Portal Uploads: 1.3 days avg (down from 2.9 in Jul) · 83% resolved same day.",
 "Tesorio Tasks: 1.1 days avg · 75% same day · sustained improvement vs. backlog.",
 "Both workflows trending down month over month.",
 "Watch: 9 Tesorio tasks still open (oldest 43 days); 9 tasks took >2 days in Aug."];
let yy=1.6;
pts.forEach(t=>{
  s.addShape(p.ShapeType.ellipse,{x:0.7,y:yy+0.06,w:0.16,h:0.16,fill:{color:CYAN},line:{type:"none"}});
  s.addText(t,{x:1.05,y:yy-0.05,w:8.3,h:0.6,fontFace:FB,fontSize:15,color:WHITE,isTextBox:true,margin:0});
  yy+=0.78;
});

p.writeFile({fileName:"/home/user/Finanzas/Portal_Metrics_Monthly_Aug-26.pptx"}).then(f=>console.log("OK ->",f));
