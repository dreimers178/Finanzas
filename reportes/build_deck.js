const pptxgen = require("pptxgenjs");
const fs = require("fs");
const p = new pptxgen();
p.layout = "LAYOUT_16x9";

const GREEN="158158", DARK="1F2A44", MUTED="6B7B84", LIGHT="F3F5F7", WHITE="FFFFFF";
const BLUE="4292E1", TEAL="53B5E0", MINT="34B78F", PURPLE="4A2ADB", ORANGE="ED8B1B", GREENA="50B432";
const FH="Calibri", FB="Calibri";
const HDR = "image/png;base64," + fs.readFileSync("/tmp/qheader.png").toString("base64");

function header(s){
  s.addImage({data:HDR,x:0,y:0,w:10,h:0.62});
  s.addText([{text:"qualtrics",options:{fontSize:17,bold:true}},{text:"XM",options:{fontSize:9,superscript:true,bold:true}}],
    {x:7.7,y:0.13,w:2.1,h:0.36,align:"right",color:WHITE,fontFace:FH,isTextBox:true,margin:0});
}
function kpiCard(s,x,y,w,val,label,color){
  s.addShape(p.ShapeType.roundRect,{x,y,w,h:1.05,fill:{color:LIGHT},line:{type:"none"},rectRadius:0.07});
  s.addText(val,{x,y:y+0.1,w,h:0.55,align:"center",fontFace:FH,fontSize:28,bold:true,color,isTextBox:true,margin:0});
  s.addText(label,{x,y:y+0.66,w,h:0.32,align:"center",fontFace:FB,fontSize:11,color:MUTED,isTextBox:true,margin:0});
}
function slideTitle(s,txt){
  s.addText(txt,{x:0.5,y:0.78,w:9,h:0.5,fontFace:FH,fontSize:24,bold:true,color:GREEN,isTextBox:true,margin:0});
}
function monthTable(s,rows,x,y,w,colW){
  const styled=rows.map(function(r,ri){
    return r.map(function(c){
      return {text:c,options:{
        fill:{color: ri===0?GREEN:(ri===rows.length-1?"E9F3EF":WHITE)},
        color: ri===0?WHITE:DARK, bold: ri===0||ri===rows.length-1,
        align:"center", valign:"middle", fontFace:FB, fontSize:11,
        border:{type:"solid",color:"DDDDDD",pt:0.5}}};
    });
  });
  s.addTable(styled,{x,y,w,colW,rowH:0.38});
}

// ================= SLIDE 1: Portal Uploads =================
let s=p.addSlide(); s.background={color:WHITE}; header(s);
slideTitle(s,"Portal Uploads  ·  Aug - 26");
kpiCard(s,0.5,1.45,2.15,"46","Uploads",BLUE);
kpiCard(s,2.78,1.45,2.15,"41","Resolved",MINT);
kpiCard(s,5.06,1.45,2.15,"0.9 d","Avg. resolution",ORANGE);
kpiCard(s,7.34,1.45,2.15,"85%","Same-day",GREENA);
s.addText("Resolution time by month",{x:0.5,y:2.75,w:5,h:0.32,fontFace:FH,fontSize:13,bold:true,color:GREEN,isTextBox:true,margin:0});
monthTable(s,[
 ["Month","Uploads","Resolved","Avg. days","% same day"],
 ["Jun - 26","58","58","2.8","69%"],
 ["Jul - 26","62","61","2.9","77%"],
 ["Aug - 26","46","41","0.9","85%"]],0.5,3.12,5.0,[1.15,1.0,1.0,0.9,0.95]);
s.addChart(p.ChartType.doughnut,[{name:"Distribution",labels:["Same day","1-2 d","2-4 d","4-7 d",">7 d"],values:[35,2,1,1,2]}],
 {x:5.9,y:2.7,w:3.8,h:2.55,chartColors:[GREENA,TEAL,BLUE,ORANGE,PURPLE],
  showLegend:true,legendPos:"r",legendColor:DARK,legendFontFace:FB,legendFontSize:10,
  showTitle:true,title:"Resolution distribution (%)",titleColor:GREEN,titleFontFace:FH,titleFontSize:13,
  showPercent:true,showValue:false,dataLabelColor:WHITE,dataLabelFontFace:FB,dataLabelFontSize:9,holeSize:58});
s.addNotes(
"Portal Uploads — this is the whole team's work: submitting our clients' invoices across every customer portal.\n"+
"In August we handled 46 upload requests and resolved 41 of them, at an average of under one day (0.9) and 85% resolved the same day.\n"+
"The month-by-month table shows the trend: we went from ~2.8 days in June to under a day in August — a clear, steady improvement.\n"+
"The donut shows how fast we close them: about 85% are done same-day, and only a small tail takes longer.\n"+
"Numbers are cut off at August 31 — September is not included.");

// ================= SLIDE 2: Tesorio Tasks =================
s=p.addSlide(); s.background={color:WHITE}; header(s);
slideTitle(s,"Tesorio Tasks  ·  Aug - 26");
kpiCard(s,0.5,1.45,2.15,"50","Completed",BLUE);
kpiCard(s,2.78,1.45,2.15,"1.1 d","Avg. resolution",ORANGE);
kpiCard(s,5.06,1.45,2.15,"74%","Same-day",GREENA);
kpiCard(s,7.34,1.45,2.15,"4","Open (Aug 31)",MINT);
s.addText("Resolution time by month",{x:0.5,y:2.75,w:5,h:0.32,fontFace:FH,fontSize:13,bold:true,color:GREEN,isTextBox:true,margin:0});
monthTable(s,[
 ["Month","Created","Completed","Avg. days","% same day"],
 ["Jun - 26","30","30","5.2","37%"],
 ["Jul - 26","80","78","2.0","76%"],
 ["Aug - 26","53","50","1.1","74%"]],0.5,3.12,5.0,[1.15,1.0,1.05,0.9,0.9]);
s.addChart(p.ChartType.bar,[{name:"Invoices",labels:["Coupa","Ariba","Other","Taulia","Oracle","URL"],values:[20,12,5,4,3,2]}],
 {x:5.9,y:2.7,w:3.8,h:2.55,barDir:"bar",chartColors:[GREEN],
  showLegend:false,showTitle:true,title:"Invoices by portal (platforms we use most)",titleColor:GREEN,titleFontFace:FH,titleFontSize:12,
  catAxisLabelColor:MUTED,catAxisLabelFontFace:FB,catAxisLabelFontSize:10,
  valAxisHidden:true,valGridLine:{style:"none"},catGridLine:{style:"none"},
  showValue:true,dataLabelColor:DARK,dataLabelFontFace:FB,dataLabelFontSize:10,dataLabelPosition:"outEnd"});
s.addNotes(
"Tesorio Tasks is how we track that same portal work end to end. In August we completed 50 tasks at 1.1 days on average, 74% same-day, and only 4 were still open at month-end.\n"+
"The bar chart shows which platforms we use most: Coupa and Ariba carry the bulk of our volume, followed by Taulia and Oracle.\n"+
"Area of opportunity: the slowest cases are the ones that come in as a raw URL, not the portals we already use — about 8 days versus under 1 day for Coupa or Ariba. If we can get proper access or credentials up front for those, we close the gap.\n"+
"What this means for us: every upload is an invoice we need the client to pay. Going from three days to under a day gets our invoices in front of them sooner and correct the first time — so we get paid faster, with fewer chases and escalations. Speed and accuracy here protect our own cash flow and bring down DSO, while making it easier for the client to process and pay. This month the team did that really well.\n"+
"Cut off at August 31 — September excluded.");

p.writeFile({fileName:"/home/user/Finanzas/Portal_Metrics_Monthly_Aug-26.pptx"}).then(f=>console.log("OK ->",f));
