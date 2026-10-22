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
function monthTable(s,rows,x,y,w,colW){ dataTable(s,rows,x,y,w,colW,0.38,11); }
function dataTable(s,rows,x,y,w,colW,rowH,fs){
  const styled=rows.map(function(r,ri){
    return r.map(function(c,ci){
      return {text:c,options:{
        fill:{color: ri===0?GREEN:(ri===rows.length-1?"E9F3EF":WHITE)},
        color: ri===0?WHITE:DARK, bold: ri===0||ri===rows.length-1,
        align: ci===0?"left":"center", valign:"middle", fontFace:FB, fontSize:fs,
        margin:[1,1,1,3],
        border:{type:"solid",color:"DDDDDD",pt:0.5}}};
    });
  });
  s.addTable(styled,{x,y,w,colW,rowH});
}
function tblTitle(s,txt,x,y,w){
  s.addText(txt,{x,y,w,h:0.28,fontFace:FH,fontSize:12,bold:true,color:GREEN,isTextBox:true,margin:0});
}

// ===== SLIDE 1: Portal Uploads · Sep - 26 =====
let s=p.addSlide(); s.background={color:WHITE}; header(s);
slideTitle(s,"Portal Uploads  ·  Sep - 26");
kpiCard(s,0.5,1.45,2.15,"85","Uploads",BLUE);
kpiCard(s,2.78,1.45,2.15,"73","Resolved",MINT);
kpiCard(s,5.06,1.45,2.15,"0.7 d","Avg. resolution",ORANGE);
kpiCard(s,7.34,1.45,2.15,"85%","Same-day",GREENA);
s.addText("Resolution time by month",{x:0.5,y:2.75,w:5,h:0.32,fontFace:FH,fontSize:13,bold:true,color:GREEN,isTextBox:true,margin:0});
monthTable(s,[
 ["Month","Uploads","Resolved","Avg. days","% same day"],
 ["Jul - 26","62","61","2.9","77%"],
 ["Aug - 26","46","43","1.7","81%"],
 ["Sep - 26","85","73","0.7","85%"]],0.5,3.12,5.0,[1.15,1.0,1.0,0.9,0.95]);
s.addChart(p.ChartType.doughnut,[{name:"Distribution",labels:["Same day","1-2 d","2-4 d","4-7 d",">7 d"],values:[62,10,0,0,1]}],
 {x:5.9,y:2.7,w:3.8,h:2.55,chartColors:[GREENA,TEAL,BLUE,ORANGE,PURPLE],
  showLegend:true,legendPos:"r",legendColor:DARK,legendFontFace:FB,legendFontSize:10,
  showTitle:true,title:"Resolution distribution (%)",titleColor:GREEN,titleFontFace:FH,titleFontSize:13,
  showPercent:true,showValue:false,dataLabelColor:WHITE,dataLabelFontFace:FB,dataLabelFontSize:9,holeSize:58});
s.addNotes(
"Portal Uploads — the whole team's work: submitting our clients' invoices across every customer portal. Scope: tasks created through September 30, with a 2-day SLA window to resolve (completed through October 2).\n"+
"In September we handled 85 upload requests and resolved 73, at 0.7 days on average and 85% resolved same-day.\n"+
"The trend keeps improving: 2.9 days in July, 1.7 in August, 0.7 in September.\n"+
"The donut shows how fast we close them — the large majority are same-day.");

// ===== SLIDE 2: Tesorio Tasks · Sep - 26 =====
s=p.addSlide(); s.background={color:WHITE}; header(s);
slideTitle(s,"Tesorio Tasks  ·  Sep - 26");
kpiCard(s,0.5,1.45,2.15,"204","Completed",BLUE);
kpiCard(s,2.78,1.45,2.15,"1.1 d","Avg. resolution",ORANGE);
kpiCard(s,5.06,1.45,2.15,"67%","Same-day",GREENA);
kpiCard(s,7.34,1.45,2.15,"38","Open (Oct 2)",MINT);
s.addText("Resolution time by month",{x:0.5,y:2.75,w:5,h:0.32,fontFace:FH,fontSize:13,bold:true,color:GREEN,isTextBox:true,margin:0});
monthTable(s,[
 ["Month","Created","Completed","Avg. days","% same day"],
 ["Jul - 26","80","80","3.9","72%"],
 ["Aug - 26","53","53","4.1","68%"],
 ["Sep - 26","240","204","1.1","67%"]],0.5,3.12,5.0,[1.15,1.0,1.05,0.9,0.9]);
s.addChart(p.ChartType.doughnut,[{name:"Distribution",labels:["Same day","1-2 d","2-4 d","4-7 d",">7 d"],values:[136,26,37,2,3]}],
 {x:5.9,y:2.7,w:3.8,h:2.55,chartColors:[GREENA,TEAL,BLUE,ORANGE,PURPLE],
  showLegend:true,legendPos:"r",legendColor:DARK,legendFontFace:FB,legendFontSize:10,
  showTitle:true,title:"Resolution distribution (%)",titleColor:GREEN,titleFontFace:FH,titleFontSize:13,
  showPercent:true,showValue:false,dataLabelColor:WHITE,dataLabelFontFace:FB,dataLabelFontSize:9,holeSize:58});
s.addNotes(
"Tesorio Tasks — how we track that same portal work end to end. Scope: created through September 30, with a 2-day SLA window to resolve (completed through October 2).\n"+
"In September we completed 204 tasks at 1.1 days on average, 67% same-day, with 38 still open past the SLA window. The 2-day window captures the many tasks created on Sep 30 that closed on Oct 1-2.\n"+
"The bar chart shows which platforms we use most: Coupa and Ariba carry the bulk of our volume, then Taulia.\n"+
"Opportunity: the slowest cases are URL-based ones with no standing access — getting access set up up front closes that gap. Speed and accuracy here get us paid faster and bring down DSO, while making it easier for the client to process and pay.");

p.writeFile({fileName:"/home/user/Finanzas/Portal_Metrics_Monthly_Sep-26.pptx"}).then(f=>console.log("OK ->",f));
