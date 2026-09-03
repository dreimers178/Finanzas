const pptxgen = require("pptxgenjs");
const fs = require("fs");
const p = new pptxgen();
p.layout = "LAYOUT_16x9"; // 10 x 5.625

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
kpiCard(s,2.78,1.45,2.15,"42","Resolved",MINT);
kpiCard(s,5.06,1.45,2.15,"1.3 d","Avg. resolution",ORANGE);
kpiCard(s,7.34,1.45,2.15,"83%","Same-day",GREENA);
s.addText("Resolution time by month",{x:0.5,y:2.75,w:5,h:0.32,fontFace:FH,fontSize:13,bold:true,color:GREEN,isTextBox:true,margin:0});
monthTable(s,[
 ["Month","Uploads","Resolved","Avg. days","% same day"],
 ["Jun - 26","58","58","2.8","69%"],
 ["Jul - 26","62","61","2.9","77%"],
 ["Aug - 26","46","42","1.3","83%"]],0.5,3.12,5.0,[1.15,1.0,1.0,0.9,0.95]);
s.addChart(p.ChartType.doughnut,[{name:"Distribution",labels:["Same day","1-2 d","2-4 d","4-7 d",">7 d"],values:[35,2,1,1,3]}],
 {x:5.9,y:2.7,w:3.8,h:2.55,chartColors:[GREENA,TEAL,BLUE,ORANGE,PURPLE],
  showLegend:true,legendPos:"r",legendColor:DARK,legendFontFace:FB,legendFontSize:10,
  showTitle:true,title:"Resolution distribution",titleColor:GREEN,titleFontFace:FH,titleFontSize:13,
  showValue:true,dataLabelColor:WHITE,dataLabelFontFace:FB,dataLabelFontSize:9,holeSize:58});

// ================= SLIDE 2: Tesorio Tasks =================
s=p.addSlide(); s.background={color:WHITE}; header(s);
slideTitle(s,"Tesorio Tasks  ·  Aug - 26");
kpiCard(s,0.5,1.45,2.15,"51","Completed",BLUE);
kpiCard(s,2.78,1.45,2.15,"1.1 d","Avg. resolution",ORANGE);
kpiCard(s,5.06,1.45,2.15,"75%","Same-day",GREENA);
kpiCard(s,7.34,1.45,2.15,"9","Open today",MINT);
s.addText("Resolution time by month",{x:0.5,y:2.75,w:5,h:0.32,fontFace:FH,fontSize:13,bold:true,color:GREEN,isTextBox:true,margin:0});
monthTable(s,[
 ["Month","Created","Completed","Avg. days","% same day"],
 ["Jun - 26","30","30","5.2","37%"],
 ["Jul - 26","78","78","2.0","76%"],
 ["Aug - 26","53","51","1.1","75%"]],0.5,3.12,5.0,[1.15,1.0,1.05,0.9,0.9]);
s.addChart(p.ChartType.bar,[{name:"Avg. days",labels:["Urgent","High","Normal"],values:[0.9,5.9,0.8]}],
 {x:5.9,y:2.7,w:3.8,h:2.55,barDir:"col",chartColors:[ORANGE],
  showLegend:false,showTitle:true,title:"Avg. days by priority",titleColor:GREEN,titleFontFace:FH,titleFontSize:13,
  catAxisLabelColor:MUTED,catAxisLabelFontFace:FB,catAxisLabelFontSize:10,
  valAxisLabelColor:MUTED,valAxisLabelFontFace:FB,valAxisLabelFontSize:9,
  valGridLine:{color:"E5E5E5",size:1},catGridLine:{style:"none"},
  showValue:true,dataLabelColor:DARK,dataLabelFontFace:FB,dataLabelFontSize:10,dataLabelPosition:"outEnd"});

p.writeFile({fileName:"/home/user/Finanzas/Portal_Metrics_Monthly_Aug-26.pptx"}).then(f=>console.log("OK ->",f));
