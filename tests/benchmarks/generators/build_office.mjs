import fs from 'node:fs/promises';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {Workbook, SpreadsheetFile, Presentation, PresentationFile} from '@oai/artifact-tool';

const base=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const kit=path.join(base,'generated/scribble-test-kit-v1');
const qa=path.join(base,'generated/qa'); await fs.mkdir(qa,{recursive:true});
const data=JSON.parse(await fs.readFile(path.join(base,'sources/atlas-v1.json'),'utf8'));
const headers=['RowID','Period','Region','Product','RevenueEUR','CostEUR'];
function table(w,name,rows){
  const s=w.worksheets.add(name); s.showGridLines=false;
  s.getRangeByIndexes(0,0,rows.length,rows[0].length).values=rows;
  s.getRangeByIndexes(0,0,1,rows[0].length).format={fill:'#18344D',font:{bold:true,color:'#FFFFFF'},rowHeight:28};
  s.getUsedRange().format.columnWidth=18; s.getUsedRange().format.rowHeight=25;
  s.freezePanes.freezeRows(1); return s;
}
for(const [name,rows] of [['Atlas-input',data.sales],['Atlas-dirty',data.dirty],['Atlas-missing',data.missing]]){
  const w=Workbook.create(); table(w,'Sales',[headers,...rows]); table(w,'Budget',[['Region','Product','BudgetEUR'],...data.budget]);
  w.recalculate(); await (await SpreadsheetFile.exportXlsx(w)).save(path.join(kit,'inputs/excel',name+'.xlsx'));
  if(name==='Atlas-input'){
    const p=await w.render({sheetName:'Sales',range:'A1:F9',scale:1,format:'png'}); await fs.writeFile(path.join(qa,'sales.png'),new Uint8Array(await p.arrayBuffer()));
  }
}
const w=Workbook.create();
const summary=table(w,'Scribble Draft',[['Metric','May','June','Change'],['Revenue EUR',null,null,null],['Cost EUR',null,null,null],['Gross profit EUR',null,null,null],['Gross margin',null,null,null],['Budget EUR',null,null,null],['Budget variance EUR',null,null,null],['Budget variance %',null,null,null]]);
table(w,'Sales',[headers,...data.sales]);table(w,'Budget',[['Region','Product','BudgetEUR'],...data.budget]);
summary.getRange('B2:D5').formulas=[['=SUM(Sales!E2:E5)','=SUM(Sales!E6:E9)','=C2/B2-1'],['=SUM(Sales!F2:F5)','=SUM(Sales!F6:F9)','=C3-B3'],['=B2-B3','=C2-C3','=C4-B4'],['=B4/B2','=C4/C2','=100*(C5-B5)']];
summary.getRange('C6:C8').formulas=[['=SUM(Budget!C2:C5)'],['=C2-C6'],['=C7/C6']];
summary.getRange('B2:C4').setNumberFormat('#,##0.00');summary.getRange('B5:D5').setNumberFormat('0.00%');summary.getRange('D2').setNumberFormat('0.00%');summary.getRange('C8').setNumberFormat('0.00%');
summary.getRange('D5').setNumberFormat('0.00" pp"');
summary.getRange('F1:G3').values=[['June','EUR'],['Actual',null],['Budget',null]];summary.getRange('G2:G3').formulas=[['=C2'],['=C6']];
const chart=summary.charts.add('bar',summary.getRange('F1:G3'));chart.title='June actual versus budget';chart.setPosition('F5','N20');chart.hasLegend=false;
chart.yAxis={min:0,max:140000,numberFormatCode:'#,##0',numberFormatSourceLinked:false,textStyle:{typeface:'Arial',fontSize:14}};
chart.xAxis={axisType:'textAxis',textStyle:{typeface:'Arial',fontSize:14}};
summary.getRange('A11:D15').values=[['June drivers','Actual EUR','Budget EUR','Variance EUR'],['North A',40000,45000,null],['South A',25000,25000,null],['North B',30000,30000,null],['South B',25000,30000,null]];
summary.getRange('D12:D15').formulas=[['=B12-C12'],['=B13-C13'],['=B14-C14'],['=B15-C15']];
summary.getRange('B12:D15').setNumberFormat('#,##0.00');summary.getRange('C6:C7').setNumberFormat('#,##0.00');summary.getRange('D3:D4').setNumberFormat('#,##0.00');summary.getRange('G2:G3').setNumberFormat('#,##0.00');
summary.getRange('A11:D11').format={fill:'#18344D',font:{bold:true,color:'#FFFFFF'}};
w.recalculate();await (await SpreadsheetFile.exportXlsx(w)).save(path.join(kit,'evaluator-only/expected-analysis.xlsx'));
await fs.writeFile(path.join(qa,'reference-values.json'),JSON.stringify(summary.getRange('A1:D8').values,null,2));
const p=await w.render({sheetName:'Scribble Draft',range:'A1:N21',scale:1,format:'png'});await fs.writeFile(path.join(qa,'analysis.png'),new Uint8Array(await p.arrayBuffer()));

function box(slide,text,left,top,width,height,size=26,color='#18344D',bold=false){
 const shape=slide.shapes.add({geometry:'textbox',position:{left,top,width,height},fill:'none',line:{fill:'none',width:0}});
 shape.text=text;shape.text.style={typeface:'Arial',fontSize:size,color,bold,autoFit:'none'};return shape;
}
function slide(deck,title,number){
 const s=deck.slides.add();s.background.fill='#FFFFFF';box(s,'ATLAS  /  JUNE 2026',60,30,1100,25,16,'#47747B',true);box(s,title,60,90,1150,100,42,'#18344D',true);box(s,'Synthetic benchmark • EUR excluding tax',60,668,1000,24,15,'#62717C');box(s,String(number),1180,668,40,24,15,'#62717C');return s;
}
const deck=Presentation.create({slideSize:{width:1280,height:720}});
let s=slide(deck,'Revenue grows but misses budget',1);
for(const [i,value,label] of [[0,'EUR 120,000','June revenue'],[1,'+20%','Versus May'],[2,'−7.69%','Versus budget']]){box(s,value,60+i*400,230,380,75,48,'#176B74',true);box(s,label,60+i*400,310,380,50,25);}
box(s,'Gross margin 38.33%  |  On-time delivery 94%',60,450,1120,70,32);box(s,'Cost and delivery actions remain planned.',60,540,1100,70,28);
s.speakerNotes.textFrame.setText('Sources: sales.csv, budget.csv, Atlas-operations-note.pdf; final finance email supersedes EUR 118000 preliminary estimate.');
s=slide(deck,'June revenue is EUR 10,000 below budget',2);
s.charts.add('bar',{position:{left:100,top:215,width:1080,height:370},categories:['Actual','Budget'],series:[{name:'Revenue EUR',values:[120000,130000],fill:'#176B74'}],barOptions:{direction:'column',grouping:'clustered'},hasLegend:false,dataLabels:{showValue:true,position:'outEnd'}});
s.speakerNotes.textFrame.setText('Source: sales.csv and budget.csv. Variance = 120000 - 130000 = -10000; percentage = -10000 / 130000 = -7.69%.');
s=slide(deck,'Gross margin declines by 1.67 points',3);
s.charts.add('bar',{position:{left:80,top:215,width:750,height:365},categories:['May','June'],series:[{name:'Gross margin %',values:[40,38.3333333333],fill:'#176B74'}],barOptions:{direction:'column',grouping:'clustered'},hasLegend:false});
box(s,'EUR 46,000',870,255,330,70,38,'#176B74',true);box(s,'June gross profit',870,325,330,50,25);box(s,'South B margin: 32%\nCost cause is unproven.',870,420,330,130,26);
s.speakerNotes.textFrame.setText('Source: sales.csv. Gross profit = revenue - cost. June: 120000 - 74000 = 46000. May margin 40%; June 38.33%; South B (25000 - 17000) / 25000 = 32%.');
s=slide(deck,'Two segments account for the budget gap',4);
box(s,'North  EUR 70,000\nSouth  EUR 50,000',70,235,520,150,36);box(s,'Product A  EUR 65,000\nProduct B  EUR 55,000',670,235,540,150,36);
box(s,'North A: EUR 5,000 below budget\nSouth B: EUR 5,000 below budget',70,445,1120,135,34,'#176B74',true);
s.speakerNotes.textFrame.setText('Sources: sales.csv and budget.csv. Grouped June revenue and row-level actual minus budget.');
s=slide(deck,'Delivery performance is below target',5);
box(s,'94%',65,240,340,110,76,'#176B74',true);box(s,'June on-time rate',65,370,370,60,26);box(s,'97% target\n3 percentage point gap\n94 of 100 orders',555,250,650,200,36);
box(s,'May was 96%. Use the current June result for this review.',65,535,1110,60,26);
s.speakerNotes.textFrame.setText('Source OPS-JUN-FINAL: Atlas-operations-note.pdf and current operations.html. Historical archive is May, not June.');
s=slide(deck,'Owners have two planned follow-ups',6);
box(s,'Mira Cole',65,225,360,55,34,'#176B74',true);box(s,'Review South B freight costs\nDue 10 July 2026',455,225,700,115,30);
box(s,'Leon Park',65,405,360,55,34,'#176B74',true);box(s,'Confirm supplier recovery plan\nDue 12 July 2026',455,405,700,115,30);
box(s,'Planned actions. No completed savings are claimed.',65,585,1120,45,25);
s.speakerNotes.textFrame.setText('Source: Atlas-operations-note.pdf. Planned actions, owners and due dates verbatim.');
for(const sl of deck.slides.items) for(const c of sl.charts.items){
 c.xAxis={visible:true,textStyle:{typeface:'Arial',fontSize:22,fill:'#18344D'}};
 c.yAxis={visible:true,min:0,textStyle:{typeface:'Arial',fontSize:18,fill:'#62717C'}};
 c.dataLabels={showValue:true,position:'outEnd',textStyle:{typeface:'Arial',fontSize:22,fill:'#18344D',bold:true}};
}
await (await PresentationFile.exportPptx(deck)).save(path.join(kit,'evaluator-only/reference-deck.pptx'));
await fs.mkdir(path.join(kit,'evaluator-only/reference-deck-slides'),{recursive:true});
let i=0;for(const sl of deck.slides.items){i++;const img=await deck.export({slide:sl,format:'png',scale:1});await fs.writeFile(path.join(kit,'evaluator-only/reference-deck-slides',`slide-${i}.png`),new Uint8Array(await img.arrayBuffer()));}
const starter=Presentation.create({slideSize:{width:1280,height:720}});s=slide(starter,'Atlas June review source brief',1);box(s,'Create draft slides for the executive review.\nPreserve this source slide.\nUse final finance and operations evidence.',65,240,1110,260,32);s.speakerNotes.textFrame.setText('SOURCE: Atlas-review-brief.docx. Preserve this original slide.');await (await PresentationFile.exportPptx(starter)).save(path.join(kit,'inputs/powerpoint/Atlas-start.pptx'));
const crowded=Presentation.create({slideSize:{width:1280,height:720}});
for(let j=0;j<6;j++){s=slide(crowded,'[Scribble draft] June review '+(j+1),j+1);box(s,'Revenue EUR 120000. Budget EUR 130000. Gross margin 38.33%. Delivery 94%. '.repeat(8),60,220,700,120,34);s.speakerNotes.textFrame.setText('Sources: sales.csv; budget.csv; Atlas-operations-note.pdf. Deliberate overflow fixture.');if(j===0)s.charts.add('bar',{position:{left:860,top:240,width:260,height:160},categories:['Actual','Budget'],series:[{name:'EUR',values:[120000,130000]}],hasLegend:false,xAxis:{textStyle:{typeface:'Arial',fontSize:8}}});}
await (await PresentationFile.exportPptx(crowded)).save(path.join(kit,'inputs/powerpoint/Atlas-crowded.pptx'));
console.log('Generated four workbooks and three presentations.');
