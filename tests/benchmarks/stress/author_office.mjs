// Authoring backend for generate_office.py. Uses the bundled Artifact Tool.
// All build JSON, PNGs, references and validation receipts are evaluator-only.
import fs from 'node:fs/promises';
import path from 'node:path';
import {pathToFileURL} from 'node:url';
import {Workbook, SpreadsheetFile, Presentation, PresentationFile} from '@oai/artifact-tool';

const root=process.env.OFFICE_STRESS_ROOT;
const skill=process.env.PRESENTATIONS_SKILL;
if(!path.isAbsolute(root??'') || !path.isAbsolute(skill??'')) throw new Error('Absolute OFFICE_STRESS_ROOT and PRESENTATIONS_SKILL required.');
const build=path.join(root,'.build');
const payload=JSON.parse(await fs.readFile(path.join(build,'office_payload.json'),'utf8'));
const selected=new Set((process.env.OFFICE_STRESS_ONLY??'').split(',').filter(Boolean));
const previews=process.env.OFFICE_STRESS_PREVIEWS!=='0';
const {resolvePresentationFont,applyPresentationChartFont,finalizePresentation}=await import(pathToFileURL(path.join(skill,'container_tools/artifact_tool_utils.mjs')).href);
const font=resolvePresentationFont({fontFamily:'Arial'});
const BLUE='#4F81BD', SOFT='#5B9BD5', BORDER='#41719C', INK='#202A35', MUTED='#596674', GRAY='#F2F2F2';
const hydrate=v=>v && typeof v==='object' && v.$date ? new Date(v.$date+'T00:00:00Z') : v;
const col=n=>{let s='';while(n>0){n--;s=String.fromCharCode(65+n%26)+s;n=Math.floor(n/26);}return s;};
const display=v=>v==null?'Unknown':Number.isInteger(v)?v.toLocaleString('en-US'):v.toLocaleString('en-US',{maximumFractionDigits:2});
const percent=v=>v==null?'Unknown':(v*100).toFixed(2)+'%';
await fs.mkdir(path.join(root,'inputs','excel'),{recursive:true});
await fs.mkdir(path.join(root,'inputs','powerpoint'),{recursive:true});
await fs.mkdir(path.join(root,'evaluator-only','powerpoint-references'),{recursive:true});
const receipts=[];

function formatSheet(sheet,rows,columns){
  sheet.showGridLines=false;
  const used=sheet.getRangeByIndexes(0,0,rows,columns);
  used.format.font={name:'Arial',size:11,color:INK};
  used.format.columnWidth=18;used.format.rowHeight=22;
  sheet.getRangeByIndexes(0,0,1,columns).format={fill:BLUE,font:{name:'Arial',bold:true,color:'#FFFFFF',size:11},rowHeight:32,wrapText:true};
  sheet.freezePanes.freezeRows(1);
}

async function authorWorkbook(spec){
  const wb=Workbook.create();
  const sheet=wb.worksheets.add('Ledger');
  const rows=[spec.columns,...spec.rows.map(row=>row.map(hydrate))];
  sheet.getRangeByIndexes(0,0,rows.length,spec.columns.length).values=rows;
  formatSheet(sheet,rows.length,spec.columns.length);
  sheet.getRange(`A2:B${rows.length}`).setNumberFormat('@');
  sheet.getRange(`C2:C${rows.length}`).setNumberFormat('yyyy-mm-dd');
  for(const field of spec.columns){
    const letter=spec.header_to_column[field];
    if(field.includes('EUR')) sheet.getRange(`${letter}2:${letter}${rows.length}`).setNumberFormat('#,##0.00;[Red](#,##0.00);"-"');
    if(['ReturnRate','Utilization'].includes(field)) sheet.getRange(`${letter}2:${letter}${rows.length}`).setNumberFormat('0.00%');
    if(field==='DueDate') sheet.getRange(`${letter}2:${letter}${rows.length}`).setNumberFormat('yyyy-mm-dd');
  }
  for(const f of spec.formulas) sheet.getRange(f.cell).formulas=[[f.formula]];
  sheet.tables.add(`A1:${col(spec.columns.length)}${rows.length}`,true,spec.id+'Ledger');
  const history=wb.worksheets.add('History');
  const hrows=[['Period',spec.primary_label,spec.secondary_label],...spec.historical_months.map(m=>[m,null,null])];
  history.getRange('A1:C6').values=hrows;formatSheet(history,6,3);history.getRange('A2:A6').setNumberFormat('@');
  for(let i=0;i<5;i++){
    const row=i+2;
    for(const [target,field] of [['B',spec.primary_field],['C',spec.secondary_field]]){
      const source=spec.header_to_column[field];
      history.getRange(target+row).formulas=[[`=SUMIF(Ledger!$B$2:$B$${rows.length},A${row},Ledger!$${source}$2:$${source}$${rows.length})`]];
    }
  }
  history.getRange('B2:C6').setNumberFormat('#,##0.00');
  history.getRange('A9:C11').merge();
  history.getRange('A9').values=[['Historical management view through May 2026. Ledger includes June source records. Blank source values mean unknown, not zero. Duplicate RowID denotes a repeated source record.']];
  history.getRange('A9:C11').format={wrapText:true,font:{name:'Arial',size:10,color:MUTED}};
  // Two series share a scale only where their units are comparable.
  const comparable=spec.domain!=='inventory';
  const chart=history.charts.add('line',history.getRange(comparable?'A1:C6':'A1:B6'));
  chart.title=spec.primary_label+' — Jan–May 2026';chart.titleTextStyle.typeface='Arial';chart.hasLegend=comparable;
  chart.xAxis={axisType:'textAxis',textStyle:{typeface:'Arial',fontSize:12}};
  chart.yAxis={min:0,numberFormatCode:'#,##0',numberFormatSourceLinked:false,textStyle:{typeface:'Arial',fontSize:12}};
  chart.setPosition('E2','N19');
  for(const [index,series] of chart.series.items.entries()) series.line={fill:index?BORDER:SOFT,style:index?'dashed':'solid',width:2};
  wb.recalculate();
  const out=path.join(build,spec.id+'.xlsx');
  await (await SpreadsheetFile.exportXlsx(wb)).save(out);
  await fs.copyFile(out,path.join(root,spec.path));
  const qa=path.join(root,'evaluator-only','previews',spec.id);await fs.mkdir(qa,{recursive:true});
  const cache=await wb.inspect({kind:'table',range:'History!A1:C6',include:'values,formulas',tableMaxRows:6,tableMaxCols:3});
  await fs.writeFile(path.join(qa,'history-readback.json'),typeof cache==='string'?cache:JSON.stringify(cache,null,2));
  if(previews){
    for(const [sheetName,range] of [['Ledger',`A1:${col(spec.columns.length)}16`],['History','A1:N20']]){
      const image=await wb.render({sheetName,range,scale:1,format:'png'});
      await fs.writeFile(path.join(qa,sheetName+'.png'),new Uint8Array(await image.arrayBuffer()));
    }
  }
  console.log(JSON.stringify({type:'workbook',id:spec.id,path:spec.path,rows:spec.rows.length}));
}

function textbox(slide,text,left,top,width,height,size=24,bold=false,color=INK){
  const shape=slide.shapes.add({geometry:'textbox',position:{left,top,width,height},fill:'none',line:{fill:'none',width:0}});
  shape.text=text;shape.text.style={typeface:font,fontSize:size,bold,color,autoFit:'none'};return shape;
}

function newSlide(deck,spec,title,n,defective=false){
  const s=deck.slides.add();s.background.fill='#FFFFFF';
  textbox(s,title,78,38,1152,72,32,true,INK);
  textbox(s,`${spec.company} | ${spec.domain} | January–June 2026`,49,116,1181,31,defective?18:19,false,BLUE);
  textbox(s,`${spec.linked_workbook} • Fictional operational source • ${spec.id}`,49,673,1114,25,13,false,MUTED);
  textbox(s,String(n),1203,676,57,34,13,false,MUTED);
  s.speakerNotes.textFrame.setText(`Fictional source: ${spec.linked_workbook}, Ledger and monthly operations records, January–June 2026. ${spec.current.complete?'Source period complete.':'June contains one unknown source amount; displayed amount is a known subtotal, not a complete total.'} Values in this source are not claims about a real company.`);
  return s;
}

function nativeTable(slide,values,{left=84,top=204,width=1112,height=360,defective=false}={}){
  const table=slide.tables.add({rows:values.length,columns:values[0].length,left,top,width,height,values,columnWidths:values[0].map((_,i)=>i===0?width*.28:width*.72/(values[0].length-1))});
  table.borders.assign({style:'solid',fill:'#D7DDE3',width:.65});
  for(let r=0;r<values.length;r++) for(let c=0;c<values[0].length;c++){
    const cell=table.getCell(r,c);cell.fill=r===0?(defective?'#E91E63':BLUE):(r%2===0?GRAY:'#FFFFFF');
    cell.text.style={typeface:font,fontSize:18,bold:r===0,color:r===0?'#FFFFFF':INK};
  }
  return table;
}

function nativeChart(slide,spec,{kind='bar',grouped=false,defective=false}={}){
  const categories=grouped?spec.by_group.map(x=>x.group):spec.monthly.map(x=>x.period);
  if(defective&&!grouped) categories[categories.length-1]='2025-06';
  const values=grouped?spec.by_group:spec.monthly;
  const compatible=spec.domain!=='inventory';
  const series=[{name:spec.primary_label,values:values.map(x=>x.primary),fill:BLUE}];
  if(compatible) series.push({name:spec.secondary_label,values:values.map(x=>x.secondary),fill:SOFT});
  if(kind==='line') for(const [i,value] of series.entries()) value.line={fill:i?SOFT:BLUE,style:'solid',width:2};
  const unit=spec.domain==='inventory'?'units':spec.domain==='workforce'||spec.domain==='projects'?'hours':'EUR';
  const ch=slide.charts.add(kind,{position:defective?{left:1080,top:211,width:680,height:371}:{left:88,top:211,width:1100,height:371},title:defective?'Historical trend — 2025':`${spec.primary_label} / ${spec.secondary_label} (${unit})`,titleTextStyle:{typeface:font,fontSize:18},categories,series,barOptions:{direction:'column',grouping:'clustered'},lineOptions:{grouping:'standard',smooth:false},hasLegend:series.length>1});
  applyPresentationChartFont(ch,{fontFamily:font});
  ch.xAxis={visible:true,textStyle:{typeface:font,fontSize:18,fill:INK}};
  ch.yAxis={visible:true,min:0,numberFormatCode:'#,##0',textStyle:{typeface:font,fontSize:17,fill:MUTED}};
  if(!compatible) textbox(slide,spec.primary_label+' only; secondary metric uses a different unit.',88,604,1100,32,18,false,MUTED);
  if(!spec.current.complete) textbox(slide,'June: known subtotal; one source value is unknown.',88,636,1100,34,18,true,'#C00000');
  return ch;
}

function buildDeck(spec,defective){
  const deck=Presentation.create({slideSize:{width:1280,height:720}});
  let s=newSlide(deck,spec,`${spec.company}: ${spec.domain} review`,1,defective);
  textbox(s,`${spec.primary_label}\n${display(spec.current.primary)}${spec.current.complete?'':' (known subtotal)'}`,85,220,600,144,36,true,BLUE);
  textbox(s,`${spec.secondary_label}\n${display(spec.current.secondary)}`,735,220,480,144,defective?30:36,false,INK);
  textbox(s,`Six months of operating records across four groups.\n${spec.ratio_label?spec.ratio_label+': '+percent(spec.current.ratio):'Inventory balances and valuation use separate units.'}`,85,440,1115,110,25);
  s=newSlide(deck,spec,`${spec.primary_label}: six-month trend`,2,defective);nativeChart(s,spec,{defective});
  s=newSlide(deck,spec,'June results by operating group',3,defective);
  nativeTable(s,[['Group',spec.primary_label,spec.secondary_label],...spec.by_group.map(x=>[x.group,display(x.primary)+(x.complete?'':'*'),display(x.secondary)])],{defective});
  if(!spec.current.complete) textbox(s,'* Known subtotal: source includes an unknown amount.',84,588,1110,38,18,false,'#C00000');
  s=newSlide(deck,spec,'Operating review and evidence boundaries',4,defective);
  const commentary=`June ${spec.primary_label.toLowerCase()}: ${display(spec.current.primary)}${spec.current.complete?'.':' — known subtotal only.'}\n\n${spec.secondary_label}: ${display(spec.current.secondary)}. ${spec.ratio_label?spec.ratio_label+': '+percent(spec.current.ratio)+'.':''}\n\nThe monthly comparison covers January to June 2026. Group comparisons use June records only. Source blanks remain unknown.\n\nOperational status and financial performance are separate: a planned action is not a completed result.`;
  if(defective) textbox(s,commentary,84,204,330,84,24,false,INK);
  else {
    const measureText=`June ${spec.primary_label.toLowerCase()}: ${display(spec.current.primary)}${spec.current.complete?'.':' — known subtotal only.'}\n\n${spec.secondary_label}: ${display(spec.current.secondary)}. ${spec.ratio_label?spec.ratio_label+': '+percent(spec.current.ratio)+'.':''}`;
    const boundaryText='The monthly comparison covers January to June 2026. Group comparisons use June records only. Source blanks remain unknown.\n\nOperational status and financial performance are separate: a planned action is not a completed result.';
    for(const x of [84,658]) s.shapes.add({geometry:'textbox',position:{left:x,top:204,width:538,height:370},fill:GRAY,line:{fill:'none',width:0}});
    textbox(s,'June measures',108,227,490,38,25,true,BLUE);
    textbox(s,`${display(spec.current.primary)} ${spec.domain==='inventory'?'units':spec.domain==='workforce'||spec.domain==='projects'?'hours':'EUR'}`,108,270,490,54,32,true,INK);
    textbox(s,measureText,108,330,490,220,22,false,INK);
    textbox(s,'Evidence boundaries',682,227,490,38,25,true,BLUE);
    textbox(s,'January–June 2026',682,270,490,54,28,true,INK);
    textbox(s,boundaryText,682,330,490,220,21,false,INK);
  }
  s=newSlide(deck,spec,'Follow-up register',5,defective);
  nativeTable(s,[['Owner','Proposed follow-up','Due date'],['Mira Cole','Reconcile North source records','2026-07-10'],['Leon Park','Review South exceptions','2026-07-12'],['Nadia Shah','Confirm East operating assumptions','2026-07-14'],['Evan Reed','Review West source completeness','2026-07-16']],{height:340});
  textbox(s,'All follow-ups are planned. Completion and savings are not confirmed.',84,589,1110,43,20,false,MUTED);
  s=newSlide(deck,spec,'Source coverage and definitions',6,defective);
  textbox(s,`Workbook: ${spec.linked_workbook}\nPeriod: January–June 2026\nGroups: North, South, East, West\nMeasures: ${spec.primary_label}; ${spec.secondary_label}\n${spec.current.complete?'June source measures are complete.':'June contains an unknown source value. Report the known subtotal.'}\nRates use the aggregate numerator and denominator, not an average of row rates.`,84,207,1110,370,25);
  for(let n=7;n<=spec.slide_count;n++){
    if(n===7){s=newSlide(deck,spec,'June group comparison',n,defective);nativeChart(s,spec,{grouped:true});}
    if(n===8){s=newSlide(deck,spec,'Monthly operating detail',n,defective);nativeTable(s,[['Month',spec.primary_label,spec.secondary_label],...spec.monthly.map(x=>[x.period,display(x.primary)+(x.complete?'':'*'),display(x.secondary)])],{top:201,height:405});}
    if(n===9){s=newSlide(deck,spec,'Interpretation limits',n,defective);textbox(s,'Period comparisons describe changes; they do not establish a cause.\n\nRead the linked workbook for transaction-level evidence.\n\nKeep missing values visible and identify the period when discussing a rate.\n\nPreserve native editable charts and tables in any revised deck.',84,206,1108,390,27);}
    if(n===10){s=newSlide(deck,spec,'Trend at a second reading',n,defective);nativeChart(s,spec,{kind:'line'});}
    if(n===11){s=newSlide(deck,spec,'Source navigation',n,defective);nativeTable(s,[['Source','Coverage','Purpose'],[spec.linked_workbook,'Ledger','Transaction or operating record detail'],[spec.linked_workbook,'History','Historical view through May'],[spec.id,'Slides 2–3','Current period and group comparison'],[spec.id,'Slide 5','Planned follow-ups, not completed outcomes']],{height:330});}
    if(n===12){s=newSlide(deck,spec,'Discussion and next actions',n,defective);textbox(s,'Confirm the scope of the requested decision.\n\nUse current-period values with their correct labels and units.\n\nResolve source gaps before asserting a complete result.\n\nTrack owners and due dates without implying that planned work is finished.',84,206,1108,390,27);}
  }
  return deck;
}

async function finishDeck(spec,defective){
  const deck=buildDeck(spec,defective);
  const staging=path.join(build,spec.id+(defective?'-repair':'-reference'));
  await fs.mkdir(staging,{recursive:true});
  const candidatePath=path.join(staging,'candidate.pptx');
  await(await PresentationFile.exportPptx(deck)).save(candidatePath);
  const finalPath=path.join(staging,'validated.pptx');
  // The finalizer refuses overwrite. Reruns write a fresh private destination.
  const revision=String(Date.now());
  const validated=finalPath.replace('.pptx','-'+revision+'.pptx');
  const chartOwners=[2,...(spec.slide_count>=7?[7]:[]),...(spec.slide_count>=10?[10]:[])];
  const tableOwners=[3,5,...(spec.slide_count>=8?[8]:[]),...(spec.slide_count>=11?[11]:[])];
  const receipt=path.join(build,'receipts',spec.id+(defective?'-repair-':'-reference-')+revision+'.json');
  await fs.mkdir(path.dirname(receipt),{recursive:true});
  const result=await finalizePresentation({workspaceDir:root,candidatePath,finalPath:validated,pythonExecutable:process.env.RUNTIME_PYTHON,integrityValidatorPath:path.join(skill,'container_tools/inspect_presentation_package_integrity.py'),layoutValidatorPath:path.join(skill,'container_tools/inspect_presentation_layout_geometry.py'),layoutArgs:['--expected-slide-size-emu','12192000,6858000',...(!defective?['--validate-heading-fit']:[]),...tableOwners.flatMap(n=>['--require-native-table-slide',String(n)])],explicitTotalSlideCount:spec.slide_count,requiredNativeChartOwnerSlides:chartOwners,requiredNativeTableOwnerSlides:tableOwners,materializeLiteralChartWorkbooks:true,nativeChartTargetApplication:'portable',fontPolicy:{basis:'design',families:[font]},verifyArtifactToolImport:true,receiptPath:receipt});
  // Deliberate repair sources have documented bad geometry. Package/native
  // integrity still applies; clean references pass the full layout checks.
  const destination=path.join(root,defective?spec.path:spec.reference_path);
  await fs.copyFile(validated,destination);
  if(!defective && !spec.repair_case) await fs.copyFile(validated,path.join(root,spec.path));
  const qa=path.join(root,'evaluator-only','previews',spec.id,defective?'source':'reference');
  await fs.mkdir(qa,{recursive:true});
  if(previews){
    let n=0;for(const sl of deck.slides.items){n++;const img=await deck.export({slide:sl,format:'png',scale:1});await fs.writeFile(path.join(qa,`slide-${String(n).padStart(2,'0')}.png`),new Uint8Array(await img.arrayBuffer()));}
  }
  receipts.push({id:spec.id,kind:defective?'intentional-repair-source':'clean-reference',slides:spec.slide_count,receipt:path.relative(root,receipt),visual_review_complete:false,native_office_verified:false});
  console.log(JSON.stringify({type:'presentation',id:spec.id,defective,slides:spec.slide_count,destination:path.relative(root,destination)}));
}

for(const spec of payload.workbooks) if(!selected.size||selected.has(spec.id)) await authorWorkbook(spec);
for(const spec of payload.presentations) if(!selected.size||selected.has(spec.id)){
  await finishDeck(spec,false);
  if(spec.repair_case) await finishDeck(spec,true);
}
// A partial repair build must not erase receipts for unchanged source files.
const receiptNames=(await fs.readdir(path.join(build,'receipts'))).sort();
const allReceipts=[];
for(const spec of payload.presentations) for(const defective of [false,...(spec.repair_case?[true]:[])]){
  const prefix=spec.id+(defective?'-repair-':'-reference-');
  const latest=receiptNames.filter(x=>x.startsWith(prefix)&&x.endsWith('.json')).at(-1);
  if(latest) allReceipts.push({id:spec.id,kind:defective?'intentional-repair-source':'clean-reference',slides:spec.slide_count,receipt:path.relative(root,path.join(build,'receipts',latest)),visual_review_complete:false,native_office_verified:false});
}
await fs.writeFile(path.join(root,'evaluator-only','office_authoring_receipts.json'),JSON.stringify({schema_version:1,scope:'Artifact Tool and package verification; not native/model acceptance',receipts:allReceipts},null,2)+'\n');
