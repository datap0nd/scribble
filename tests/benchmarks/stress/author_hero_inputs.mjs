// Deterministic high-value benchmark workbooks. Model-visible files contain
// sources only; exact translations and table matrices stay evaluator-only.
import fs from 'node:fs/promises';
import path from 'node:path';
import {Workbook, SpreadsheetFile} from '@oai/artifact-tool';

const root=process.env.OFFICE_STRESS_ROOT;
if(!path.isAbsolute(root??'')) throw new Error('Absolute OFFICE_STRESS_ROOT required.');
const BLUE='#1428A0', CYAN='#00A9E0', INK='#111827', PALE='#EEF2FF', BORDER='#CBD5E1';
const koreanSheets=[
  {name:'Operations',rows:[
    [['구분','Category'],['담당자','Owner'],['상태','Status'],['마감일','Due date']],
    [['물류','Logistics'],['김민준','Min-jun Kim'],['완료','Complete'],['2026년 9월 20일','2026-09-20']],
    [['품질','Quality'],['이서연','Seo-yeon Lee'],['진행 중','In progress'],['2026년 9월 22일','2026-09-22']],
    [['판매','Sales'],['박지훈','Ji-hoon Park'],['검토 필요','Review required'],['2026년 9월 25일','2026-09-25']],
    [['고객 지원','Customer support'],['최유진','Yu-jin Choi'],['대기 중','Pending'],['2026년 9월 28일','2026-09-28']]]},
  {name:'Finance',rows:[
    [['항목','Item'],['부서','Department'],['금액','Amount'],['비고','Notes']],
    [['매출','Revenue'],['모바일 사업부','Mobile division'],['1250000','1250000'],['확정','Confirmed']],
    [['비용','Cost'],['운영팀','Operations team'],['820000','820000'],['잠정','Provisional']],
    [['예산','Budget'],['재무팀','Finance team'],['900000','900000'],['승인됨','Approved']],
    [['절감액','Savings'],['조달팀','Procurement team'],['45000','45000'],['검증 필요','Verification required']]]},
  {name:'People',rows:[
    [['이름','Name'],['직책','Role'],['지역','Region'],['근무 형태','Work arrangement']],
    [['정하늘','Ha-neul Jeong'],['프로젝트 관리자','Project manager'],['서울','Seoul'],['사무실 근무','Office-based']],
    [['오지민','Ji-min Oh'],['데이터 분석가','Data analyst'],['부산','Busan'],['재택근무','Remote']],
    [['윤서준','Seo-jun Yoon'],['품질 책임자','Quality lead'],['수원','Suwon'],['혼합 근무','Hybrid']],
    [['한예린','Ye-rin Han'],['재무 담당자','Finance specialist'],['인천','Incheon'],['사무실 근무','Office-based']]]},
  {name:'Risks',rows:[
    [['위험','Risk'],['영향','Impact'],['가능성','Likelihood'],['대응 방안','Mitigation']],
    [['납품 지연','Delivery delay'],['높음','High'],['중간','Medium'],['대체 공급업체 확인','Confirm alternate supplier']],
    [['환율 변동','Exchange-rate volatility'],['중간','Medium'],['높음','High'],['주간 환율 검토','Review exchange rate weekly']],
    [['품질 문제','Quality issue'],['높음','High'],['낮음','Low'],['추가 검사 실시','Perform additional inspection']],
    [['인력 부족','Staff shortage'],['중간','Medium'],['중간','Medium'],['임시 인력 확보','Secure temporary staff']]]}
];

const wordTables=[
  {name:'Revenue',rows:[['Month','Product','Region','Revenue (AED)'],['2026-06','Fold','UAE','420000'],['2026-07','Flip','UAE','365000'],['2026-08','Fold','KSA','448000'],['2026-09','Flip','KSA','392000']]},
  {name:'Costs',rows:[['Cost center','Owner','Budget (AED)','Actual (AED)'],['Marketing','Amina Rahman','180000','172500'],['Logistics','Omar Saleh','145000','151200'],['Retail','Leila Haddad','210000','205400'],['Support','Daniel Kim','98000','96300']]},
  {name:'Pipeline',rows:[['Opportunity','Stage','Probability','Value (AED)'],['Enterprise A','Proposal','70%','310000'],['Retail B','Negotiation','85%','225000'],['Carrier C','Discovery','40%','480000'],['Online D','Contract','95%','155000']]},
  {name:'Headcount',rows:[['Team','Plan','Actual','Open roles'],['Sales','24','22','2'],['Operations','18','17','1'],['Finance','9','9','0'],['Support','26','23','3']]},
  {name:'Inventory',rows:[['SKU','Opening units','Shipped units','Closing units'],['FOLD-256-BLK','420','175','245'],['FOLD-512-SLV','310','128','182'],['FLIP-256-BLU','515','204','311'],['FLIP-512-GLD','275','96','179']]},
  {name:'Projects',rows:[['Project','Owner','Status','Due date'],['Store refresh','Noura Ali','On track','2026-09-30'],['Trade-in portal','Yusuf Khan','At risk','2026-10-08'],['Care training','Sara Lim','On track','2026-10-15'],['Forecast model','Ibrahim Noor','Blocked','2026-10-20']]},
  {name:'Risks',rows:[['Risk ID','Risk','Severity','Mitigation'],['R-101','Carrier delay','High','Add backup carrier'],['R-102','FX movement','Medium','Weekly hedge review'],['R-103','Portal outage','High','Complete failover drill'],['R-104','Training gap','Low','Add two workshops']]},
  {name:'Actions',rows:[['Action ID','Action','Owner','Target date'],['A-201','Approve backup carrier','Omar Saleh','2026-09-21'],['A-202','Validate portal failover','Yusuf Khan','2026-09-24'],['A-203','Confirm store staffing','Noura Ali','2026-09-26'],['A-204','Publish forecast baseline','Ibrahim Noor','2026-09-29']]}
];

function style(sheet,rowCount,columnCount){
  sheet.showGridLines=false;
  const used=sheet.getRangeByIndexes(0,0,rowCount,columnCount);
  used.format={font:{name:'Arial',size:11,color:INK},rowHeight:23,verticalAlignment:'center'};
  used.format.columnWidth=24;
  sheet.getRangeByIndexes(0,0,1,columnCount).format={fill:BLUE,font:{name:'Arial',size:11,bold:true,color:'#FFFFFF'},rowHeight:31,wrapText:true};
  for(let r=1;r<rowCount;r++) if(r%2===0) sheet.getRangeByIndexes(r,0,1,columnCount).format.fill=PALE;
  used.format.borders={top:{style:'continuous',color:BORDER},bottom:{style:'continuous',color:BORDER},left:{style:'continuous',color:BORDER},right:{style:'continuous',color:BORDER},insideHorizontal:{style:'continuous',color:BORDER},insideVertical:{style:'continuous',color:BORDER}};
  sheet.freezePanes.freezeRows(1);
}

async function saveWorkbook(fileName,specs,selector){
  const wb=Workbook.create();
  for(const spec of specs){
    const values=selector(spec);
    const sheet=wb.worksheets.add(spec.name);
    sheet.getRangeByIndexes(0,0,values.length,values[0].length).values=values;
    style(sheet,values.length,values[0].length);
    sheet.tables.add(`A1:D${values.length}`,true,`Tbl${spec.name.replace(/[^A-Za-z0-9]/g,'')}`);
  }
  const target=path.join(root,'inputs','excel',fileName);
  await fs.mkdir(path.dirname(target),{recursive:true});
  await (await SpreadsheetFile.exportXlsx(wb)).save(target);
  const qa=path.join(root,'evaluator-only','previews',path.parse(fileName).name);await fs.mkdir(qa,{recursive:true});
  for(const spec of specs){
    const values=selector(spec);
    const image=await wb.render({sheetName:spec.name,range:`A1:D${values.length}`,scale:1.5,format:'png'});
    await fs.writeFile(path.join(qa,spec.name+'.png'),new Uint8Array(await image.arrayBuffer()));
  }
}

await saveWorkbook('KoreanOperations.xlsx',koreanSheets,s=>s.rows.map(row=>row.map(pair=>pair[0])));
await saveWorkbook('EightSheetWorkbook.xlsx',wordTables,s=>s.rows);
const translated=koreanSheets.map(sheet=>({name:sheet.name,cells:sheet.rows.flatMap((row,r)=>row.map((pair,c)=>({row:r+1,column:c+1,text:pair[1]})))}));
await fs.writeFile(path.join(root,'evaluator-only','hero_inputs.json'),JSON.stringify({schema:1,korean:{path:'inputs/excel/KoreanOperations.xlsx',sheets:translated},word:{path:'inputs/excel/EightSheetWorkbook.xlsx',tables:wordTables.map(x=>x.rows)}},null,2)+'\n');
console.log(JSON.stringify({workbooks:2,korean_cells:translated.reduce((n,s)=>n+s.cells.length,0),word_tables:wordTables.length}));
