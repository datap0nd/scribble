"""Offline evidence inspection. Never executes model tools, formulas or macros.
Machine findings are combined with an explicit native/visual review, never inferred from file presence.
"""
import argparse, hashlib, json, re, zipfile
from pathlib import Path, PurePosixPath
from xml.etree import ElementTree as ET
from io import BytesIO

def digest(data):return hashlib.sha256(data).hexdigest()
def read_archive(path):
    with zipfile.ZipFile(path) as z:
        seen=set();size=0;out={}
        for item in z.infolist():
            p=PurePosixPath(item.filename)
            if p.is_absolute() or '..' in p.parts or '\\' in item.filename or item.filename in seen:raise ValueError('Unsafe/duplicate archive entry')
            seen.add(item.filename);size+=item.file_size
            if size>800*1024*1024:raise ValueError('Archive expands beyond 800 MB')
            if not item.is_dir():out[item.filename]=z.read(item)
        return out

def inspect_office(data,ext):
    with zipfile.ZipFile(BytesIO(data)) as z:
        if sum(i.file_size for i in z.infolist())>100*1024*1024:raise ValueError('Oversized Office package')
        names=z.namelist()
        if any('vbaProject' in n for n in names):raise ValueError('Unexpected macro payload')
        xml={n:ET.fromstring(z.read(n)) for n in names if n.endswith('.xml') and (n.startswith('xl/worksheets/') or re.match(r'ppt/(slides/slide\d+|notesSlides/notesSlide\d+|(?:slides/)?charts/chart\d+)\.xml$',n) or n=='word/document.xml')}
        text='\n'.join(' '.join(n.itertext()) for n in xml.values())
        formulas=sum(len(n.findall('.//{*}f')) for name,n in xml.items() if name.startswith('xl/worksheets/'))
        formula_errors=[v.text for n in xml.values() for c in n.findall('.//{*}c') if c.get('t')=='e' for v in c.findall('{*}v')]
        numbers=[]
        for token in re.findall(r'(?<![A-Za-z])[-+−]?\d[\d,]*(?:\.\d+)?',text):
            try:numbers.append(float(token.replace(',','').replace('−','-')))
            except ValueError:pass
        if ext=='xlsx':
            for name in names:
                if name=='xl/sharedStrings.xml':text+='\n'+' '.join(ET.fromstring(z.read(name)).itertext())
        charts=[n for n in names if re.match(r'(?:xl/(?:drawings/)?|ppt/(?:slides/)?)charts/chart\d+\.xml$',n)]
        slides=[n for n in names if re.match(r'ppt/slides/slide\d+\.xml$',n)]
        notes=[' '.join(v.itertext()) for n,v in xml.items() if n.startswith('ppt/notesSlides/')]
        return {'text':text,'numbers':numbers,'formula_count':formulas,'formula_errors':formula_errors,'chart_count':len(charts),'slide_count':len(slides),'notes_count':len(notes),'substantive_notes_count':sum(len(n.strip())>=30 for n in notes)}

def evaluate(run_zip,kit,review=None):
    files=read_archive(run_zip);run=json.loads(files['run.json']);manifest=json.loads(files['export-manifest.json']);checks=[]
    def check(name,passed,detail='',hard=True):checks.append({'name':name,'passed':bool(passed),'detail':detail,'hard':hard})
    check('export_hashes',all(f['path'] in files and len(files[f['path']])==f['size'] and digest(files[f['path']])==f['sha256'] for f in manifest['files']))
    listed={f['path'] for f in manifest['files']};check('export_inventory',set(files)==listed|{'export-manifest.json'})
    check('fixture_version',digest((kit/'manifest.json').read_bytes())==run['manifest_sha256'])
    cases=json.loads((kit/'operator/cases.json').read_text(encoding='utf-8'));case=next(c for c in cases if c['id']==run['case_id'])
    check('trace_complete',run.get('trace_complete') and run.get('status')=='finished')
    check('unassisted',not run.get('assisted'))
    events=[json.loads(x) for x in files['timeline.jsonl'].decode('utf-8-sig').splitlines() if x.strip()]
    check('events_present',bool(events));check('run_correlation',all(e.get('run_id')==run['run_id'] for e in events))
    keys=[(e.get('instance_id'),e.get('sequence')) for e in events];check('unique_event_sequences',len(keys)==len(set(keys)))
    requests=[e for e in events if e['stage']=='inference_request'];responses=[e for e in events if e['stage']=='inference_response']
    check('inference_recorded',bool(requests) and bool(responses));check('lifecycle_finished',any(e['stage']=='case_finished' for e in events))
    request_text=json.dumps(requests);check('no_oracle_paths',not any(s in request_text.lower() for s in ['evaluator-only','reference-deck.pptx','expected-analysis.xlsx','answers.json']))
    reasoning=False;models=set();response_text=[]
    for e in responses:
        d=e.get('detail',{});models.add(d.get('model','unknown'))
        try:
            response=json.loads(d.get('response','{}'));msg=response.get('choices',[{}])[0].get('message',{})
            reasoning=reasoning or bool(msg.get('reasoning_content') or msg.get('reasoning'))
            response_text.append(str(msg.get('content') or ''))
        except (ValueError,IndexError,TypeError):pass
    artifacts={};available=set()
    for name,data in files.items():
        if not name.startswith('artifacts/') or name.endswith('.receipt.json'):continue
        ext=Path(name).suffix.lstrip('.');available.add(ext)
        if ext in ['xlsx','pptx','docx']:
            try:artifacts[name]=inspect_office(data,ext)
            except (ValueError,ET.ParseError,zipfile.BadZipFile) as e:check('readable_'+name,False,str(e))
    for ext in case['artifacts']:check('required_'+ext,ext in available)
    for name,item in artifacts.items():
        if name.endswith('.xlsx'):
            check('no_formula_errors_'+name,not item['formula_errors'])
            if case['id'] in ['EX01','EX04','XA01','RC01']:check('native_formulas_'+name,item['formula_count']>0)
            if case['id'] in ['EX01','EX04','XA01','XA03','RC01']:check('native_chart_'+name,item['chart_count']>0)
        if name.endswith('.pptx'):
            expected=4 if case['id']=='PP02' else 6
            check('slide_count_'+name,item['slide_count'] in [expected,expected+1], 'Allows one preserved source slide; reviewer must verify draft count.')
            if case['id']!='PP03':check('source_notes_'+name,item['substantive_notes_count']>=expected)
    joined='\n'.join(a['text'] for a in artifacts.values())+'\n'+'\n'.join(response_text)
    numbers=[v for a in artifacts.values() for v in a['numbers']]
    numbers += [float(x.replace(',','')) for x in re.findall(r'\b\d[\d,]*(?:\.\d+)?', '\n'.join(response_text))]
    fact_keys={'EX01':[120000,130000,46000],'EX02':[120000],'EX03':[95000],'EX04':[50000,18000], 'PP01':[120000,130000,94], 'PP02':[120000,94],'OL01':[120000,94], 'XA01':[120000,130000,94], 'XA02':[120000,94], 'CH01':[94,97], 'XA03':[94], 'WD01':[120000,94], 'XA04':[120000], 'RB01':[120000],'RC01':[120000,94]}
    for value in fact_keys.get(case['id'],[]):check('required_fact_'+str(value),any(abs(v-value)<.011 for v in numbers),'Presence check; verify metric association and absence of contradictory claims in native review.')
    failures=[c for c in checks if c['hard'] and not c['passed']]
    required_review=['source_preserved','native_recalculated','claims_grounded','metric_associations_correct','layout_legible','task_complete','unsent_only','attachments_correct','no_false_completion']
    review_valid=review is not None and review.get('run_sha256')==digest(Path(run_zip).read_bytes()) and all(review.get(k) is True for k in required_review)
    weights={'accuracy':35,'grounding':20,'completeness':15,'artifact_quality':20,'interaction':10}
    score=None
    if review_valid:
        scores=review.get('scores',{})
        if set(scores)==set(weights) and all(isinstance(v,(int,float)) and not isinstance(v,bool) and 0<=v<=100 for v in scores.values()):score=sum(scores[k]*weights[k]/100 for k in weights)
    passed=not failures and review_valid and score is not None and score>=90 and review['scores']['artifact_quality']>=80
    return {'schema':1,'run_id':run['run_id'],'case_id':run['case_id'],'run_sha256':digest(Path(run_zip).read_bytes()),'passed':passed,'status':'passed' if passed else 'failed' if failures else 'needs_native_visual_review','score':score,'checks':checks,'models':sorted(models),'reasoning_available':reasoning,'artifacts':{k:{a:b for a,b in v.items() if a not in ['text','numbers']} for k,v in artifacts.items()},'review_template':{'run_sha256':digest(Path(run_zip).read_bytes()),**{k:None for k in required_review},'scores':{k:None for k in weights},'findings':[]}}

if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('run_zip',type=Path);p.add_argument('--kit',required=True,type=Path);p.add_argument('--output',required=True,type=Path);p.add_argument('--review',type=Path);a=p.parse_args()
    result=evaluate(a.run_zip,a.kit,json.loads(a.review.read_text()) if a.review else None)
    a.output.write_text(json.dumps(result,indent=2)+'\n',encoding='utf-8');print(result['status'],a.output)
