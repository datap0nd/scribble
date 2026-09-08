"""Compare evaluator reports without discarding failed attempts."""
import argparse,json,statistics
from pathlib import Path

def compare(paths):
    rows=[json.loads(p.read_text(encoding='utf-8-sig')) for p in paths]
    groups={}
    for r in rows:
        key=r['case_id']+' / '+','.join(r.get('models',[]))
        g=groups.setdefault(key,{'attempts':0,'passes':0,'scores':[],'runs':[],'failed_checks':{}})
        g['attempts']+=1;g['passes']+=bool(r.get('passed'));g['runs'].append(r['run_id'])
        if r.get('score') is not None:g['scores'].append(r['score'])
        for c in r.get('checks',[]):
            if not c['passed']:g['failed_checks'][c['name']]=g['failed_checks'].get(c['name'],0)+1
    for g in groups.values():
        g['pass_rate']=g['passes']/g['attempts'];g['median_score']=statistics.median(g['scores']) if g['scores'] else None
        if len(set(g['runs']))!=len(g['runs']):raise ValueError('Duplicate run included in comparison')
    return {'schema':1,'groups':groups,'note':'Compare only matching fixture, build and model settings. Small samples do not establish reliability.'}

if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('reports',nargs='+',type=Path);p.add_argument('--output',type=Path,required=True);a=p.parse_args()
    a.output.write_text(json.dumps(compare(a.reports),indent=2)+'\n',encoding='utf-8')
