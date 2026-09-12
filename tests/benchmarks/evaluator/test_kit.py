from pathlib import Path
import csv, hashlib, io, json, unittest, zipfile, email, email.policy, tempfile
from evaluate import inspect_office, read_archive, evaluate

BASE=Path(__file__).resolve().parents[1]
class KitTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.files=read_archive(BASE/'releases/scribble-test-kit-v1.zip')
        cls.prefix='scribble-test-kit-v1/'
    def data(self,p):return self.files[self.prefix+p]
    def test_hashes_and_roles(self):
        m=json.loads(self.data('manifest.json'));self.assertEqual(len(m['files']),len(self.files)-1)
        for f in m['files']:
            b=self.data(f['path']);self.assertEqual(f['sha256'],hashlib.sha256(b).hexdigest());self.assertEqual(f['size'],len(b))
    def test_independent_arithmetic(self):
        rows=list(csv.DictReader(io.StringIO(self.data('inputs/data/sales.csv').decode())))
        june=[r for r in rows if r['Period']=='2026-06'];may=[r for r in rows if r['Period']=='2026-05']
        rev=lambda rs:sum(int(r['RevenueEUR']) for r in rs)
        cost=lambda rs:sum(int(r['CostEUR']) for r in rs)
        self.assertEqual(rev(june),120000);self.assertEqual(cost(june),74000);self.assertEqual(rev(may),100000)
        self.assertAlmostEqual((rev(june)-cost(june))/rev(june),.38333333333333336)
        self.assertEqual(len({r['RowID'] for r in rows}),8)
    def test_case_inputs_exist(self):
        cases=json.loads(self.data('operator/cases.json'));self.assertEqual(len(cases),16)
        self.assertEqual({c['host'] for c in cases},{'Excel','PowerPoint','Word','Outlook','Chrome'})
        for c in cases:
            for p in c['inputs']:self.assertIn(self.prefix+p,self.files);self.assertTrue(p.startswith('inputs/'))
        missing=next(c for c in cases if c['id']=='EX03');self.assertEqual(missing['inputs'],['inputs/excel/Atlas-missing.xlsx'])
    def test_email_attachment_bytes(self):
        source={Path(p).name:b for p,b in self.files.items() if '/inputs/' in p and not p.endswith('.eml')}
        for p,b in self.files.items():
            if '/inputs/outlook/' not in p:continue
            msg=email.message_from_bytes(b,policy=email.policy.default)
            self.assertIn('example.test',msg['From'])
            for part in msg.iter_attachments():self.assertEqual(part.get_payload(decode=True),source[part.get_filename()])
    def test_native_references(self):
        book=inspect_office(self.data('evaluator-only/expected-analysis.xlsx'),'xlsx')
        self.assertGreater(book['formula_count'],10);self.assertEqual(book['formula_errors'],[]);self.assertGreater(book['chart_count'],0)
        deck=inspect_office(self.data('evaluator-only/reference-deck.pptx'),'pptx');self.assertEqual(deck['slide_count'],6);self.assertGreaterEqual(deck['chart_count'],2);self.assertEqual(deck['notes_count'],6)
    def test_no_macro_or_real_domains(self):
        for p,b in self.files.items():
            if p.endswith(('.xlsx','.pptx','.docx')):
                with zipfile.ZipFile(io.BytesIO(b)) as z:self.assertFalse(any('vbaProject' in n for n in z.namelist()))
            if '/inputs/browser/' in p:self.assertNotIn(b'https://',b)
    def test_claimed_success_without_outputs_does_not_pass(self):
        with tempfile.TemporaryDirectory() as tmp:
            root=Path(tmp);kit=root/'kit';kit.mkdir();(kit/'operator').mkdir()
            (kit/'manifest.json').write_bytes(self.data('manifest.json'));(kit/'operator/cases.json').write_bytes(self.data('operator/cases.json'))
            run={'run_id':'a'*32,'case_id':'EX01','status':'finished','trace_complete':True,'manifest_sha256':hashlib.sha256(self.data('manifest.json')).hexdigest()}
            event={'schema':1,'run_id':'a'*32,'task_id':'t','instance_id':'i','sequence':1,'utc':'2026-07-03T12:00:00Z','stage':'case_finished','detail':{'content':'I created the workbook successfully.'}}
            payload={'run.json':json.dumps(run).encode(),'timeline.jsonl':json.dumps(event).encode()}
            manifest={'files':[{'path':p,'size':len(b),'sha256':hashlib.sha256(b).hexdigest()} for p,b in payload.items()]}
            payload['export-manifest.json']=json.dumps(manifest).encode()
            archive=root/'run.zip'
            with zipfile.ZipFile(archive,'w') as z:
                for p,b in payload.items():z.writestr(p,b)
            report=evaluate(archive,kit)
            self.assertFalse(report['passed']);self.assertTrue(any(c['name']=='required_final_xlsx' and not c['passed'] for c in report['checks']))

if __name__=='__main__':unittest.main()
