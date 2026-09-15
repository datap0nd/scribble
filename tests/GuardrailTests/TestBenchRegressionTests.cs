using System;
using System.IO;
using System.IO.Compression;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Web.Script.Serialization;
using Scribble.Office;
using Scribble.Outlook;
using Scribble.Testing;

namespace GuardrailTests
{
    internal static class TestBenchRegressionTests
    {
        private static void Check(bool value, string message) { if (!value) throw new InvalidOperationException(message); }

        public static void OperationsPdf()
        {
            var directory = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory);
            while (directory != null && !File.Exists(Path.Combine(directory.FullName, "tests", "benchmarks", "releases", "scribble-test-kit-v1.zip"))) directory = directory.Parent;
            Check(directory != null, "The checked-in benchmark kit was not found.");
            var path = Path.Combine(Path.GetTempPath(), "scribble-pdf-regression-" + Guid.NewGuid().ToString("N") + ".pdf");
            try
            {
                using (var archive = ZipFile.OpenRead(Path.Combine(directory.FullName, "tests", "benchmarks", "releases", "scribble-test-kit-v1.zip")))
                {
                    var entry = archive.Entries.Single(e => e.FullName.Replace('\\', '/').EndsWith("inputs/pdf/Atlas-operations-note.pdf", StringComparison.Ordinal));
                    entry.ExtractToFile(path);
                }
                var fileText = PdfTextExtractor.Extract(path, 200000, CancellationToken.None);
                var byteText = PdfTextExtractor.Extract(File.ReadAllBytes(path), 200000);
                foreach (var text in new[] { fileText, byteText })
                    foreach (var fact in new[] { "Mira Cole", "Leon Park", "94", "97", "10 July 2026", "12 July 2026" })
                        Check(text.Contains(fact), "Ordered ASCII85/Flate extraction lost " + fact + ".");
                Check(fileText == byteText, "File and attachment PDF readers disagree.");
                using (var cancelled = new CancellationTokenSource())
                {
                    cancelled.Cancel();
                    var stopped = false;
                    try { PdfTextExtractor.Extract(path, 200000, cancelled.Token); }
                    catch (OperationCanceledException) { stopped = true; }
                    Check(stopped, "PDF extraction ignored Stop.");
                }
            }
            finally { if (File.Exists(path)) File.Delete(path); }
        }

        public static void ExcelErrors()
        {
            Check(ExcelErrorValue.Text(-2146826273) == "#VALUE!", "The work-PC VT_ERROR was treated as a number.");
            Check(ExcelErrorValue.Text(-2146826281) == "#DIV/0!", "The margin division error was missed.");
            Check(ExcelErrorValue.Text(new ErrorWrapper(-2146826273)) == "#VALUE!", "Wrapped HRESULT lost its error identity.");
            Check(ExcelErrorValue.Text(new ErrorWrapper(2007)) == "#DIV/0!", "Wrapped CVErr was missed.");
            Check(ExcelErrorValue.Text(2023) == "#REF!", "CVErr was missed.");
            foreach (var value in new object[] { 2007d, -2146826273d, 46000, "2007", null })
                Check(ExcelErrorValue.Text(value) == null, "An ordinary cell value was called an Excel error.");
            var normalized = TestLabEvaluator.NormalizeExcelErrors("R8C4: -2146826273 | formula: =Sales!E2+E3\nR9C4: -2146826281 | formula: =0/0\nR10C1: -2146826273");
            Check(normalized.Contains("R8C4: #VALUE!") && normalized.Contains("R9C4: #DIV/0!") && normalized.Contains("R10C1: -2146826273"), "Legacy native formula evidence was not normalized precisely.");
        }

        public static void ReportFacts()
        {
            const string correct = "To: review@example.test\nUnsent: true\nRevenue: EUR 120000; Budget gap: -10000 (-7.69%); Gross profit: EUR 46000; Cost: EUR 74000; Weighted gross margin: 38.33%; margin change: -1.67 pp; Delivery: 94% against target 97%.\nPlanned actions: Mira Cole reviews freight costs by 10 July 2026. Leon Park confirms recovery by 12 July 2026.";
            var valid = TestLabEvaluator.CheckOutputFacts("OL02", correct);
            Check(valid.All(c => c.passed), "The correct control failed: " + string.Join("; ", valid.Where(c => !c.passed).Select(c => c.name)));
            Check(TestLabEvaluator.CheckOutputFacts("OL02", correct + "\nJune revenue 120000, May revenue 100000; North revenue 70000; May gross margin 40%.").All(c => c.passed), "Historical or regional metrics were mistaken for June totals.");
            Check(TestLabEvaluator.CheckOutputFacts("OL02", correct + "\nGross profit was EUR 40000 in May; June gross profit was EUR 46000.").All(c => c.passed), "A period stated after a historical value was mistaken for June.");
            Check(TestLabEvaluator.CheckOutputFacts("OL02", correct.Replace("Mira Cole reviews freight costs by 10 July 2026. Leon Park confirms recovery by 12 July 2026.",
                "10 July 2026: Mira Cole reviews freight costs. 12 July 2026: Leon Park confirms recovery.")).All(c => c.passed), "Dates preceding their owners were rejected.");
            Check(TestLabEvaluator.CheckOutputFacts("OL02", correct.Replace("7.69%", "8%") + "\nReference 7.69%.").Any(c => c.name == "consistent_budget_gap_percentage" && !c.passed), "An incorrect budget percentage escaped its metric association.");
            var wrong = correct.Replace("46000", "52000").Replace("38.33%", "43.33%").Replace("-1.67 pp", "+3.33 pp");
            var bad = TestLabEvaluator.CheckOutputFacts("OL02", wrong + "\nUnrelated reference numbers: 46000, 38.33, -1.67.");
            foreach (var check in new[] { "consistent_june_profit", "consistent_june_margin", "consistent_june_margin_change" })
                Check(bad.Any(c => c.name == check && !c.passed), "Wrong numeric association escaped: " + check);
            Check(TestLabEvaluator.CheckOutputFacts("OL02", correct.Replace("Gross profit: EUR 46000", "Gross profit (EUR): 52000")).Any(c => c.name == "consistent_june_profit" && !c.passed), "A wrong profit escaped through a unit-bearing label.");
            Check(TestLabEvaluator.CheckOutputFacts("OL02", correct.Replace("Weighted gross margin: 38.33%", "Weighted gross margin (%): 43.33") + "\nReference: 38.33%.").Any(c => c.name == "consistent_june_margin" && !c.passed), "A wrong margin escaped through a unit-bearing label.");
            foreach (var pair in new[] { new[] { "10 July 2026", "12 July 2026" }, new[] { "review@example.test", "wrong@example.test" }, new[] { "Unsent: true", "Unsent: false" } })
                Check(TestLabEvaluator.CheckOutputFacts("OL02", correct.Replace(pair[0], pair[1])).Any(c => !c.passed), "The draft's date, recipient or unsent state escaped validation.");
            var formulaOnly = TestLabEvaluator.CheckOutputFacts("EX02", "R4C2: 0 | formula: =120000*0");
            Check(formulaOnly.Any(c => !c.passed), "Formula operands were accepted as calculated answers.");
            Check(TestLabEvaluator.CheckOutputFacts("EX05", "120000 74000 46000 0.383333333333").All(c => c.passed), "A native percentage fraction was rejected.");
            Check(TestLabEvaluator.CheckOutputFacts("EX02", "Worksheet: Scribble Draft\nR1C1: 0\nSeries 1: =SERIES(120000,Sheet1!A1:A2,Sheet1!B1:B2,1)").Any(c => !c.passed), "A chart's formula operands were accepted as answers.");
            Check(TestLabEvaluator.CheckOutputFacts("PP01", "Slide 1\n[Scribble draft]\nShape 1 | type 1 | x,y,w,h: 120000,94,97,100").Any(c => c.name == "required_final_fact_120000" && !c.passed), "Shape geometry was accepted as a financial fact.");
        }

        public static void NativeOutputBoundaries()
        {
            const string source = "Worksheet: Sales | used range: A1:B2\nR1C1: 50000 | formula: =SUM(A2:A3)\nR1C2: #VALUE! | formula: =A1+1\nNative charts: 1\n";
            const string emptyDraft = "Worksheet: Scribble Draft | used range: A1:B1\nR1C1: 50000\nR1C2: 18000\nNative charts: 0\n";
            var missing = TestLabEvaluator.CheckOutputStructure("EX04", "xlsx", source + emptyDraft);
            foreach (var name in new[] { "native_formulas", "native_chart" })
                Check(missing.Any(c => c.name == name && !c.passed), "Source structure satisfied a new draft requirement: " + name);
            Check(missing.Single(c => c.name == "no_formula_errors").passed, "An unchanged source error failed the new draft.");
            const string goodDraft = "Worksheet: Scribble Draft | used range: A1:B1\nR1C1: 50000 | formula: =SUM(Sales!E6:E9)\nR1C2: 18000\nNative charts: 1\n";
            Check(TestLabEvaluator.CheckOutputStructure("EX04", "xlsx", source + goodDraft).All(c => c.passed), "Valid draft structure was confused with preserved source contents.");
            Check(TestLabEvaluator.CheckOutputStructure("EX04", "xlsx", goodDraft.Replace("R1C2: 18000", "R1C2: -2146826273 | formula: =Sales!E2+E3"), true).Any(c => c.name == "no_formula_errors" && !c.passed), "A legacy signed error escaped the output boundary.");
            var sources = string.Join("\n", Enumerable.Range(1, 6).Select(n => "Slide " + n + "\nPreserved source"));
            var oneDraft = "Slide count: 7\n" + sources + "\nSlide 7\n[Scribble draft] Incomplete";
            Check(TestLabEvaluator.CheckOutputStructure("PP03", "pptx", oneDraft).Any(c => !c.passed), "Six source slides plus one new slide satisfied a six-slide task.");
            var fiveDrafts = "Slide count: 6\nSlide 1\nPreserved source\n" + string.Join("\n", Enumerable.Range(2, 5).Select(n => "Slide " + n + "\n[Scribble draft] New"));
            Check(TestLabEvaluator.CheckOutputStructure("PP01", "pptx", fiveDrafts).Any(c => !c.passed), "A preserved starter slide substituted for a missing draft slide.");
            var sixDrafts = string.Join("\n", Enumerable.Range(7, 6).Select(n => "Slide " + n + "\n[Scribble draft] New"));
            Check(TestLabEvaluator.CheckOutputStructure("PP03", "pptx", "Slide count: 12\n" + sources + "\n" + sixDrafts).All(c => c.passed), "Preserving six source slides caused a correct six-slide draft to fail.");
            Check(TestLabEvaluator.CheckOutputStructure("PP04", "pptx", "Slide count: 6\n" + sixDrafts, true).Any(c => c.name == "native_chart" && !c.passed), "A financial chart review passed without an editable chart.");
            Check(TestLabEvaluator.CheckOutputStructure("PP04", "pptx", "Slide count: 6\n" + sixDrafts + "\nNative chart type: 51", true).All(c => c.passed), "A boundary-aware six-slide draft with a native chart failed.");
        }

        public static void NativeMailHeaders()
        {
            var serializer = new JavaScriptSerializer();
            const string body = "To: review@example.test\nUnsent: true\nQuoted draft request.";
            var good = serializer.Serialize(new { to = "review@example.test", cc = "", bcc = "", unsent = true, text = body });
            Check(TestLabEvaluator.CheckMailReadback(good).All(c => c.passed), "Correct structured native mail metadata failed.");
            foreach (var metadata in new[] {
                new { to = "wrong@example.test", cc = "", bcc = "", unsent = true, text = body },
                new { to = "review@example.test", cc = "wrong@example.test", bcc = "", unsent = true, text = body },
                new { to = "review@example.test", cc = "", bcc = "wrong@example.test", unsent = true, text = body },
                new { to = "review@example.test", cc = "", bcc = "", unsent = false, text = body } })
                Check(TestLabEvaluator.CheckMailReadback(serializer.Serialize(metadata)).Any(c => !c.passed), "Body text overrode wrong native recipients or sent state.");
            var legacy = "To: wrong@example.test\nCC: \nSubject: Test\nUnsent: true\n" + body;
            Check(TestLabEvaluator.CheckMailReadback(serializer.Serialize(new { text = legacy })).Any(c => c.name == "draft_recipient" && !c.passed), "A quoted To header overrode the first native header.");
            Check(TestLabEvaluator.CheckMailReadback(serializer.Serialize(new { text = "To: review@example.test\nBCC: hidden@example.test\nUnsent: true\n" + body })).Any(c => !c.passed), "A legacy BCC recipient escaped validation.");
        }

        public static void NativePackageBoundaries()
        {
            const string goodMemory = "Worksheet: Scribble Draft\nR1C1: 50000 | formula: =SUM(Sales!A1:A2)\nR1C2: 18000\nNative charts: 1\n";
            var good = EvaluatePackages("EX04", new[] { WorkbookPackage(true, false) }, goodMemory);
            Check(good.checks.All(c => c.passed), "Source errors in the full disk package failed an authoritative output-only readback: " + string.Join(",", good.findings));

            var badStructure = EvaluatePackages("EX04", new[] { WorkbookPackage(false, false) }, null);
            foreach (var name in new[] { "native_formulas_", "native_chart_" })
                Check(badStructure.checks.Any(c => c.name.StartsWith(name, StringComparison.Ordinal) && !c.passed), "File-only source structure satisfied the draft requirement: " + name);
            Check(badStructure.checks.Where(c => c.name.StartsWith("no_formula_errors_", StringComparison.Ordinal)).All(c => c.passed), "A preserved source error leaked into file-only output grading.");

            var wrong = EvaluatePackages("EX02", new[] { WorkbookPackage(false, true) }, null);
            Check(wrong.checks.Any(c => c.name == "required_final_fact_120000" && !c.passed), "Source shared strings or formula operands satisfied the file-only numeric oracle.");

            var two = EvaluatePackages("EX04", new[] { WorkbookPackage(true, false), WorkbookPackage(false, false) }, goodMemory);
            Check(two.checks.Any(c => c.name.StartsWith("native_chart_", StringComparison.Ordinal) && c.name.Contains("-output-2.xlsx") && !c.passed),
                "A readback for workbook one suppressed grading workbook two.");

            var ambiguous = EvaluatePackages("EX04", new[] { WorkbookPackage(true, false) }, null, "Worksheet: Scribble Draft\nR1C1: 50000\nNative charts: 1");
            Check(ambiguous.checks.Any(c => c.name.StartsWith("output_boundary_", StringComparison.Ordinal) && !c.passed), "A pre-existing draft label was accepted as proof of new output ownership.");
            var legacy = EvaluatePackages("EX04", new[] { WorkbookPackage(true, false) }, goodMemory, "Worksheet: Scribble Draft\nR1C1: 50000\nNative charts: 1", "missing");
            Check(legacy.checks.Any(c => c.name.StartsWith("output_boundary_", StringComparison.Ordinal) && !c.passed), "A legacy memory readback bypassed an ambiguous source boundary.");
            foreach (var mode in new[] { "false", "corrupt" })
            {
                var invalid = EvaluatePackages("EX04", new[] { WorkbookPackage(true, false) }, goodMemory, boundaryMode: mode);
                Check(invalid.checks.Any(c => c.name.StartsWith("output_boundary_", StringComparison.Ordinal) && !c.passed), "A " + mode + " boundary flag did not produce an explicit failed check.");
            }
            Check(EvaluatePackages("EX04", new[] { WorkbookPackage(true, false) }, goodMemory, boundaryMode: "missing").checks.All(c => c.passed), "An unambiguous legacy source boundary was rejected.");
            Check(EvaluatePackages("EX04", new[] { WorkbookPackage(true, false) }, goodMemory.Replace("Scribble Draft", "Margin Audit")).checks.All(c => c.passed), "An authentic run-owned renamed output sheet was ignored.");
        }

        private static byte[] WorkbookPackage(bool goodDraft, bool formulaOnly)
        {
            const string sheets = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            const string rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            const string rels = "http://schemas.openxmlformats.org/package/2006/relationships";
            var parts = new Dictionary<string, string> {
                { "xl/workbook.xml", "<workbook xmlns='" + sheets + "' xmlns:r='" + rel + "'><sheets><sheet name='Sales' sheetId='1' r:id='rId1'/><sheet name='Scribble Draft' sheetId='2' r:id='rId2'/></sheets></workbook>" },
                { "xl/_rels/workbook.xml.rels", "<Relationships xmlns='" + rels + "'><Relationship Id='rId1' Target='worksheets/sheet1.xml'/><Relationship Id='rId2' Target='worksheets/sheet2.xml'/></Relationships>" },
                { "xl/sharedStrings.xml", "<sst xmlns='" + sheets + "'><si><t>120000</t></si></sst>" },
                { "xl/worksheets/sheet1.xml", "<worksheet xmlns='" + sheets + "' xmlns:r='" + rel + "'><sheetData><row r='1'><c r='A1' t='s'><v>0</v></c><c r='B1'><f>50000+18000</f><v>68000</v></c><c r='C1' t='e'><f>1+&quot;text&quot;</f><v>#VALUE!</v></c></row></sheetData><drawing r:id='drawing'/></worksheet>" },
                { "xl/worksheets/_rels/sheet1.xml.rels", "<Relationships xmlns='" + rels + "'><Relationship Id='drawing' Target='../drawings/drawing1.xml'/></Relationships>" },
                { "xl/drawings/drawing1.xml", "<drawing xmlns:r='" + rel + "'><chart r:id='chart'/></drawing>" },
                { "xl/drawings/_rels/drawing1.xml.rels", "<Relationships xmlns='" + rels + "'><Relationship Id='chart' Target='../charts/chart1.xml'/></Relationships>" },
                { "xl/charts/chart1.xml", "<chart/>" }
            };
            var value = formulaOnly ? "<f>120000*0</f><v>0</v>" : goodDraft ? "<f>SUM(Sales!A1:A2)</f><v>50000</v>" : "<v>50000</v>";
            parts["xl/worksheets/sheet2.xml"] = "<worksheet xmlns='" + sheets + "' xmlns:r='" + rel + "'><sheetData><row r='1'><c r='A1'>" + value + "</c><c r='B1'><v>18000</v></c></row></sheetData>" + (goodDraft ? "<drawing r:id='drawing'/>" : "") + "</worksheet>";
            if (goodDraft) parts["xl/worksheets/_rels/sheet2.xml.rels"] = "<Relationships xmlns='" + rels + "'><Relationship Id='drawing' Target='../drawings/drawing1.xml'/></Relationships>";
            using (var stream = new MemoryStream())
            {
                using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, true))
                    foreach (var part in parts) WriteEntry(archive, part.Key, part.Value);
                return stream.ToArray();
            }
        }

        private static TestLabEvaluation EvaluatePackages(string caseId, byte[][] packages, string memory, string source = "Worksheet: Sales\nR1C1: 120000\nNative charts: 1", string boundaryMode = "valid")
        {
            var root = Path.Combine(Path.GetTempPath(), "scribble-package-regression-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(Path.Combine(root, "operator"));
            var serializer = new JavaScriptSerializer();
            try
            {
                File.WriteAllText(Path.Combine(root, "operator", "cases.json"), serializer.Serialize(new[] { new { id = caseId, artifacts = new[] { "xlsx" } } }));
                var path = Path.Combine(root, "evidence.zip");
                using (var archive = ZipFile.Open(path, ZipArchiveMode.Create))
                {
                    WriteEntry(archive, "run.json", serializer.Serialize(new { run_id = "package-regression", case_id = caseId, fixture_root = root, status = "finished", trace_complete = true, assisted = false }));
                    WriteEntry(archive, "artifacts/before-Excel-source-source-1-readback.json", serializer.Serialize(new { text = source }));
                    WriteEntry(archive, "artifacts/after-Excel-final-source-1-readback.json", serializer.Serialize(new { text = source }));
                    for (var i = 0; i < packages.Length; i++)
                        using (var output = archive.CreateEntry("artifacts/disk-Excel-final-output-" + (i + 1) + ".xlsx").Open()) output.Write(packages[i], 0, packages[i].Length);
                    if (memory != null)
                    {
                        var readback = new Dictionary<string, object> { { "native_readback", true }, { "run_created_output", true }, { "artifact_extension", "xlsx" }, { "text", memory } };
                        if (boundaryMode != "missing") readback["output_boundary"] = boundaryMode == "corrupt" ? (object)"not a boolean" : boundaryMode != "false";
                        WriteEntry(archive, "artifacts/memory-Excel-final-output-1-readback.json", serializer.Serialize(readback));
                    }
                }
                return TestLabEvaluator.Evaluate(path, Path.Combine(root, "evaluation"));
            }
            finally { Directory.Delete(root, true); }
        }

        private static void WriteEntry(ZipArchive archive, string name, string contents)
        { using (var writer = new StreamWriter(archive.CreateEntry(name).Open())) writer.Write(contents); }
    }
}
