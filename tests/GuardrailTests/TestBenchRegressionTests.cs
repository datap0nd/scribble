using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
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
            Check(TestLabEvaluator.CheckOutputFacts("OL02", correct.Replace("7.69%", "8%") + "\nReference 7.69%.").Any(c => c.name == "consistent_budget_gap_percentage" && !c.passed), "An incorrect budget percentage escaped its metric association.");
            var wrong = correct.Replace("46000", "52000").Replace("38.33%", "43.33%").Replace("-1.67 pp", "+3.33 pp");
            var bad = TestLabEvaluator.CheckOutputFacts("OL02", wrong + "\nUnrelated reference numbers: 46000, 38.33, -1.67.");
            foreach (var check in new[] { "consistent_june_profit", "consistent_june_margin", "consistent_june_margin_change" })
                Check(bad.Any(c => c.name == check && !c.passed), "Wrong numeric association escaped: " + check);
            foreach (var pair in new[] { new[] { "10 July 2026", "12 July 2026" }, new[] { "review@example.test", "wrong@example.test" }, new[] { "Unsent: true", "Unsent: false" } })
                Check(TestLabEvaluator.CheckOutputFacts("OL02", correct.Replace(pair[0], pair[1])).Any(c => !c.passed), "The draft's date, recipient or unsent state escaped validation.");
            var formulaOnly = TestLabEvaluator.CheckOutputFacts("EX02", "R4C2: 0 | formula: =120000*0");
            Check(formulaOnly.Any(c => !c.passed), "Formula operands were accepted as calculated answers.");
            Check(TestLabEvaluator.CheckOutputFacts("EX05", "120000 74000 46000 0.383333333333").All(c => c.passed), "A native percentage fraction was rejected.");
        }
    }
}
