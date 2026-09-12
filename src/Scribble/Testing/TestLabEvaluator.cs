using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;
using System.Xml.Linq;

namespace Scribble.Testing
{
    // Deterministic post-capture checks. This boundary reads sealed evidence;
    // model responses cannot add rules, change scores, or execute content.
    public static class TestLabEvaluator
    {
        public static TestLabEvaluation Evaluate(string evidence, string destination)
        {
            var checks = new List<TestLabCheck>(); var json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
            LabRun run; LabCase testCase; var finalNumbers = new List<double>(); var finalText = new StringBuilder();
            using (var archive = ZipFile.OpenRead(evidence))
            {
                run = json.Deserialize<LabRun>(Read(archive, "run.json"));
                testCase = TestLabSuite.Read<LabCase[]>(TestLab.SafeChild(run.fixture_root, "operator/cases.json")).Single(c => c.id == run.case_id);
                Add(checks, "trace_complete", run.trace_complete && run.status == "finished", "The capture must reach a terminal case event.");
                Add(checks, "unassisted", !run.assisted, "Manual intervention is retained but cannot count as an unassisted pass.");
                foreach (var extension in testCase.artifacts ?? new string[0])
                {
                    var suffix = "." + extension;
                    Add(checks, "required_final_" + extension, archive.Entries.Any(e => e.FullName.StartsWith("artifacts/", StringComparison.Ordinal) &&
                        e.FullName.IndexOf("-final-output-", StringComparison.OrdinalIgnoreCase) >= 0 && e.FullName.EndsWith(suffix, StringComparison.OrdinalIgnoreCase)),
                        "Only a final run-owned output satisfies this requirement.");
                }
                foreach (var entry in archive.Entries.Where(e => e.FullName.StartsWith("artifacts/", StringComparison.Ordinal) &&
                    e.FullName.IndexOf("-final-output-", StringComparison.OrdinalIgnoreCase) >= 0 &&
                    new[] { ".xlsx", ".pptx", ".docx" }.Contains(Path.GetExtension(e.FullName).ToLowerInvariant())))
                {
                    try
                    {
                        var inspection = InspectOffice(entry); finalNumbers.AddRange(inspection.numbers); finalText.AppendLine(inspection.text);
                        if (entry.FullName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                        {
                            Add(checks, "no_formula_errors_" + entry.Name, inspection.formulaErrors == 0, "Native formula error cells: " + inspection.formulaErrors);
                            if (new[] { "EX01", "EX04", "XA01", "RC01" }.Contains(run.case_id))
                                Add(checks, "native_formulas_" + entry.Name, inspection.formulas > 0, "Formula count: " + inspection.formulas);
                            if (new[] { "EX01", "EX04", "XA01", "XA03", "RC01" }.Contains(run.case_id))
                                Add(checks, "native_chart_" + entry.Name, inspection.charts > 0, "Chart count: " + inspection.charts);
                        }
                        if (entry.FullName.EndsWith(".pptx", StringComparison.OrdinalIgnoreCase))
                        {
                            var expected = run.case_id == "PP02" ? 4 : 6;
                            Add(checks, "final_slide_count_" + entry.Name, inspection.slides == expected || inspection.slides == expected + 1,
                                "Expected " + expected + " draft slides (plus at most one preserved source); observed " + inspection.slides + ".");
                        }
                    }
                    catch (Exception error) { Add(checks, "readable_" + entry.Name, false, error.GetType().Name + ": " + error.Message); }
                }
                if ((testCase.artifacts ?? new string[0]).Length == 0)
                {
                    var timeline = Read(archive, "timeline.jsonl"); finalText.AppendLine(timeline);
                    finalNumbers.AddRange(Numbers(timeline));
                }
                CheckFacts(checks, run.case_id, finalNumbers);
                CheckSourcePreservation(checks, archive);
            }
            var failed = checks.Where(c => c.hard && !c.passed).ToArray();
            var evaluation = new TestLabEvaluation { schema = 1, run_id = run.run_id, case_id = run.case_id,
                evidence_sha256 = TestLab.FileHash(evidence), status = failed.Length > 0 ? "failed" : "needs_native_visual_review",
                passed = false, checks = checks.ToArray(), findings = failed.Select(c => c.name + ": " + c.detail).ToArray(),
                note = failed.Length > 0 ? "Deterministic hard gates failed. Native visual review remains required." : "Deterministic gates passed. Native visual review and scoring remain required." };
            Directory.CreateDirectory(destination);
            File.WriteAllText(Path.Combine(destination, "evaluation.json"), TestLab.Serialize(evaluation), new UTF8Encoding(false));
            return evaluation;
        }

        private static void CheckSourcePreservation(List<TestLabCheck> checks, ZipArchive archive)
        {
            var before = Readbacks(archive, "-source-source-"); var after = Readbacks(archive, "-final-source-");
            foreach (var pair in before)
            {
                string value;
                Add(checks, "source_preserved_" + pair.Key, after.TryGetValue(pair.Key, out value) && value == pair.Value,
                    after.ContainsKey(pair.Key) ? "Native source readback changed between pre-write and final capture." : "Final source readback is missing.");
            }
        }

        private static Dictionary<string, string> Readbacks(ZipArchive archive, string marker)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase); var json = new JavaScriptSerializer();
            foreach (var entry in archive.Entries.Where(e => e.FullName.IndexOf(marker, StringComparison.OrdinalIgnoreCase) >= 0 && e.FullName.EndsWith("-readback.json", StringComparison.OrdinalIgnoreCase)))
            {
                var data = json.Deserialize<Dictionary<string, object>>(Read(archive, entry.FullName)); object text;
                var logical = entry.Name.Substring(entry.Name.IndexOf('-') + 1);
                var markerAt = logical.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
                var suffix = logical.Substring(markerAt + marker.Length).Replace("-readback.json", "");
                if (data.TryGetValue("text", out text)) result[logical.Substring(0, markerAt) + "-" + suffix] = Convert.ToString(text);
            }
            return result;
        }

        private static void CheckFacts(List<TestLabCheck> checks, string caseId, List<double> numbers)
        {
            var facts = new Dictionary<string, double[]> {
                { "EX01", new[] { 120000d, 130000d, 46000d } }, { "EX02", new[] { 120000d } },
                { "EX03", new[] { 95000d } }, { "EX04", new[] { 50000d, 18000d } },
                { "PP01", new[] { 120000d, 130000d, 94d } }, { "PP02", new[] { 120000d, 94d } },
                { "OL01", new[] { 120000d, 94d } }, { "XA01", new[] { 120000d, 130000d, 94d } },
                { "XA02", new[] { 120000d, 94d } }, { "CH01", new[] { 94d, 97d } },
                { "XA03", new[] { 94d } }, { "WD01", new[] { 120000d, 94d } },
                { "XA04", new[] { 120000d } }, { "RB01", new[] { 120000d } }, { "RC01", new[] { 120000d, 94d } } };
            double[] expected;
            if (!facts.TryGetValue(caseId, out expected)) return;
            foreach (var value in expected) Add(checks, "required_final_fact_" + value.ToString(CultureInfo.InvariantCulture),
                numbers.Any(n => Math.Abs(n - value) < .011), "Expected value was not found in the final run-owned evidence boundary.");
        }

        private static OfficeInspection InspectOffice(ZipArchiveEntry entry)
        {
            using (var bytes = new MemoryStream())
            {
                using (var input = entry.Open()) input.CopyTo(bytes); bytes.Position = 0;
                using (var package = new ZipArchive(bytes, ZipArchiveMode.Read))
                {
                    if (package.Entries.Sum(e => e.Length) > 100L * 1024 * 1024) throw new InvalidDataException("Expanded package exceeds 100 MB.");
                    if (package.Entries.Any(e => e.FullName.IndexOf("vbaProject", StringComparison.OrdinalIgnoreCase) >= 0)) throw new InvalidDataException("Unexpected macro payload.");
                    var text = new StringBuilder(); int formulas = 0, formulaErrors = 0;
                    foreach (var part in package.Entries.Where(e => e.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) &&
                        (e.FullName.StartsWith("xl/worksheets/", StringComparison.Ordinal) || e.FullName.StartsWith("ppt/slides/", StringComparison.Ordinal) ||
                         e.FullName.StartsWith("ppt/notesSlides/", StringComparison.Ordinal) || e.FullName == "word/document.xml" || e.FullName == "xl/sharedStrings.xml")))
                    {
                        using (var stream = part.Open())
                        {
                            var document = XDocument.Load(stream); text.AppendLine(string.Join(" ", document.Descendants().Select(n => n.Value)));
                            if (part.FullName.StartsWith("xl/worksheets/", StringComparison.Ordinal)) {
                                formulas += document.Descendants().Count(n => n.Name.LocalName == "f");
                                formulaErrors += document.Descendants().Count(n => n.Name.LocalName == "c" && (string)n.Attribute("t") == "e");
                            }
                        }
                    }
                    var value = text.ToString();
                    return new OfficeInspection { text = value, numbers = Numbers(value), formulas = formulas,
                        formulaErrors = formulaErrors, charts = package.Entries.Count(e => Regex.IsMatch(e.FullName, @"(?:xl|ppt)/.+charts/chart\d+\.xml$")),
                        slides = package.Entries.Count(e => Regex.IsMatch(e.FullName, @"^ppt/slides/slide\d+\.xml$")) };
                }
            }
        }

        private static List<double> Numbers(string text)
        {
            var result = new List<double>();
            foreach (Match match in Regex.Matches(text ?? "", @"(?<![A-Za-z])[-+\u2212]?\d[\d,]*(?:\.\d+)?"))
            { double value; if (double.TryParse(match.Value.Replace(",", "").Replace('\u2212', '-'), NumberStyles.Float, CultureInfo.InvariantCulture, out value)) result.Add(value); }
            return result;
        }

        private static string Read(ZipArchive archive, string name)
        { var entry = archive.GetEntry(name); if (entry == null) return ""; using (var reader = new StreamReader(entry.Open())) return reader.ReadToEnd(); }

        private static void Add(List<TestLabCheck> checks, string name, bool passed, string detail)
        { checks.Add(new TestLabCheck { name = name, passed = passed, detail = detail, hard = true }); }

        private sealed class OfficeInspection
        { public string text; public List<double> numbers; public int formulas, formulaErrors, charts, slides; }
    }

    public sealed class TestLabEvaluation
    {
        public int schema { get; set; }
        public string run_id { get; set; }
        public string case_id { get; set; }
        public string evidence_sha256 { get; set; }
        public string status { get; set; }
        public bool passed { get; set; }
        public TestLabCheck[] checks { get; set; }
        public string[] findings { get; set; }
        public string note { get; set; }
    }

    public sealed class TestLabCheck
    { public string name { get; set; } public bool passed { get; set; } public string detail { get; set; } public bool hard { get; set; } }
}
