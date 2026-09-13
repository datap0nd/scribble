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
            LabRun run; LabCase testCase; var finalText = new StringBuilder();
            using (var archive = ZipFile.OpenRead(evidence))
            {
                run = json.Deserialize<LabRun>(Read(archive, "run.json"));
                testCase = TestLabSuite.Read<LabCase[]>(TestLab.SafeChild(run.fixture_root, "operator/cases.json")).Single(c => c.id == run.case_id);
                Add(checks, "trace_complete", run.trace_complete && run.status == "finished", "The capture must reach a terminal case event.");
                Add(checks, "unassisted", !run.assisted, "Manual intervention is retained but cannot count as an unassisted pass.");
                foreach (var extension in testCase.artifacts ?? new string[0])
                {
                    var suffix = "." + extension;
                    Add(checks, "required_final_" + extension,
                        archive.Entries.Any(e =>
                            e.FullName.StartsWith("artifacts/", StringComparison.Ordinal) &&
                            e.FullName.IndexOf("-final-output-", StringComparison.OrdinalIgnoreCase) >= 0 &&
                            e.FullName.EndsWith(suffix, StringComparison.OrdinalIgnoreCase)) ||
                        HasMemoryOutput(archive, extension),
                        "Only a final run-owned native output or in-memory native readback satisfies this requirement.");
                }
                var memoryOutputs = MemoryOutputs(archive).ToArray();
                foreach (var memory in memoryOutputs)
                {
                    var output = OutputText(memory.extension, memory.text);
                    finalText.AppendLine(output);
                }
                foreach (var entry in archive.Entries.Where(e => e.FullName.StartsWith("artifacts/", StringComparison.Ordinal) &&
                    e.FullName.IndexOf("-final-output-", StringComparison.OrdinalIgnoreCase) >= 0 &&
                    new[] { ".xlsx", ".pptx", ".docx" }.Contains(Path.GetExtension(e.FullName).ToLowerInvariant())))
                {
                    try
                    {
                        var inspection = InspectOffice(entry);
                        if (!memoryOutputs.Any(m => "." + m.extension == Path.GetExtension(entry.Name))) finalText.AppendLine(inspection.text);
                        if (entry.FullName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                        {
                            Add(checks, "no_formula_errors_" + entry.Name, inspection.formulaErrors == 0, "Native formula error cells: " + inspection.formulaErrors);
                            if (new[] { "EX01", "EX04", "EX05", "XA01", "RC01" }.Contains(run.case_id))
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
                AddMemoryStructureChecks(checks, run.case_id, memoryOutputs);
                if ((testCase.artifacts ?? new string[0]).Length == 0)
                {
                    // Input sources and tool observations already contain the
                    // right numbers. Only the final assistant answer can satisfy
                    // a chat-only test, never a copied source or user prompt.
                    var answer = FinalAnswer(Read(archive, "timeline.jsonl"));
                    finalText.AppendLine(answer);
                }
                checks.AddRange(CheckOutputFacts(run.case_id, finalText.ToString()));
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

        private static bool HasMemoryOutput(ZipArchive archive, string extension)
        {
            return MemoryOutputs(archive).Any(c => string.Equals(
                c.extension,
                extension,
                StringComparison.OrdinalIgnoreCase));
        }
        public static string FinalAnswer(string timeline)
        {
            var answer = "";
            var json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
            foreach (var line in (timeline ?? "").Split('\n').Where(l => !string.IsNullOrWhiteSpace(l)))
            {
                var entry = json.Deserialize<Dictionary<string, object>>(line); object stage, detail, type, text;
                if (!entry.TryGetValue("stage", out stage) || Convert.ToString(stage) != "pane_event" ||
                    !entry.TryGetValue("detail", out detail)) continue;
                var payload = detail as Dictionary<string, object>;
                if (payload != null && payload.TryGetValue("type", out type) && Convert.ToString(type) == "assistant" &&
                    payload.TryGetValue("text", out text)) answer = Convert.ToString(text);
            }
            return answer;
        }

        public static string BlockingDetails(string timeline)
        {
            string question = null, rejected = null;
            var json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
            foreach (var line in (timeline ?? "").Split('\n').Where(l => !string.IsNullOrWhiteSpace(l)))
            {
                var entry = json.Deserialize<Dictionary<string, object>>(line); object stage, raw, value;
                if (!entry.TryGetValue("stage", out stage) || !entry.TryGetValue("detail", out raw)) continue;
                var detail = raw as Dictionary<string, object>; if (detail == null) continue;
                if (Convert.ToString(stage) == "pane_event" && detail.TryGetValue("type", out value) && Convert.ToString(value) == "askUser" && detail.TryGetValue("question", out value))
                    question = "Question: " + Convert.ToString(value);
                if (Convert.ToString(stage) == "argument_validation_failed")
                    rejected = "Last rejected tool call: " + json.Serialize(detail);
            }
            return string.Join("\n", new[] { question, rejected }.Where(v => v != null));
        }

        public static string OutputText(string extension, string text)
        {
            if (extension == "pptx") {
                var slides = Regex.Split(text ?? "", @"(?m)(?=^Slide \d+\r?$)");
                var draftSlides = slides.Where(s => s.Contains("[Scribble draft]")).ToArray();
                if (draftSlides.Length > 0) return "Draft slide count: " + draftSlides.Length + "\n" + string.Join("\n", draftSlides);
            }
            if (extension != "xlsx") return text ?? "";
            var sections = Regex.Split(text ?? "", @"(?m)(?=^Worksheet: )");
            var drafts = sections.Where(s => s.StartsWith("Worksheet: Scribble Draft", StringComparison.Ordinal)).ToArray();
            return drafts.Length > 0 ? string.Join("\n", drafts) : text ?? "";
        }

        private static IEnumerable<MemoryOutput> MemoryOutputs(ZipArchive archive)
        {
            var json = new JavaScriptSerializer();
            foreach (var entry in archive.Entries.Where(e =>
                e.FullName.StartsWith("artifacts/", StringComparison.Ordinal) &&
                e.FullName.IndexOf("-final-output-", StringComparison.OrdinalIgnoreCase) >= 0 &&
                e.FullName.EndsWith("-readback.json", StringComparison.OrdinalIgnoreCase)))
            {
                Dictionary<string, object> data;
                try { data = json.Deserialize<Dictionary<string, object>>(Read(archive, entry.FullName)); }
                catch { continue; }
                object native;
                object created;
                object extension;
                object text;
                if (!data.TryGetValue("native_readback", out native) ||
                    !Convert.ToBoolean(native) ||
                    !data.TryGetValue("run_created_output", out created) ||
                    !Convert.ToBoolean(created) ||
                    !data.TryGetValue("artifact_extension", out extension) ||
                    !data.TryGetValue("text", out text))
                    continue;
                yield return new MemoryOutput {
                    extension = Convert.ToString(extension),
                    text = Convert.ToString(text) ?? "",
                    name = entry.Name
                };
            }
        }

        private static void AddMemoryStructureChecks(
            List<TestLabCheck> checks,
            string caseId,
            MemoryOutput[] outputs)
        {
            foreach (var output in outputs.Where(o => o.extension == "xlsx"))
            {
                var formulas = Regex.Matches(output.text, @"\| formula: =").Count;
                var formulaErrors = Regex.Matches(
                    NormalizeExcelErrors(output.text),
                    @"#(?:REF!|NAME\?|VALUE!|DIV/0!|NULL!|NUM!|N/A|SPILL!|CALC!|GETTING_DATA|EXCEL_ERROR\()",
                    RegexOptions.IgnoreCase).Count;
                var charts = Regex.Matches(output.text, @"Native charts: ([1-9][0-9]*)").Count;
                Add(checks, "no_formula_errors_" + output.name, formulaErrors == 0,
                    "Native formula error cells in memory readback: " + formulaErrors);
                if (new[] { "EX01", "EX04", "EX05", "XA01", "RC01" }.Contains(caseId))
                    Add(checks, "native_formulas_" + output.name, formulas > 0,
                        "Formula count in memory readback: " + formulas);
                if (new[] { "EX01", "EX04", "XA01", "XA03", "RC01" }.Contains(caseId))
                    Add(checks, "native_chart_" + output.name, charts > 0,
                        "Chart-bearing sheets in memory readback: " + charts);
            }
            foreach (var output in outputs.Where(o => o.extension == "pptx"))
            {
                var match = Regex.Match(output.text, @"Slide count: (\d+)");
                var observed = match.Success ? Convert.ToInt32(match.Groups[1].Value) : 0;
                var expected = caseId == "PP02" ? 4 : 6;
                Add(checks, "final_slide_count_" + output.name,
                    observed == expected || observed == expected + 1,
                    "Expected " + expected + " draft slides (plus at most one preserved source); observed " + observed + ".");
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

        public static string NormalizeExcelErrors(string text)
        {
            // Older work-PC evidence serialized VT_ERROR as signed integers.
            // Only translate a cell value carrying a native formula receipt.
            return Regex.Replace(text ?? "", @"(?m)^(R\d+C\d+: )(-\d+)(?=\s*\| formula:)", m => {
                int code; return int.TryParse(m.Groups[2].Value, out code)
                    ? m.Groups[1].Value + (Scribble.Office.ExcelErrorValue.Text(code) ?? m.Groups[2].Value) : m.Value;
            });
        }

        public static TestLabCheck[] CheckOutputFacts(string caseId, string output)
        {
            var checks = new List<TestLabCheck>();
            // Cell coordinates and formula operands are not calculated answers.
            var values = Regex.Replace(output ?? "", @"(?m)^R\d+C\d+: ([^\r\n]*)", m =>
                m.Groups[1].Value.Split('|')[0]);
            var numbers = Numbers(values);
            var facts = new Dictionary<string, double[]> {
                { "EX01", new[] { 120000d, 130000d, 46000d } }, { "EX02", new[] { 120000d } },
                { "EX03", new[] { 95000d } }, { "EX04", new[] { 50000d, 18000d } },
                { "PP01", new[] { 120000d, 130000d, 94d } }, { "PP02", new[] { 120000d, 94d } },
                { "OL01", new[] { 120000d, 94d } }, { "XA01", new[] { 120000d, 130000d, 94d } },
                { "XA02", new[] { 120000d, 94d } }, { "CH01", new[] { 94d, 97d } },
                { "XA03", new[] { 94d } }, { "WD01", new[] { 120000d, 94d } },
                { "XA04", new[] { 120000d } }, { "RB01", new[] { 120000d } }, { "RC01", new[] { 120000d, 94d } },
                { "EX05", new[] { 120000d, 74000d, 46000d } }, { "PP04", new[] { 120000d, 130000d, 94d } },
                { "OL02", new[] { 120000d, 94d } } };
            var finance = new[] { "EX01", "PP01", "PP02", "PP04", "OL01", "XA01", "XA02", "XA04", "RC01" };
            if (finance.Contains(caseId))
                facts[caseId] = facts[caseId].Concat(new[] { 130000d, 74000d, 46000d, -10000d }).Distinct().ToArray();
            if (caseId == "OL02") facts[caseId] = facts[caseId].Concat(new[] { -10000d }).ToArray();
            double[] expected;
            if (!facts.TryGetValue(caseId, out expected)) return checks.ToArray();
            foreach (var value in expected) Add(checks, "required_final_fact_" + value.ToString(CultureInfo.InvariantCulture),
                numbers.Any(n => Math.Abs(n - value) < .011), "Expected value was not found in the final run-owned evidence boundary.");
            if (finance.Contains(caseId) || caseId == "EX05" || caseId == "OL02")
            {
                RequirePercent(checks, numbers, "june_margin", 38.33);
                if (finance.Contains(caseId)) { RequirePercent(checks, numbers, "growth", 20); RequirePercent(checks, numbers, "budget_gap", -7.69, true); }
            }
            if (new[] { "PP01", "PP02", "PP04", "OL01", "XA01", "XA02", "XA04", "RC01", "OL02" }.Contains(caseId))
            {
                RequirePercent(checks, numbers, "delivery", 94);
                RequirePercent(checks, numbers, "delivery_target", 97);
            }
            if (caseId == "OL01" || caseId == "OL02") CheckFinanceStatements(checks, output ?? "");
            if (caseId == "OL02")
            {
                Add(checks, "draft_recipient", Regex.IsMatch(output ?? "", @"(?im)^To:\s*review@example\.test\s*$"), "The draft must be addressed only to review@example.test.");
                Add(checks, "draft_unsent", Regex.IsMatch(output ?? "", @"(?im)^Unsent:\s*true\s*$"), "The recorded native draft must remain unsent.");
                foreach (var owner in new[] { "Mira Cole", "Leon Park" })
                    Add(checks, "action_owner_" + owner.Replace(" ", "_"), (output ?? "").IndexOf(owner, StringComparison.OrdinalIgnoreCase) >= 0, "Required planned action owner is missing: " + owner);
                CheckActionDate(checks, output ?? "", "Mira Cole", "Leon Park", 10);
                CheckActionDate(checks, output ?? "", "Leon Park", "Mira Cole", 12);
            }
            return checks.ToArray();
        }

        private static void CheckActionDate(List<TestLabCheck> checks, string output, string owner, string other, int day)
        {
            var date = @"(?:\b" + day + @"(?:th)?\s+July\s+2026\b|\bJuly\s+" + day + @"(?:th)?,?\s+2026\b|\b2026-07-" + day + @"\b)";
            Add(checks, "action_date_" + owner.Replace(" ", "_"), Regex.IsMatch(output,
                Regex.Escape(owner) + @"(?:(?!" + Regex.Escape(other) + @").){0,240}?" + date,
                RegexOptions.IgnoreCase | RegexOptions.Singleline), owner + " must retain the due date " + day + " July 2026.");
        }

        private static void RequirePercent(List<TestLabCheck> checks, List<double> numbers, string name, double expected, bool magnitude = false)
        {
            Add(checks, "required_percentage_" + name, numbers.Any(n =>
                Math.Abs((magnitude ? Math.Abs(n) : n) - (magnitude ? Math.Abs(expected) : expected)) < .011 ||
                (Math.Abs(n) < 1 && Math.Abs((magnitude ? Math.Abs(n * 100) : n * 100) - (magnitude ? Math.Abs(expected) : expected)) < .011)),
                "Expected " + name + ": " + expected.ToString(CultureInfo.InvariantCulture) + "%; a worksheet fraction is also accepted.");
        }

        private static void CheckFinanceStatements(List<TestLabCheck> checks, string output)
        {
            // Check the association, even when a correct number appears elsewhere
            // in the same draft. Only explicit metric statements are recognized;
            // full visual/semantic review is still required for other layouts.
            var metrics = new[] {
                new[] { "profit", @"(?:gross\s+profit|\bGP\b)", "46000" },
                new[] { "margin", @"\b(?:weighted\s+)?(?:gross\s+)?margin\b", "38.33" },
                new[] { "revenue", @"(?:June\s+)?revenue", "120000" },
                new[] { "cost", @"(?:June\s+)?(?:total\s+)?costs?", "74000" },
                new[] { "margin_change", @"\bmargin\s+(?:change|delta)\b", "-1.67" }
            };
            foreach (var metric in metrics)
            {
                var observed = new List<double>();
                foreach (var line in Regex.Split(output, @"[\r\n;]+|\.\s+(?=[A-Z])"))
                {
                    foreach (Match match in Regex.Matches(line,
                        metric[1] + @"[\s:*]*(?:(?:EUR|€|is|was|of|at|equals|=)\s*)*(?<value>[-+\u2212]?\d[\d,]*(?:\.\d+)?)", RegexOptions.IgnoreCase))
                    {
                        var prefix = line.Substring(0, match.Groups["value"].Index);
                        var periods = Regex.Matches(prefix, @"\b(?:May|June)\b", RegexOptions.IgnoreCase);
                        if (periods.Count > 0 && periods[periods.Count - 1].Value.Equals("May", StringComparison.OrdinalIgnoreCase)) continue;
                        if (Regex.IsMatch(line.Substring(0, match.Index), @"\b(?:North|South|Product\s+[AB])\s*$", RegexOptions.IgnoreCase)) continue;
                        observed.AddRange(Numbers(match.Groups["value"].Value));
                    }
                }
                var expected = double.Parse(metric[2], CultureInfo.InvariantCulture);
                Add(checks, "consistent_june_" + metric[0], observed.All(n => Math.Abs(n - expected) < .011),
                    "June " + metric[0] + " must be " + metric[2] + ". Recognized values: " + string.Join(", ", observed.Select(n => n.ToString(CultureInfo.InvariantCulture))) + ".");
            }
            var gaps = Regex.Matches(output, @"\bbudget\s+(?:gap|variance|shortfall)\b[^\r\n;%]{0,65}?(?<value>[-+\u2212]?\d+(?:\.\d+)?)\s*%", RegexOptions.IgnoreCase)
                .Cast<Match>().SelectMany(m => Numbers(m.Groups["value"].Value)).ToArray();
            Add(checks, "consistent_budget_gap_percentage", gaps.All(n => Math.Abs(Math.Abs(n) - 7.69) < .011),
                "An explicitly stated June budget gap percentage must be 7.69% below budget.");
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
                        formulaErrors = formulaErrors, charts = package.Entries.Count(e => Regex.IsMatch(e.FullName, @"^(?:xl|ppt)/charts/chart\d+\.xml$")),
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
        private sealed class MemoryOutput
        { public string extension; public string text; public string name; }
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
