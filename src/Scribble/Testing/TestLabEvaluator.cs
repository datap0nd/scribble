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
                    if (memory.extension == "xlsx" || memory.extension == "pptx" || memory.extension == "docx")
                        Add(checks, "output_boundary_" + memory.name,
                            memory.outputBoundary || (!memory.invalidBoundary && LegacyBoundaryClear(archive, memory.name)),
                            "Output-only native evidence is required when the captured source boundary is missing or already contains draft markers. An invalid boundary flag cannot establish ownership.");
                    var output = memory.outputBoundary ? memory.text : OutputText(memory.extension, memory.text);
                    finalText.AppendLine(output);
                    if (run.case_id == "OL02" && memory.extension == "msg")
                        checks.AddRange(CheckMailReadback(memory.metadata));
                }
                foreach (var entry in archive.Entries.Where(e => e.FullName.StartsWith("artifacts/", StringComparison.Ordinal) &&
                    e.FullName.IndexOf("-final-output-", StringComparison.OrdinalIgnoreCase) >= 0 &&
                    new[] { ".xlsx", ".pptx", ".docx" }.Contains(Path.GetExtension(e.FullName).ToLowerInvariant())))
                {
                    try
                    {
                        var inspection = InspectOffice(entry);
                        // The disk package may also contain unchanged source
                        // sheets/slides. Grade the captured output boundary once;
                        // keep the full native package available for inspection.
                        if (!memoryOutputs.Any(m => SameOutput(m.name, entry.Name)))
                        {
                            Add(checks, "output_boundary_" + entry.Name, LegacyBoundaryClear(archive, entry.Name),
                                "File-only evidence needs a captured source boundary without pre-existing draft markers; otherwise a native output-only readback is required.");
                            finalText.AppendLine(inspection.text);
                            AddNamedChecks(checks, CheckOutputStructure(run.case_id, Path.GetExtension(entry.Name).TrimStart('.'), inspection.text, true), entry.Name);
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
                checks.AddRange(CheckOutputFacts(run.case_id, finalText.ToString(), false));
                if (run.case_id == "OL02" && !memoryOutputs.Any(m => m.extension == "msg"))
                    Add(checks, "native_mail_headers", false, "Native draft recipient and unsent metadata are missing.");
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
                object boundary;
                var hasBoundary = data.TryGetValue("output_boundary", out boundary);
                yield return new MemoryOutput {
                    extension = Convert.ToString(extension).ToLowerInvariant(),
                    text = Convert.ToString(text) ?? "",
                    name = entry.Name,
                    outputBoundary = hasBoundary && boundary is bool && (bool)boundary,
                    invalidBoundary = hasBoundary && (!(boundary is bool) || !(bool)boundary),
                    metadata = data
                };
            }
        }

        private static void AddMemoryStructureChecks(
            List<TestLabCheck> checks,
            string caseId,
            MemoryOutput[] outputs)
        {
            foreach (var output in outputs)
                AddNamedChecks(checks, CheckOutputStructure(caseId, output.extension, output.text, output.outputBoundary), output.name);
        }

        private static bool SameOutput(string first, string second)
        {
            const string pattern = @"(Excel|PowerPoint|Outlook|Word)-final-output-(\d+)(?:-readback\.json|\.[^.]+)$";
            var left = Regex.Match(first, pattern, RegexOptions.IgnoreCase); var right = Regex.Match(second, pattern, RegexOptions.IgnoreCase);
            return left.Success && right.Success && string.Equals(left.Groups[1].Value, right.Groups[1].Value, StringComparison.OrdinalIgnoreCase) &&
                left.Groups[2].Value == right.Groups[2].Value;
        }

        private static void AddNamedChecks(List<TestLabCheck> checks, TestLabCheck[] additions, string name)
        {
            foreach (var check in additions) { check.name += "_" + name; checks.Add(check); }
        }

        private static bool LegacyBoundaryClear(ZipArchive archive, string name)
        {
            var host = Regex.Match(name, @"(Excel|PowerPoint|Word)-final-output-", RegexOptions.IgnoreCase).Groups[1].Value;
            var sources = Readbacks(archive, "-source-source-").Where(p => p.Key.StartsWith(host + "-", StringComparison.OrdinalIgnoreCase)).Select(p => p.Value).ToArray();
            return sources.Length > 0 && sources.All(s => !s.Contains("Worksheet: Scribble Draft") && !s.Contains("[Scribble draft]"));
        }

        public static TestLabCheck[] CheckOutputStructure(string caseId, string extension, string text, bool outputBoundary = false)
        {
            var checks = new List<TestLabCheck>();
            var output = outputBoundary ? text ?? "" : OutputText(extension, text);
            if (extension == "xlsx")
            {
                var formulas = Regex.Matches(output, @"(?m)^R\d+C\d+: [^\r\n]*\| formula: =").Count;
                var formulaErrors = Regex.Matches(
                    NormalizeExcelErrors(output),
                    @"(?m)^R\d+C\d+: #(?:REF!|NAME\?|VALUE!|DIV/0!|NULL!|NUM!|N/A|SPILL!|CALC!|GETTING_DATA|EXCEL_ERROR\(\d+\))(?=\s*(?:\||$))",
                    RegexOptions.IgnoreCase).Count;
                var charts = Regex.Matches(output, @"(?m)^Native charts: ([1-9][0-9]*)\r?$").Count;
                Add(checks, "no_formula_errors", formulaErrors == 0,
                    "Native formula error cells in the output boundary: " + formulaErrors);
                if (new[] { "EX01", "EX04", "EX05", "XA01", "RC01" }.Contains(caseId))
                    Add(checks, "native_formulas", formulas > 0,
                        "Formula count in the output boundary: " + formulas);
                if (new[] { "EX01", "EX04", "XA01", "XA03", "RC01" }.Contains(caseId))
                    Add(checks, "native_chart", charts > 0,
                        "Chart-bearing output sheets: " + charts);
            }
            if (extension == "pptx")
            {
                var match = Regex.Match(output, outputBoundary ? @"(?m)^Slide count: (\d+)\r?$" : @"(?m)^Draft slide count: (\d+)\r?$");
                var observed = match.Success ? Convert.ToInt32(match.Groups[1].Value) : 0;
                var expected = caseId == "PP02" ? 4 : 6;
                Add(checks, "final_slide_count", observed == expected,
                    "Expected " + expected + " new draft slides; observed " + observed + ". Preserved source slides do not count.");
                if (caseId == "PP04") Add(checks, "native_chart", Regex.IsMatch(output, @"(?m)^Native chart type: "),
                    "The financial review must contain a native editable chart in its draft slides.");
            }
            return checks.ToArray();
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
        { return CheckOutputFacts(caseId, output, true); }

        private static TestLabCheck[] CheckOutputFacts(string caseId, string output, bool checkMailHeaders)
        {
            var checks = new List<TestLabCheck>();
            // Cell coordinates and formula operands are not calculated answers.
            var values = Regex.Replace(output ?? "", @"(?m)^R\d+C\d+: ([^\r\n]*)", m =>
                m.Groups[1].Value.Split('|')[0]);
            values = Regex.Replace(values, @"(?m)^(?:Worksheet: |Shape \d+ \||Native charts?: |Native chart type: |Chart data readback: |Series \d+: |Chart \d+ \||Slide count: |Draft slide count: |Slide \d+\r?$)[^\r\n]*", "");
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
                if (checkMailHeaders) checks.AddRange(CheckLegacyMailHeaders(output ?? ""));
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
            var clauses = Regex.Split(output, @"[;\r\n]+|\.\s+|,\s*(?=(?:Mira Cole|Leon Park)\b)", RegexOptions.IgnoreCase);
            Add(checks, "action_date_" + owner.Replace(" ", "_"), clauses.Any(clause => clause.Length <= 500 &&
                clause.IndexOf(owner, StringComparison.OrdinalIgnoreCase) >= 0 && clause.IndexOf(other, StringComparison.OrdinalIgnoreCase) < 0 &&
                Regex.IsMatch(clause, date, RegexOptions.IgnoreCase)), owner + " must retain the due date " + day + " July 2026.");
        }

        public static TestLabCheck[] CheckMailReadback(string readbackJson)
        {
            return CheckMailReadback(new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(readbackJson));
        }

        private static TestLabCheck[] CheckMailReadback(Dictionary<string, object> data)
        {
            // New captures keep native headers outside the model-authored body.
            // Never fall back to quoted body headers when structured fields exist.
            if (data.ContainsKey("to") || data.ContainsKey("cc") || data.ContainsKey("bcc"))
                return CheckMailHeaders(Field(data, "to"), Field(data, "cc"), Field(data, "bcc"), Field(data, "unsent"));
            return CheckLegacyMailHeaders(Field(data, "text"));
        }

        private static string Field(Dictionary<string, object> data, string name)
        { object value; return data.TryGetValue(name, out value) ? Convert.ToString(value) ?? "" : ""; }

        private static TestLabCheck[] CheckLegacyMailHeaders(string text)
        {
            var headers = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            var lines = (text ?? "").Split('\n');
            // The collector writes one fixed header block, then the body. Stop
            // at its Unsent field so a body line cannot impersonate native data.
            if (lines.Length > 0 && lines[0].StartsWith("To:", StringComparison.OrdinalIgnoreCase))
                foreach (var line in lines)
                {
                    var match = Regex.Match(line, @"^(To|CC|BCC|Subject|Unsent):\s*(.*?)\r?$", RegexOptions.IgnoreCase);
                    if (!match.Success || headers.ContainsKey(match.Groups[1].Value)) break;
                    headers.Add(match.Groups[1].Value, match.Groups[2].Value);
                    if (match.Groups[1].Value.Equals("Unsent", StringComparison.OrdinalIgnoreCase)) break;
                }
            return CheckMailHeaders(Field(headers, "To"), Field(headers, "CC"), Field(headers, "BCC"), Field(headers, "Unsent"));
        }

        private static TestLabCheck[] CheckMailHeaders(string to, string cc, string bcc, string unsent)
        {
            var checks = new List<TestLabCheck>();
            Add(checks, "draft_recipient", string.Equals(to.Trim(), "review@example.test", StringComparison.OrdinalIgnoreCase) &&
                string.IsNullOrWhiteSpace(cc) && string.IsNullOrWhiteSpace(bcc),
                "The native draft must be addressed only to review@example.test, with no CC or BCC recipients.");
            Add(checks, "draft_unsent", string.Equals(unsent.Trim(), "true", StringComparison.OrdinalIgnoreCase),
                "The recorded native draft must remain unsent.");
            return checks.ToArray();
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
                        metric[1] + @"[\s*]*(?:\(\s*(?:EUR|€|%|pp|percentage points)\s*\)[\s*]*)?[:\s*]*(?:(?:EUR|€|is|was|of|at|equals|=)\s*)*(?<value>[-+\u2212]?\d[\d,]*(?:\.\d+)?)", RegexOptions.IgnoreCase))
                    {
                        var prefix = line.Substring(0, match.Groups["value"].Index);
                        var periods = Regex.Matches(prefix, @"\b(?:May|June)\b", RegexOptions.IgnoreCase);
                        if (periods.Count > 0 && periods[periods.Count - 1].Value.Equals("May", StringComparison.OrdinalIgnoreCase)) continue;
                        var suffix = line.Substring(match.Index + match.Length);
                        if (Regex.IsMatch(suffix, @"^\s*%?\s*(?:(?:in|for|during)\s+|\()May\b", RegexOptions.IgnoreCase)) continue;
                        if (Regex.IsMatch(line.Substring(0, match.Index), @"\b(?:North|South|Product\s+[AB])\s*$", RegexOptions.IgnoreCase)) continue;
                        observed.AddRange(Numbers(match.Groups["value"].Value).Select(n =>
                            metric[0].StartsWith("margin", StringComparison.Ordinal) && Math.Abs(n) < 1 ? n * 100 : n));
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
                    // Read actual leaf values in draft parts only. XML ancestor
                    // Values duplicate text and concatenate formula operands with
                    // cached results, allowing source facts to pass as answers.
                    var text = entry.FullName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ? ReadPackageExcel(package) :
                        entry.FullName.EndsWith(".pptx", StringComparison.OrdinalIgnoreCase) ? ReadPackageSlides(package) :
                        LeafText(PackageXml(package, "word/document.xml"));
                    return new OfficeInspection { text = text };
                }
            }
        }

        private static XDocument PackageXml(ZipArchive package, string path)
        {
            var entry = package.GetEntry(path); if (entry == null) return null;
            using (var stream = entry.Open()) return XDocument.Load(stream);
        }

        private static string LeafText(XDocument document)
        { return document == null ? "" : string.Join("\n", document.Descendants().Where(n => n.Name.LocalName == "t").Select(n => n.Value)); }

        private static Dictionary<string, string> PackageRelationships(ZipArchive package, string part)
        {
            var slash = part.LastIndexOf('/');
            var rels = PackageXml(package, part.Substring(0, slash + 1) + "_rels/" + part.Substring(slash + 1) + ".rels");
            var result = new Dictionary<string, string>();
            if (rels != null) foreach (var item in rels.Descendants().Where(n => n.Name.LocalName == "Relationship" && (string)n.Attribute("TargetMode") != "External"))
            {
                var id = (string)item.Attribute("Id"); var target = (string)item.Attribute("Target");
                if (id != null && target != null)
                    result[id] = new Uri(new Uri("https://package.invalid/" + part), target).AbsolutePath.TrimStart('/');
            }
            return result;
        }

        private static string RelationshipTarget(XElement element, Dictionary<string, string> relationships)
        {
            var id = (string)element.Attribute(XName.Get("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"));
            string target; return id != null && relationships.TryGetValue(id, out target) ? target : null;
        }

        private static int PackageCharts(ZipArchive package, string part, XDocument document, int depth = 0)
        {
            if (document == null || depth > 2) return 0;
            var relationships = PackageRelationships(package, part); int count = 0;
            foreach (var node in document.Descendants().Where(n => n.Name.LocalName == "chart" || n.Name.LocalName == "drawing"))
            {
                var target = RelationshipTarget(node, relationships); if (target == null) continue;
                if (node.Name.LocalName == "chart") { if (package.GetEntry(target) != null) count++; }
                else count += PackageCharts(package, target, PackageXml(package, target), depth + 1);
            }
            return count;
        }

        private static string ReadPackageExcel(ZipArchive package)
        {
            var text = new StringBuilder(); var workbook = PackageXml(package, "xl/workbook.xml");
            if (workbook == null) return "";
            var relationships = PackageRelationships(package, "xl/workbook.xml");
            var shared = PackageXml(package, "xl/sharedStrings.xml");
            var strings = shared == null ? new string[0] : shared.Descendants().Where(n => n.Name.LocalName == "si")
                .Select(n => string.Concat(n.Descendants().Where(t => t.Name.LocalName == "t").Select(t => t.Value))).ToArray();
            foreach (var sheet in workbook.Descendants().Where(n => n.Name.LocalName == "sheet"))
            {
                var name = (string)sheet.Attribute("name") ?? "";
                if (name != "Scribble Draft" && !name.StartsWith("Scribble Draft ", StringComparison.Ordinal)) continue;
                var part = RelationshipTarget(sheet, relationships); var document = part == null ? null : PackageXml(package, part);
                if (document == null) continue;
                text.AppendLine("Worksheet: " + name);
                foreach (var cell in document.Descendants().Where(n => n.Name.LocalName == "c"))
                {
                    var formula = cell.Elements().FirstOrDefault(n => n.Name.LocalName == "f");
                    var value = cell.Elements().FirstOrDefault(n => n.Name.LocalName == "v"); var contents = value == null ? "" : value.Value;
                    var type = (string)cell.Attribute("t");
                    if (type == "s") { int index; contents = int.TryParse(contents, out index) && index >= 0 && index < strings.Length ? strings[index] : ""; }
                    else if (type == "inlineStr") contents = string.Concat(cell.Descendants().Where(n => n.Name.LocalName == "t").Select(n => n.Value));
                    var coordinate = Regex.Match((string)cell.Attribute("r") ?? "A1", @"^([A-Z]+)(\d+)$"); int column = 0;
                    foreach (var letter in coordinate.Groups[1].Value) column = column * 26 + letter - 'A' + 1;
                    text.Append("R").Append(coordinate.Groups[2].Value).Append("C").Append(column).Append(": ").Append(contents);
                    if (formula != null) text.Append(" | formula: =").Append(formula.Value);
                    text.AppendLine();
                }
                text.AppendLine("Native charts: " + PackageCharts(package, part, document));
            }
            return text.ToString();
        }

        private static string ReadPackageSlides(ZipArchive package)
        {
            var slides = new List<string>();
            foreach (var part in package.Entries.Where(e => Regex.IsMatch(e.FullName, @"^ppt/slides/slide\d+\.xml$")))
            {
                var document = PackageXml(package, part.FullName); var text = LeafText(document);
                if (!text.Contains("[Scribble draft]")) continue;
                var slide = new StringBuilder(text);
                for (var n = 0; n < PackageCharts(package, part.FullName, document); n++) slide.AppendLine("\nNative chart type: editable OOXML chart");
                var relationships = PackageRelationships(package, part.FullName);
                foreach (var chart in document.Descendants().Where(n => n.Name.LocalName == "chart").Take(32))
                {
                    var target = RelationshipTarget(chart, relationships); var data = target == null ? null : PackageXml(package, target);
                    if (data == null) continue;
                    foreach (var series in data.Descendants().Where(n => n.Name.LocalName == "ser").Take(32))
                    {
                        var name = series.Elements().FirstOrDefault(n => n.Name.LocalName == "tx");
                        var category = series.Elements().FirstOrDefault(n => n.Name.LocalName == "cat" || n.Name.LocalName == "xVal");
                        var value = series.Elements().FirstOrDefault(n => n.Name.LocalName == "val" || n.Name.LocalName == "yVal");
                        slide.AppendLine("\n" + BenchmarkArtifactCollector.FormatChartSeries(name == null ? "" : name.Descendants().FirstOrDefault(n => n.Name.LocalName == "v")?.Value,
                            CachedChartPoints(category), CachedChartPoints(value)));
                    }
                }
                foreach (var notes in PackageRelationships(package, part.FullName).Values.Where(p => p.StartsWith("ppt/notesSlides/", StringComparison.Ordinal)))
                    slide.AppendLine("\nNotes: " + LeafText(PackageXml(package, notes)));
                slides.Add(slide.ToString());
            }
            return "Slide count: " + slides.Count + "\n" + string.Join("\n", slides.Select((s, i) => "Slide " + (i + 1) + "\n" + s));
        }

        private static object[] CachedChartPoints(XElement parent)
        {
            if (parent == null) return new object[0];
            var points = new object[128]; var count = 0;
            foreach (var point in parent.Descendants().Where(n => n.Name.LocalName == "pt").Take(128))
            {
                int index; if (!int.TryParse((string)point.Attribute("idx"), out index) || index < 0 || index >= points.Length) continue;
                points[index] = point.Elements().FirstOrDefault(n => n.Name.LocalName == "v")?.Value;
                count = Math.Max(count, index + 1);
            }
            return points.Take(count).ToArray();
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
        { public string text; }
        private sealed class MemoryOutput
        { public string extension; public string text; public string name; public bool outputBoundary, invalidBoundary; public Dictionary<string, object> metadata; }
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
