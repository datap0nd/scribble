using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;

namespace Scribble.Testing
{
    public static class TestLabStressEvaluator
    {
        public static TestLabCheck[] Evaluate(ZipArchive archive, LabRun run, LabCase testCase, IEnumerable<TestLabCheck> existing)
        {
            var results = new List<TestLabCheck>();
            try
            {
                var manifest = TestLab.VerifyKit(run.fixture_root);
                if (manifest.suite_id != "scribble-stress-v1" || TestLab.FileHash(Path.Combine(run.fixture_root, "manifest.json")) != run.manifest_sha256)
                    throw new InvalidDataException("The captured stress manifest changed.");
                if (string.IsNullOrEmpty(testCase.oracle_ref) || !testCase.oracle_ref.StartsWith("evaluator-only/cases/", StringComparison.Ordinal) ||
                    !manifest.files.Any(f => f.path == testCase.oracle_ref)) throw new InvalidDataException("A verified evaluator-only oracle is required.");
                var json = new JavaScriptSerializer { MaxJsonLength = 2 * 1024 * 1024 };
                var oracle = json.Deserialize<Dictionary<string, object>>(File.ReadAllText(TestLab.SafeChild(run.fixture_root, testCase.oracle_ref)));
                if (Text(oracle, "id") != run.case_id) throw new InvalidDataException("Oracle case identity mismatch.");
                var captures = new List<StressReadback>();
                foreach (var entry in archive.Entries.Where(e => e.FullName.StartsWith("artifacts/", StringComparison.Ordinal) &&
                    e.FullName.IndexOf("-final-output-", StringComparison.Ordinal) >= 0 && e.FullName.EndsWith("-readback.json", StringComparison.Ordinal)))
                {
                    using (var reader = new StreamReader(entry.Open()))
                    {
                        var capture = json.Deserialize<StressReadback>(reader.ReadToEnd());
                        if (capture.run_id == run.run_id && capture.native_readback && capture.run_created_output && capture.output_boundary)
                            captures.Add(capture);
                    }
                }
                string timeline = "";
                var timelineEntry = archive.GetEntry("timeline.jsonl");
                if (timelineEntry != null) using (var reader = new StreamReader(timelineEntry.Open())) timeline = reader.ReadToEnd();
                var answer = TestLabEvaluator.FinalAnswer(timeline);
                var text = string.Join("\n", captures.Select(c => c.text)) + (testCase.artifacts.Length == 0 ? "\n" + answer : "");
                var native = captures.Where(c => c.stress_native != null).Select(c => c.stress_native).ToArray();
                foreach (var warning in native.SelectMany(n => n.measurement_warnings ?? new string[0]).Distinct())
                    results.Add(new TestLabCheck { name = "stress_native_measurement_review", passed = false, hard = false, detail = warning });
                var rules = Items(Value(oracle, "checks")).Cast<Dictionary<string, object>>().ToArray();
                if (rules.Length == 0) throw new InvalidDataException("The stress case has no independent checks.");
                var index = 0;
                foreach (var rule in rules)
                {
                    var kind = Text(rule, "kind"); string reason;
                    bool passed;
                    try { passed = Check(rule, captures.ToArray(), native, text, answer, timeline, existing.ToArray(), run, out reason); }
                    catch (Exception error) { passed = false; reason = "Cannot evaluate rule: " + error.Message; }
                    results.Add(new TestLabCheck { name = "stress_" + (++index) + "_" + kind, passed = passed, hard = true, detail = reason });
                }
            }
            catch (Exception error) { results.Add(new TestLabCheck { name = "stress_oracle_integrity", passed = false, hard = true, detail = error.Message }); }
            return results.ToArray();
        }

        private static bool Check(Dictionary<string, object> rule, StressReadback[] captures, StressNative[] native,
            string text, string answer, string timeline, TestLabCheck[] existing, LabRun run, out string reason)
        {
            var kind = Text(rule, "kind");
            reason = "Independent " + kind + " check.";
            switch (kind)
            {
                case "native_artifact":
                    reason = "A final output-only native " + Text(rule, "extension") + " readback is required.";
                    return captures.Any(c => c.artifact_extension == Text(rule, "extension"));
                case "source_unchanged":
                    var path = Text(rule, "path");
                    reason = "Verified fixture bytes and any opened native source must remain unchanged: " + path;
                    return run.input_paths.Contains(path) && existing.Where(c => c.name.StartsWith("source_preserved_", StringComparison.Ordinal)).All(c => c.passed);
                case "numeric_cell":
                    var wantedSheet = Text(rule, "sheet"); var address = Text(rule, "cell").ToUpperInvariant();
                    var match = Regex.Match(address, "^([A-Z]{1,3})([1-9][0-9]*)$");
                    if (!match.Success) throw new InvalidDataException("Invalid expected cell address.");
                    int column = 0; foreach (var ch in match.Groups[1].Value) column = column * 26 + ch - 'A' + 1;
                    var row = int.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture);
                    var sheets = native.Where(n => n.host == "Excel" && string.IsNullOrEmpty(n.error)).SelectMany(n => n.sheets)
                        .Where(s => Regex.IsMatch(s.name, "^" + Regex.Escape(wantedSheet).Replace("\\*", ".*") + "$"));
                    var candidates = sheets.Where(s => s.recalculated).SelectMany(s => s.cells).Where(c => c.row == row && c.column == column).ToArray();
                    reason = wantedSheet + "!" + address + " must recalculate to " + Text(rule, "expected") +
                        (Flag(rule, "formula_required") ? " using a native formula." : ".");
                    return candidates.Length == 1 && Near(candidates[0].value, Value(rule, "expected"), Number(rule, "tolerance", .01)) &&
                        (!Flag(rule, "formula_required") || (candidates[0].formula ?? "").StartsWith("=", StringComparison.Ordinal));
                case "required_text":
                    reason = "Required facts must occur in the final answer or native output, never only in inputs.";
                    return Strings(Value(rule, "values")).Length > 0 && Strings(Value(rule, "values")).All(v => Contains(text, v));
                case "cell_text":
                    var textCells = OutputCells(native, Text(rule, "sheet"), Text(rule, "cell"));
                    reason = Text(rule, "sheet") + "!" + Text(rule, "cell") + " must contain the requested text: " + Text(rule, "expected");
                    return textCells.Length == 1 && string.Equals(Convert.ToString(textCells[0].value, CultureInfo.InvariantCulture).Trim(),
                        Text(rule, "expected").Trim(), StringComparison.OrdinalIgnoreCase);
                case "native_chart":
                    var chartHost = Text(rule, "host"); var chartExtension = Text(rule, "artifact_extension");
                    if (!(chartHost == "Excel" && chartExtension == "xlsx") && !(chartHost == "PowerPoint" && chartExtension == "pptx"))
                        throw new InvalidDataException("Native chart rules require an explicit valid host and artifact extension.");
                    reason = "Native " + chartHost + " chart categories and calculated series must match the oracle in order.";
                    var chartOutputs = captures.Where(c => c.artifact_extension == chartExtension && c.stress_native?.host == chartHost &&
                        string.IsNullOrEmpty(c.stress_native.error)).Select(c => c.stress_native);
                    var charts = chartOutputs.SelectMany(n => chartHost == "Excel" ? n.sheets.SelectMany(s => s.charts) : n.slides.SelectMany(s => s.charts));
                    return charts.Any(c => ChartMatches(c, rule));
                case "mail_search":
                    var expected = Strings(Value(rule, "expected_ids")).Distinct().OrderBy(x => x, StringComparer.Ordinal).ToArray();
                    var observed = Regex.Matches(answer ?? "", @"\bMAIL[0-9]{4}\b").Cast<Match>().Select(m => m.Value).Distinct().OrderBy(x => x, StringComparer.Ordinal).ToArray();
                    reason = "Final email IDs must exactly match the oracle (expected " + expected.Length + ", observed " + observed.Length + "); native mailbox search evidence is also required.";
                    return expected.SequenceEqual(observed) && (int)Number(rule, "expected_count", expected.Length) == expected.Length &&
                        ExactMailCount(answer, expected.Length) &&
                        CompletedMailboxSearch(timeline);
                case "presentation":
                    reason = "The new deck must have the requested slides, facts, native text fit, and the supplied Samsung theme.";
                    return captures.Where(c => c.artifact_extension == "pptx" && c.stress_native?.host == "PowerPoint" && string.IsNullOrEmpty(c.stress_native.error))
                        .Any(c => PresentationMatches(c.stress_native, rule, run, c.text));
                case "draft_mail":
                    reason = "Native unsent draft recipients and content must match the oracle.";
                    return captures.Where(c => c.artifact_extension == "msg" && c.unsent).Any(c =>
                        Headers(c.to, Strings(Value(rule, "to"))) && Headers(c.cc, Strings(Value(rule, "cc"))) && Headers(c.bcc, Strings(Value(rule, "bcc"))) &&
                        (string.IsNullOrEmpty(Text(rule, "subject_contains")) || Contains(c.subject, Text(rule, "subject_contains"))) &&
                        Strings(Value(rule, "body_contains")).All(v => Contains(c.body, v)) && BodyMetrics(c.body, rule));
                default:
                    reason = "Unsupported oracle rule: " + kind + ". This case cannot pass until its independent checker is implemented.";
                    return false;
            }
        }

        private static StressCell[] OutputCells(StressNative[] native, string sheet, string address)
        {
            var match = Regex.Match((address ?? "").ToUpperInvariant(), "^([A-Z]{1,3})([1-9][0-9]*)$");
            if (!match.Success) throw new InvalidDataException("Invalid expected cell address.");
            int column = 0; foreach (var ch in match.Groups[1].Value) column = column * 26 + ch - 'A' + 1;
            var row = int.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture);
            return native.Where(n => n.host == "Excel" && string.IsNullOrEmpty(n.error)).SelectMany(n => n.sheets)
                .Where(s => s.recalculated && Regex.IsMatch(s.name, "^" + Regex.Escape(sheet).Replace("\\*", ".*") + "$"))
                .SelectMany(s => s.cells).Where(c => c.row == row && c.column == column).ToArray();
        }

        private static bool ExactMailCount(string answer, int expected)
        {
            var lines = (answer ?? "").Split('\n').Where(line => Regex.IsMatch(line, @"\bTotal\s+matches\s*:", RegexOptions.IgnoreCase)).ToArray();
            if (lines.Length != 1) return false;
            var match = Regex.Match(lines[0].Trim(), @"^Total matches\s*:\s*([0-9]+)$", RegexOptions.IgnoreCase);
            int count;
            return match.Success && int.TryParse(match.Groups[1].Value, out count) && count == expected;
        }

        private static bool BodyMetrics(string body, Dictionary<string, object> rule)
        {
            var metrics = Items(Value(rule, "body_metrics")).Cast<Dictionary<string, object>>().ToArray();
            if (metrics.Length == 0) return true;
            var labels = metrics.Select(m => Text(m, "label")).ToArray();
            if (labels.Any(string.IsNullOrWhiteSpace)) return false;
            var pattern = @"(?<!\w)(?:" + string.Join("|", labels.Select(Regex.Escape)) + @")(?!\w)";
            var occurrences = Regex.Matches(body ?? "", pattern, RegexOptions.IgnoreCase).Cast<Match>().ToArray();
            return metrics.All(metric => {
                var matches = occurrences.Select((match, index) => new { match, index })
                .Where(item => string.Equals(item.match.Value, Text(metric, "label"), StringComparison.OrdinalIgnoreCase))
                .ToArray();
                return matches.Length > 0 && matches.All(item => {
                    var start = item.match.Index + item.match.Length;
                    var end = item.index + 1 < occurrences.Length ? occurrences[item.index + 1].Index : body.Length;
                    var clause = body.Substring(start, Math.Min(200, end - start)).Split('\n', '\r', ';')[0];
                    var numbers = Numbers(clause).ToArray();
                    return numbers.Length > 0 && Near(numbers[0], Value(metric, "expected"), .01);
                });
            });
        }

        private static bool ChartMatches(StressChart chart, Dictionary<string, object> rule)
        {
            if (!string.IsNullOrEmpty(chart.error)) return false;
            var categories = Strings(Value(rule, "category_labels"));
            var series = Items(Value(rule, "series_values")).ToArray();
            if (categories.Length == 0 || series.Length == 0) return false;
            var expected = series[0] is IEnumerable && !(series[0] is string) ? series.Select(v => Items(v).ToArray()).ToArray() : new[] { series };
            if (expected.Length != chart.series.Length) return false;
            for (int i = 0; i < expected.Length; i++)
            {
                var actual = chart.series[i];
                if (!categories.SequenceEqual(actual.categories) || actual.values.Length != expected[i].Length) return false;
                for (int j = 0; j < actual.values.Length; j++) if (!Near(actual.values[j], expected[i][j], Number(rule, "tolerance", .01))) return false;
            }
            return (!Flag(rule, "zero_baseline") || chart.minimum.HasValue && Math.Abs(chart.minimum.Value) <= .001) &&
                (string.IsNullOrEmpty(Text(rule, "units")) || Contains(chart.title, Text(rule, "units")));
        }
        private static bool PresentationMatches(StressNative native, Dictionary<string, object> rule, LabRun run, string text)
        {
            if (native.slides.Length != (int)Number(rule, "slide_count", -1)) return false;
            if (!Strings(Value(rule, "required_facts")).All(v => Contains(text, v))) return false;
            var themePath = Text(rule, "theme_ref");
            if (string.IsNullOrEmpty(themePath) || !themePath.StartsWith("evaluator-only/", StringComparison.Ordinal)) return false;
            var manifest = TestLabSuite.Read<KitManifest>(Path.Combine(run.fixture_root, "manifest.json"));
            if (!manifest.files.Any(f => f.path == themePath)) return false;
            var theme = new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(File.ReadAllText(TestLab.SafeChild(run.fixture_root, themePath)));
            if (!Near(native.width, Value(theme, "width"), 1) || !Near(native.height, Value(theme, "height"), 1)) return false;
            var fonts = Strings(Value(theme, "fonts")); var palette = Strings(Value(theme, "palette"));
            if (fonts.Length == 0 || palette.Length == 0) return false;
            foreach (var slide in native.slides)
            {
                var textShapes = slide.shapes.Where(s => !string.IsNullOrWhiteSpace(s.text)).ToArray();
                if (textShapes.Length == 0) return false;
                foreach (var shape in slide.shapes)
                {
                    if (shape.x < -1 || shape.y < -1 || shape.x + shape.width > native.width + 1 || shape.y + shape.height > native.height + 1) return false;
                    if (!string.IsNullOrEmpty(shape.fill_color) && !palette.Any(c => string.Equals(c.TrimStart('#'), shape.fill_color.TrimStart('#'), StringComparison.OrdinalIgnoreCase))) return false;
                    if (string.IsNullOrWhiteSpace(shape.text)) continue;
                    // Samsung's source/footer/folio bands deliberately use
                    // compact labels; ordinary body copy and native table cells
                    // have different documented minimum sizes.
                    var footer = shape.y >= native.height * .925 && shape.y + shape.height <= native.height + 1;
                    var unitLabel = shape.x >= native.width * .79 && shape.width <= native.width * .18 &&
                        shape.y >= native.height * .20 && shape.y + shape.height <= native.height * .26;
                    var minimumFont = shape.is_table_cell ? 7.5 : footer ? 7 : unitLabel ? 8 : 14;
                    if (shape.bound_width > shape.available_width + 2 || shape.bound_height > shape.available_height + 2 ||
                        shape.font_size < minimumFont - .01 ||
                        !fonts.Contains(shape.font, StringComparer.OrdinalIgnoreCase) || !palette.Any(c => string.Equals(c.TrimStart('#'), (shape.color ?? "").TrimStart('#'), StringComparison.OrdinalIgnoreCase))) return false;
                }
                foreach (var series in slide.charts.SelectMany(c => c.series))
                    foreach (var color in new[] { series.fill_color, series.line_color }.Where(c => !string.IsNullOrEmpty(c)))
                        if (!palette.Any(c => string.Equals(c.TrimStart('#'), color.TrimStart('#'), StringComparison.OrdinalIgnoreCase))) return false;
                for (int a = 0; a < textShapes.Length; a++) for (int b = a + 1; b < textShapes.Length; b++)
                {
                    var left = textShapes[a]; var right = textShapes[b];
                    if (Math.Min(left.x + left.width, right.x + right.width) - Math.Max(left.x, right.x) > 2 &&
                        Math.Min(left.y + left.height, right.y + right.height) - Math.Max(left.y, right.y) > 2) return false;
                }
            }
            return true;
        }
        private static bool Headers(string value, string[] expected)
        {
            var actual = Regex.Matches(value ?? "", @"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}").Cast<Match>().Select(m => m.Value.ToLowerInvariant()).Distinct().OrderBy(x => x).ToArray();
            return actual.SequenceEqual(expected.Select(x => x.ToLowerInvariant()).Distinct().OrderBy(x => x));
        }
        private static bool Contains(string text, string wanted)
        {
            double expected;
            if (!double.TryParse(wanted, NumberStyles.Float, CultureInfo.InvariantCulture, out expected))
                return !string.IsNullOrWhiteSpace(wanted) && (text ?? "").IndexOf(wanted, StringComparison.OrdinalIgnoreCase) >= 0;
            return Numbers(text).Any(number => Near(number, expected, .01));
        }
        private static IEnumerable<double> Numbers(string text)
        {
            foreach (Match match in Regex.Matches(text ?? "",
                @"(?<![A-Za-z0-9_.,])(?:(?:EUR|USD|AED|GBP)\s*)?(?<number>[-+\u2212]?(?:(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d*)?|\.\d+)(?:[eE][-+]?\d+)?)(?![A-Za-z0-9_]|[.,]\d)",
                RegexOptions.IgnoreCase))
            {
                double number;
                if (double.TryParse(match.Groups["number"].Value.Replace(",", "").Replace('\u2212', '-'), NumberStyles.Float,
                    CultureInfo.InvariantCulture, out number) && !double.IsNaN(number) && !double.IsInfinity(number)) yield return number;
            }
        }
        private static bool CompletedMailboxSearch(string timeline)
        {
            var states = new Dictionary<string, bool>();
            var json = new JavaScriptSerializer { MaxJsonLength = 2 * 1024 * 1024 };
            foreach (var line in (timeline ?? "").Split('\n').Where(l => !string.IsNullOrWhiteSpace(l)))
            {
                var entry = json.Deserialize<Dictionary<string, object>>(line);
                if (Text(entry, "stage") != "tool_result") continue;
                var detail = Value(entry, "detail") as Dictionary<string, object>;
                if (detail == null || Text(detail, "name") != "search_mailbox") continue;
                var response = json.Deserialize<Dictionary<string, object>>(Text(detail, "Content"));
                var progress = Value(response, "progress") as Dictionary<string, object>;
                if (progress == null || string.IsNullOrEmpty(Text(progress, "cursor_id"))) continue;
                states[Text(progress, "cursor_id")] = Flag(response, "enumeration_complete") && !Flag(response, "truncated");
            }
            return states.Count > 0 && states.Values.All(v => v);
        }
        private static object Value(Dictionary<string, object> data, string key) { object value; return data.TryGetValue(key, out value) ? value : null; }
        private static string Text(Dictionary<string, object> data, string key) { return Convert.ToString(Value(data, key), CultureInfo.InvariantCulture); }
        private static bool Flag(Dictionary<string, object> data, string key) { return Value(data, key) is bool && (bool)Value(data, key); }
        private static double Number(Dictionary<string, object> data, string key, double fallback) { double number; return double.TryParse(Text(data, key), NumberStyles.Float, CultureInfo.InvariantCulture, out number) ? number : fallback; }
        private static bool Near(object value, object expected, double tolerance)
        {
            double actual, target;
            return tolerance >= 0 && tolerance <= 1 && double.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), NumberStyles.Float, CultureInfo.InvariantCulture, out actual) &&
                double.TryParse(Convert.ToString(expected, CultureInfo.InvariantCulture), NumberStyles.Float, CultureInfo.InvariantCulture, out target) &&
                !double.IsNaN(actual) && !double.IsInfinity(actual) && Math.Abs(actual - target) <= tolerance;
        }
        private static IEnumerable<object> Items(object value) { return value == null ? new object[0] : (value as IEnumerable)?.Cast<object>() ?? new[] { value }; }
        private static string[] Strings(object value) { return value is string ? new[] { (string)value } : Items(value).Select(v => Convert.ToString(v, CultureInfo.InvariantCulture)).ToArray(); }
    }
    public sealed class StressReadback
    {
        public string run_id, artifact_extension, text, to, cc, bcc, subject, body;
        public bool native_readback, run_created_output, output_boundary, unsent;
        public StressNative stress_native;
    }
}
