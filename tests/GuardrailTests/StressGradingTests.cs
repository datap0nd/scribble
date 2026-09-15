using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Collections.Generic;
using Scribble.Testing;

namespace GuardrailTests
{
    internal static class StressGradingTests
    {
        private static void Check(bool condition, string message) { if (!condition) throw new Exception(message); }
        public static void BudgetRequiresFiniteTotal()
        {
            var actual = new Scribble.Configuration.AppSettings { BaseUrl = "https://openrouter.ai/api/v1", Model = "qwen/qwen3.8-27b", ApiKey = "fixture-key" };
            var expected = new Scribble.Configuration.AppSettings { BaseUrl = actual.BaseUrl, Model = actual.Model, ApiKey = actual.ApiKey };
            TestLabStressBudget.ValidateRequestSettings(actual, expected, actual.Model);
            expected.ApiKey = "new-capped-key";
            bool staleRejected = false;
            try { TestLabStressBudget.ValidateRequestSettings(actual, expected, actual.Model); } catch (InvalidOperationException) { staleRejected = true; }
            Check(staleRejected, "An existing pane could use its old uncapped key after the operator verified a different key.");
            Check(TestLabStressBudget.Validate("{\"data\":{\"limit\":10,\"limit_remaining\":10,\"usage\":0,\"limit_reset\":null}}").limit == 10,
                "The requested finite budget was rejected.");
            foreach (var response in new[] {
                "{\"data\":{\"limit\":null,\"limit_remaining\":null,\"usage\":0}}",
                "{\"data\":{\"limit\":100,\"limit_remaining\":100,\"usage\":0}}",
                "{\"data\":{\"limit\":10,\"limit_remaining\":10,\"usage\":0,\"limit_reset\":\"monthly\"}}",
                "{\"data\":{\"limit\":10,\"limit_remaining\":0.1,\"usage\":9.9}}",
                "{\"data\":{\"limit\":10,\"limit_remaining\":10,\"usage\":0,\"is_management_key\":true}}" })
            {
                bool failed = false; try { TestLabStressBudget.Validate(response); } catch (InvalidOperationException) { failed = true; }
                Check(failed, "An unbounded, resetting, exhausted or management key passed the stress budget.");
            }
            var folder = Path.Combine(Path.GetTempPath(), "scribble-provider-stop-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(folder);
            try
            {
                var state = new SuiteState { id = "persisted-provider-stop", fixtureSuiteId = "scribble-stress-v1", folder = folder, caseId = "EX01" };
                var path = Path.Combine(folder, "provider-stop.json");
                TestLabStressBudget.RecordProviderResponse(state, 429);
                Check(!File.Exists(path), "A transient provider response permanently stopped the stress attempt.");
                TestLabStressBudget.RecordProviderResponse(state, 402);
                var original = File.ReadAllText(path);
                TestLabStressBudget.RecordProviderResponse(state, 401);
                Check(File.ReadAllText(path) == original, "A later response replaced the first provider rejection.");
                // Reconstruct the state, as another Office process or the suite
                // runner would. The gate must fail before settings or HTTP work.
                for (var attempt = 0; attempt < 2; attempt++)
                {
                    var restored = new SuiteState { id = state.id, fixtureSuiteId = state.fixtureSuiteId, folder = folder, caseId = "EX02" };
                    bool stopped = false;
                    try { TestLabStressBudget.CheckAsync(restored, System.Threading.CancellationToken.None).GetAwaiter().GetResult(); }
                    catch (InvalidOperationException error) { stopped = error.Message.Contains("HTTP 402") && error.Message.Contains("EX01"); }
                    Check(stopped, "A persisted HTTP402 allowed another budget check to proceed toward inference.");
                }
            }
            finally { Directory.Delete(folder, true); }
        }
        public static void NativeCellsRequireRecalculation()
        {
            var oracle = new { id = "EX01", checks = new object[] { new { kind = "numeric_cell", sheet = "Scribble Draft*", cell = "B4", expected = 46000, tolerance = .01, formula_required = true } } };
            Func<double, bool, bool, bool> evaluate = (number, recalculated, outputBoundary) =>
            {
                var capture = new StressReadback { run_id = "test", native_readback = true, run_created_output = true, output_boundary = outputBoundary,
                    artifact_extension = "xlsx", text = "Profit 46,000. Source inputs already contain the expected answer.",
                    stress_native = new StressNative { host = "Excel", sheets = new[] { new StressSheet { name = "Scribble Draft", recalculated = recalculated,
                        cells = new[] { new StressCell { row = 4, column = 2, value = number, formula = "=SUM(Ledger!I2:I9)" } } } } } };
                return Evaluate(oracle, new[] { capture }, "", new[] { "xlsx" }).All(c => c.passed);
            };
            Check(evaluate(46000, true, true), "A correct recalculated native formula was rejected.");
            Check(!evaluate(48000, true, true), "Wrong output passed using the correct number elsewhere.");
            Check(!evaluate(46000, false, true), "A stale cached value passed without recalculation.");
            Check(!evaluate(46000, true, false), "A source-only readback satisfied an output check.");
        }
        public static void MailRequiresCompleteSearchAndExactIds()
        {
            var oracle = new { id = "OL01", checks = new object[] { new { kind = "mail_search", expected_ids = new[] { "MAIL0001" }, expected_count = 1 } } };
            Func<bool, string, bool> evaluate = (complete, answer) => {
                var tool = new { stage = "tool_result", detail = new { name = "search_mailbox", Content = TestLab.Serialize(new {
                    progress = new { cursor_id = "cursor1" }, enumeration_complete = complete, truncated = !complete }) } };
                var final = new { stage = "pane_event", detail = new { type = "assistant", text = answer } };
                return Evaluate(oracle, new StressReadback[0], TestLab.Serialize(tool) + "\n" + TestLab.Serialize(final), new string[0]).All(c => c.passed);
            };
            Check(evaluate(true, "MAIL0001\nTotal matches: 1"), "Complete native search with exact IDs was rejected.");
            Check(!evaluate(false, "MAIL0001\nTotal matches: 1"), "A partial search passed.");
            Check(!evaluate(true, "MAIL0001 MAIL0002\nTotal matches: 2"), "A false positive passed.");
            Check(!evaluate(true, "MAIL0001\nTotal matches: 10"), "An incorrect stated total passed.");
            foreach (var invalid in new[] { "1,000", "1.5", "1\nTotal matches: 5" })
                Check(!evaluate(true, "MAIL0001\nTotal matches: " + invalid), "A partial or contradictory total passed: " + invalid);
        }
        public static void IncompleteCellsAndDraftMetricsRemainGrounded()
        {
            var cellOracle = new { id = "EX01", checks = new object[] { new { kind = "cell_text", sheet = "Scribble Draft*", cell = "F4", expected = "incomplete" } } };
            var cell = new StressCell { row = 4, column = 6, value = "incomplete", formula = "=IF(COUNT(Source!B2:B4)<3,\"incomplete\",SUM(Source!B2:B4))" };
            var workbook = new StressReadback { run_id = "test", native_readback = true, run_created_output = true, output_boundary = true, artifact_extension = "xlsx",
                stress_native = new StressNative { host = "Excel", sheets = new[] { new StressSheet { name = "Scribble Draft", recalculated = true, cells = new[] { cell } } } } };
            Check(Evaluate(cellOracle, new[] { workbook }, "", new[] { "xlsx" }).All(c => c.passed), "Explicit incomplete rate was rejected.");
            cell.value = .42;
            Check(Evaluate(cellOracle, new[] { workbook }, "", new[] { "xlsx" }).Any(c => !c.passed), "An invented rate passed an incomplete-cell requirement.");
            var mailOracle = new { id = "EX01", checks = new object[] { new { kind = "draft_mail", to = new[] { "review@example.test" }, cc = new string[0], bcc = new string[0],
                body_metrics = new[] { new { label = "Revenue", expected = 120000 }, new { label = "Cost", expected = 74000 } } } } };
            var mail = new StressReadback { run_id = "test", native_readback = true, run_created_output = true, output_boundary = true, artifact_extension = "msg",
                unsent = true, to = "review@example.test", cc = "", bcc = "", body = "Revenue: EUR 120,000; Cost: EUR 74,000" };
            Check(Evaluate(mailOracle, new[] { mail }, "", new[] { "msg" }).All(c => c.passed), "Correct labeled native mail metrics were rejected.");
            mail.body = "Revenue: EUR 74,000; Cost: EUR 120,000";
            Check(Evaluate(mailOracle, new[] { mail }, "", new[] { "msg" }).Any(c => !c.passed), "Swapped metric labels passed merely because both numbers were present.");
            foreach (var invalid in new[] {
                "Revenue: EUR 74,000 (previously EUR 120,000); Cost: EUR 120,000 (previously EUR 74,000)",
                "Revenue: EUR 120,000; Cost: EUR 74,000; Revenue: EUR 74,000" })
            {
                mail.body = invalid;
                Check(Evaluate(mailOracle, new[] { mail }, "", new[] { "msg" }).Any(c => !c.passed), "An incorrect first amount or contradictory labeled amount passed.");
            }
        }
        public static void ReportRetainsTwoHundredResults()
        {
            var root = Path.Combine(Path.GetTempPath(), "scribble-stress-report-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(root);
            try
            {
                var ids = Enumerable.Range(1, 60).Select(n => "EX" + n.ToString("00")).Concat(Enumerable.Range(1, 70).Select(n => "OL" + n.ToString("00")))
                    .Concat(Enumerable.Range(1, 50).Select(n => "PP" + n.ToString("00"))).Concat(Enumerable.Range(1, 20).Select(n => "XA" + n.ToString("00"))).ToArray();
                var results = ids.Select(id => new SuiteCaseResult { id = id, host = "Synthetic fixture", status = "not_run", error = "No model request was submitted." }).ToArray();
                var state = new SuiteState { id = "stress-report-fixture", folder = root, commit = "packaging fixture", kitHash = "fixture" };
                var path = TestLabSuiteReport.Create(state, results);
                var html = File.ReadAllText(Path.Combine(root, "report.html"));
                Check(TestLabPdfWriter.IsValid(path) && ids.All(html.Contains), "The 200-case report omitted case results or failed PDF validation.");
                Check(html.Contains("No model request was submitted."), "Report hid the lack of model execution.");
            }
            finally { Directory.Delete(root, true); }
        }
        private static TestLabCheck[] Evaluate(object oracle, StressReadback[] captures, string timeline, string[] artifacts)
        {
            var root = Path.Combine(Path.GetTempPath(), "scribble-stress-grader-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(root);
            try
            {
                var id = captures.Length > 0 ? "EX01" : "OL01";
                var testCase = new LabCase { id = id, host = captures.Length > 0 ? "Excel" : "Outlook", prompt = "Use synthetic inputs.",
                    inputs = new string[0], artifacts = artifacts, oracle_ref = "evaluator-only/cases/" + id + ".json" };
                Action<string, object> write = (path, value) => { var file = Path.Combine(root, path); Directory.CreateDirectory(Path.GetDirectoryName(file)); File.WriteAllText(file, TestLab.Serialize(value)); };
                write(testCase.oracle_ref, oracle); write("operator/cases.json", new[] { testCase }); write("operator/mail-index.json", new MailFixture[0]);
                var files = Directory.GetFiles(root, "*", SearchOption.AllDirectories).Select(p => new KitFile { path = p.Substring(root.Length + 1).Replace('\\', '/'), size = new FileInfo(p).Length, sha256 = TestLab.FileHash(p) }).ToArray();
                write("manifest.json", new KitManifest { schema = 1, suite_id = "scribble-stress-v1", files = files });
                var run = new LabRun { run_id = "test", case_id = id, suite_id = "scribble-stress-v1", fixture_root = root,
                    manifest_sha256 = TestLab.FileHash(Path.Combine(root, "manifest.json")), input_paths = new string[0] };
                using (var buffer = new MemoryStream())
                {
                    using (var output = new ZipArchive(buffer, ZipArchiveMode.Create, true))
                    {
                        for (int i = 0; i < captures.Length; i++) using (var writer = new StreamWriter(output.CreateEntry("artifacts/Excel-final-output-" + i + "-readback.json").Open())) writer.Write(TestLab.Serialize(captures[i]));
                        using (var writer = new StreamWriter(output.CreateEntry("timeline.jsonl").Open())) writer.Write(timeline);
                    }
                    buffer.Position = 0;
                    using (var input = new ZipArchive(buffer, ZipArchiveMode.Read)) return TestLabStressEvaluator.Evaluate(input, run, testCase, new TestLabCheck[0]);
                }
            }
            finally { Directory.Delete(root, true); }
        }
    }
}
