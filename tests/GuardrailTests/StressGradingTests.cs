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
            Check(TestLabStressBudget.Validate("{\"data\":{\"limit\":20,\"limit_remaining\":10.5,\"usage\":9.5,\"limit_reset\":null}}").limit == 20,
                "The approved same-key checkpoint budget was rejected.");
            Check(TestLabStressBudget.Validate("{\"data\":{\"limit\":20,\"limit_remaining\":5.2,\"usage\":14.8,\"limit_reset\":null}}").limit == 20,
                "The final native-validation allowance was rejected.");
            Check(TestLabStressBudget.Validate("{\"data\":{\"limit\":20,\"limit_remaining\":1.7,\"usage\":18.3}}").limit == 20,
                "The reserved final native-validation allowance was rejected.");
            foreach (var response in new[] {
                "{\"data\":{\"limit\":null,\"limit_remaining\":null,\"usage\":0}}",
                "{\"data\":{\"limit\":100,\"limit_remaining\":100,\"usage\":0}}",
                "{\"data\":{\"limit\":10,\"limit_remaining\":10,\"usage\":0,\"limit_reset\":\"monthly\"}}",
                "{\"data\":{\"limit\":10,\"limit_remaining\":0.1,\"usage\":9.9}}",
                "{\"data\":{\"limit\":20,\"limit_remaining\":0.9,\"usage\":19.1}}",
                "{\"data\":{\"limit\":10,\"limit_remaining\":10,\"usage\":0,\"is_management_key\":true}}" })
            {
                bool failed = false; try { TestLabStressBudget.Validate(response); } catch (InvalidOperationException) { failed = true; }
                Check(failed, "An unbounded, resetting, exhausted, over-checkpoint or management key passed the stress budget.");
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
        public static void NativeChartsStayInRequestedHost()
        {
            var chartRule = new Dictionary<string, object> { ["kind"] = "native_chart", ["host"] = "PowerPoint", ["artifact_extension"] = "pptx",
                ["category_labels"] = new[] { "2026-05", "2026-06" }, ["series_values"] = new[] { 100000, 120000 }, ["zero_baseline"] = true, ["units"] = "EUR" };
            var oracle = new { id = "EX01", checks = new object[] { new { kind = "native_artifact", extension = "pptx" }, chartRule } };
            Func<StressChart> correctChart = () => new StressChart { title = "Revenue EUR", minimum = 0, series = new[] {
                new StressSeries { name = "Revenue", categories = new[] { "2026-05", "2026-06" }, values = new[] { "100000", "120000" } } } };
            var slide = new StressSlide { number = 1, charts = new[] { correctChart() } };
            var sheet = new StressSheet { name = "Scribble Draft", charts = new[] { correctChart() } };
            var workbook = new StressReadback { run_id = "test", native_readback = true, run_created_output = true, output_boundary = true,
                artifact_extension = "xlsx", stress_native = new StressNative { host = "Excel", sheets = new[] { sheet } } };
            var presentation = new StressReadback { run_id = "test", native_readback = true, run_created_output = true, output_boundary = true,
                artifact_extension = "pptx", stress_native = new StressNative { host = "PowerPoint", slides = new[] { slide } } };
            Func<TestLabCheck[]> evaluate = () => Evaluate(oracle, new[] { workbook, presentation }, "", new[] { "xlsx", "pptx" });
            Func<bool> passes = () => evaluate().All(c => c.passed);
            Check(passes(), "A valid PowerPoint chart with an additional matching Excel chart was rejected.");
            slide.charts = new StressChart[0];
            Check(!passes(), "A matching Excel chart rescued a chartless PowerPoint output.");
            slide.charts = new[] { correctChart() }; slide.charts[0].series[0].values[1] = "130000";
            Check(!passes(), "A matching Excel chart rescued incorrect PowerPoint chart values.");
            slide.charts = new[] { correctChart() }; presentation.output_boundary = false;
            Check(!passes(), "A source-only PowerPoint chart satisfied an output chart rule.");
            presentation.output_boundary = true; presentation.artifact_extension = "xlsx";
            Check(evaluate().Any(c => c.name.EndsWith("_native_chart", StringComparison.Ordinal) && !c.passed),
                "A chart capture with the wrong artifact type satisfied the PowerPoint rule.");
            presentation.artifact_extension = "pptx";
            chartRule["host"] = "Excel"; chartRule["artifact_extension"] = "xlsx";
            Check(passes(), "A valid Excel chart was rejected by its explicit host scope.");
            sheet.charts = new StressChart[0];
            Check(!passes(), "A PowerPoint chart rescued a chartless Excel output.");
            sheet.charts = new[] { correctChart() };
            foreach (var invalid in new[] { "", "Word", "PowerPoint" })
            {
                chartRule["host"] = invalid;
                Check(!passes(), "An absent or invalid native chart host/extension scope was accepted: " + invalid);
            }
            chartRule["host"] = "Excel"; chartRule.Remove("artifact_extension");
            Check(!passes(), "A native chart rule without an artifact extension was accepted.");
        }
        public static void CoverSourceLineIsFooterMetadata()
        {
            var theme = new { width = 960, height = 540, fonts = new[] { "Arial", "Arial Narrow" },
                palette = new[] { "FFFFFF", "000000", "4F81BD", "7F7F7F" } };
            var oracle = new { id = "EX01", checks = new object[] { new { kind = "presentation", slide_count = 1,
                theme_ref = "evaluator-only/theme.json", required_facts = new object[0] } } };
            var slide = new StressSlide { number = 1, shapes = new[] {
                new StressShape { name = "Title", text = "Atlas Components", font = "Arial", color = "#000000",
                    x = 44, y = 152, width = 810, height = 184, bound_width = 500, bound_height = 80,
                    available_width = 806, available_height = 182, font_size = 65 },
                new StressShape { name = "Source", text = "Source references and evidence: see speaker notes.", font = "Arial Narrow", color = "#000000",
                    x = 36, y = 490.32, width = 835, height = 16.2, bound_width = 140, bound_height = 8.4,
                    available_width = 831, available_height = 14.2, font_size = 7 } } };
            var capture = new StressReadback { run_id = "test", native_readback = true, run_created_output = true, output_boundary = true,
                artifact_extension = "pptx", text = "Atlas Components", stress_native = new StressNative { host = "PowerPoint", width = 960, height = 540,
                    slides = new[] { slide } } };
            Check(Evaluate(oracle, new[] { capture }, "", new[] { "pptx" }, theme).All(c => c.passed),
                "The Samsung cover source line was incorrectly graded as undersized body copy.");
            slide.shapes[1].y = 470;
            Check(Evaluate(oracle, new[] { capture }, "", new[] { "pptx" }, theme).Any(c => !c.passed),
                "Undersized body copy outside the footer band passed presentation grading.");
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
        public static void HeroCasesRequireExactNativeAndBrowserEvidence()
        {
            var workbookOracle = new { id = "EX01", checks = new object[] { new { kind = "workbook_exact_text", sheets = new[] {
                new { name = "Operations", cells = new[] { new { row = 1, column = 1, text = "Category" }, new { row = 2, column = 1, text = "Logistics" } } } } } } };
            var workbook = new StressReadback { run_id = "test", native_readback = true, run_created_output = true, output_boundary = true, artifact_extension = "xlsx",
                stress_native = new StressNative { host = "Excel", sheets = new[] { new StressSheet { name = "Operations", recalculated = true, cells = new[] {
                    new StressCell { row = 1, column = 1, value = "Category" }, new StressCell { row = 2, column = 1, value = "Logistics" } } } } } };
            Check(Evaluate(workbookOracle, new[] { workbook }, "", new[] { "xlsx" }).All(c => c.passed), "Exact translated workbook cells were rejected.");
            workbook.stress_native.sheets[0].cells[1].value = "물류";
            Check(Evaluate(workbookOracle, new[] { workbook }, "", new[] { "xlsx" }).Any(c => !c.passed), "Remaining Korean text passed the full-workbook translation gate.");

            var expectedTable = new[] { new[] { "Month", "Revenue" }, new[] { "2026-09", "420000" } };
            var wordOracle = new { id = "EX01", checks = new object[] { new { kind = "word_tables", tables = new[] { expectedTable } } } };
            var word = new StressReadback { run_id = "test", native_readback = true, run_created_output = true, output_boundary = true, artifact_extension = "docx",
                stress_native = new StressNative { host = "Word", tables = new[] { new StressTable { number = 1, rows = 2, columns = 2, borders = true, cells = new[] {
                    new StressWordCell { row = 1, column = 1, text = "Month", bold = true }, new StressWordCell { row = 1, column = 2, text = "Revenue", bold = true },
                    new StressWordCell { row = 2, column = 1, text = "2026-09" }, new StressWordCell { row = 2, column = 2, text = "420000" } } } } } };
            Check(Evaluate(wordOracle, new[] { word }, "", new[] { "docx" }).All(c => c.passed), "Exact bordered Word table was rejected.");
            word.stress_native.tables[0].cells[3].text = "42000";
            Check(Evaluate(wordOracle, new[] { word }, "", new[] { "docx" }).Any(c => !c.passed), "A dropped Word table digit passed exact transfer grading.");

            var browserOracle = new { id = "OL01", checks = new object[] { new { kind = "browser_evidence", purchasedProduct = "Galaxy Z Fold8",
                tradeInProduct = "Apple iPhone 16 Pro", storage = "256 GB", condition = "Flawless", market = "United Arab Emirates", currency = "AED",
                allowed_hosts = new[] { "www.samsungtradein.ae" } } } };
            var evidence = new { purchasedProduct = "Galaxy Z Fold8", tradeInProduct = "Apple iPhone 16 Pro", storage = "256 GB", condition = "Flawless",
                market = "United Arab Emirates", amount = "1,660", currency = "AED", caveat = "Estimate subject to inspection", sourceUrl = "https://www.samsungtradein.ae/ae-en/result" };
            var tool = new { stage = "tool_result", detail = new { name = "browser_record_evidence", Content = "[VERIFIED_BROWSER_EVIDENCE]\n" + TestLab.Serialize(evidence) } };
            var final = new { stage = "pane_event", detail = new { type = "assistant", text = "AED 1,660. Estimate subject to inspection. Source: www.samsungtradein.ae" } };
            Check(Evaluate(browserOracle, new StressReadback[0], TestLab.Serialize(tool) + "\n" + TestLab.Serialize(final), new string[0]).All(c => c.passed), "Verified live browser evidence was rejected.");
            var wrong = new { stage = "pane_event", detail = new { type = "assistant", text = "AED 1,760. Estimate subject to inspection. Source: www.samsungtradein.ae" } };
            Check(Evaluate(browserOracle, new StressReadback[0], TestLab.Serialize(tool) + "\n" + TestLab.Serialize(wrong), new string[0]).Any(c => !c.passed), "An answer that changed the verified amount passed.");

            var liveCase = new LabCase { host = "Chrome", browser_allowed_hosts = new[] { "www.samsungtradein.ae" },
                browser_start_url = "https://www.samsungtradein.ae/ae-en/" };
            Check(TestLab.IsBrowserSourceAllowed(liveCase, new Uri("https://www.samsungtradein.ae/ae-en/result?quote=1")), "The exact case allow-list rejected its HTTPS result page.");
            Check(TestLab.BrowserStartUri(liveCase).AbsoluteUri == liveCase.browser_start_url, "The verified live browser start URL changed.");
            Check(!TestLab.IsBrowserSourceAllowed(liveCase, new Uri("https://evil.example/")), "An unrelated live host escaped the case allow-list.");
            Check(!TestLab.IsBrowserSourceAllowed(liveCase, new Uri("https://user@www.samsungtradein.ae/")), "A credential-bearing URL escaped the case allow-list.");
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
        private static TestLabCheck[] Evaluate(object oracle, StressReadback[] captures, string timeline, string[] artifacts, object theme = null)
        {
            var root = Path.Combine(Path.GetTempPath(), "scribble-stress-grader-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(root);
            try
            {
                var id = captures.Length > 0 ? "EX01" : "OL01";
                var testCase = new LabCase { id = id, host = captures.Length > 0 ? "Excel" : "Outlook", prompt = "Use synthetic inputs.",
                    inputs = new string[0], artifacts = artifacts, oracle_ref = "evaluator-only/cases/" + id + ".json" };
                Action<string, object> write = (path, value) => { var file = Path.Combine(root, path); Directory.CreateDirectory(Path.GetDirectoryName(file)); File.WriteAllText(file, TestLab.Serialize(value)); };
                write(testCase.oracle_ref, oracle); write("operator/cases.json", new[] { testCase }); write("operator/mail-index.json", new MailFixture[0]);
                if (theme != null) write("evaluator-only/theme.json", theme);
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
