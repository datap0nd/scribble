using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading;
using System.Web.Script.Serialization;
using Scribble.Testing;

namespace GuardrailTests
{
    internal static class StressKitTests
    {
        private static void Check(bool value, string message) { if (!value) throw new Exception(message); }
        private static void Rejected(Action action, string message)
        {
            try { action(); } catch (ArgumentException) { return; } catch (InvalidDataException) { return; }
            throw new Exception(message);
        }
        private static void Write(string root, string path, object value)
        {
            var target = Path.Combine(root, path.Replace('/', Path.DirectorySeparatorChar));
            Directory.CreateDirectory(Path.GetDirectoryName(target));
            File.WriteAllText(target, new JavaScriptSerializer().Serialize(value));
        }

        public static void ExternalOptionsAndExactScope()
        {
            var path = Path.Combine(Path.GetTempPath(), "stress-result.json");
            var kit = Path.Combine(Path.GetTempPath(), "stress-kit");
            var hash = new string('a', 64);
            var args = new[] { "--test-lab-run", "--result-json", path, "--kit", kit, "--kit-sha256", hash };
            var options = TestLabOperatorOptions.Parse(args.Concat(new[] { "--cases", "ex01,ol70,pp50" }).ToArray());
            Check(options.RequestedCaseIds.SequenceEqual(new[] { "EX01", "OL70", "PP50" }) && options.KitPath == kit && options.KitSha256 == hash,
                "External operator selection changed IDs or kit identity.");
            Rejected(() => TestLabOperatorOptions.Parse(args.Take(5).ToArray()), "Unpinned external kit accepted.");
            Rejected(() => TestLabOperatorOptions.Parse(args.Concat(new[] { "--cases", "EX01,EX01" }).ToArray()), "Duplicate selected ID accepted.");
            Rejected(() => TestLabOperatorOptions.Parse(args.Concat(new[] { "--case", "EX01", "--cases", "EX02" }).ToArray()), "Ambiguous selection accepted.");
            Rejected(() => TestLabOperatorOptions.Parse(args.Concat(new[] { "--cases", "EX01," }).ToArray()), "Empty selected ID accepted.");
            var catalog = Enumerable.Range(1, 70).Select(n => new LabCase { id = "OL" + n.ToString("00"), host = "Outlook" }).ToArray();
            var selected = TestLabSuite.SelectRequestedCases(catalog, new[] { "OL70", "OL01" });
            Check(selected.Select(c => c.id).SequenceEqual(new[] { "OL70", "OL01" }), "Explicit execution order changed.");
            var full = Enumerable.Range(1, 60).Select(n => new LabCase { id = "EX" + n.ToString("00"), host = "Excel" })
                .Concat(Enumerable.Range(1, 70).Select(n => new LabCase { id = "OL" + n.ToString("00"), host = n == 70 ? "Chrome" : "Outlook" }))
                .Concat(Enumerable.Range(1, 50).Select(n => new LabCase { id = "PP" + n.ToString("00"), host = "PowerPoint" }))
                .Concat(Enumerable.Range(1, 20).Select(n => new LabCase { id = "XA" + n.ToString("00"), host = "Excel" })).ToArray();
            Check(TestLabSuite.SelectRequestedCases(full, null).Length == 200, "The full external stress scope omitted its Chrome case.");
            Check(TestLabSuite.SelectCases(full.Take(19).Concat(new[] { new LabCase { id = "CH01", host = "Chrome" } }).ToArray(), null).All(c => c.host != "Chrome"),
                "The legacy default suite unexpectedly selected Chrome.");
            Rejected(() => TestLabSuite.SelectRequestedCases(catalog, new[] { "OL71" }), "Absent case accepted.");
            Rejected(() => TestLabSuite.SelectRequestedCases(catalog.Concat(catalog.Take(1)).ToArray(), null), "Duplicate catalog accepted.");
            var requested = catalog.Select(c => c.id).ToArray();
            var results = catalog.Select(c => new SuiteCaseResult { id = c.id, status = "needs_review" }).ToArray();
            var result = TestLabOperatorResult.ClassifyRequested(requested, null, null, results, "report.pdf", true, false, null);
            Check(result.harness_complete && result.requested_count == 70 && result.completed_count == 70 && result.correctness_status == "needs_review",
                "External complete scope was measured against the bundled16 cases.");
            Check(!TestLabOperatorResult.ClassifyRequested(requested, null, null, results.Take(69).ToArray(), "report.pdf", true, false, null).harness_complete,
                "An incomplete external scope counted as complete.");
            var saved = new Scribble.Configuration.AppSettings { Model = "qwen/qwen3.8-27b", ApiKey = "fake-test-key",
                BaseUrl = "https://openrouter.ai/api/v1", UseToneProfile = true, ToneProfile = "private wording",
                DraftRules = "private rules", GeminiRefreshToken = "fake-private-token", UseGeminiSignIn = true };
            saved.Topics.Add(new Scribble.Configuration.TopicConfig());
            saved.McpServers.Add(new Scribble.Configuration.McpServerConfig());
            var isolated = TestLabStressSettings.Isolate(saved);
            Check(isolated.Model == saved.Model && isolated.ApiKey == saved.ApiKey && isolated.BaseUrl == saved.BaseUrl &&
                !isolated.UseToneProfile && isolated.ToneProfile == "" && isolated.DraftRules == "" && isolated.Topics.Count == 0 &&
                isolated.McpServers.Count == 0 && isolated.GeminiRefreshToken == "" && !isolated.SwitchToVisionModelForImages,
                "Stress pane settings retained personal context or changed the explicit inference identity.");
            Check(saved.UseToneProfile && saved.ToneProfile == "private wording" && saved.Topics.Count == 1 && saved.McpServers.Count == 1 &&
                saved.GeminiRefreshToken == "fake-private-token", "Stress isolation mutated the saved settings object.");
        }

        public static void ExternalSnapshotAndProjection()
        {
            // Pure file validation: no Office, model, session, or Test Lab state.
            var root = Path.Combine(Path.GetTempPath(), "scribble-stress-contract-" + Guid.NewGuid().ToString("N"));
            var source = Path.Combine(root, "source");
            Directory.CreateDirectory(source);
            try
            {
                var testCase = new LabCase { id = "EX60", host = "Excel", prompt = "Audit the selected source.",
                    inputs = new[] { "inputs/excel/WB20.xlsx" }, artifacts = new[] { "xlsx" }, oracle_ref = "evaluator-only/cases/EX60.json" };
                Write(source, "inputs/excel/WB20.xlsx", "source20");
                Write(source, "inputs/excel/WB19.xlsx", "source19");
                Write(source, "inputs/outlook/MAIL0001.eml", "mail body");
                Write(source, "operator/mail-index.json", new[] { new MailFixture { path = "inputs/outlook/MAIL0001.eml", attachments = new[] { "inputs/excel/WB19.xlsx" } } });
                Write(source, "operator/cases.json", new[] { testCase });
                Write(source, "evaluator-only/cases/EX60.json", new { id = "EX60", checks = new[] { new { kind = "numeric_cell", expected = 77 } } });
                Write(source, "evaluator-only/unrelated.json", new { secret = "do not project" });
                var files = Directory.GetFiles(source, "*", SearchOption.AllDirectories).Select(path => new KitFile {
                    path = path.Substring(source.Length + 1).Replace('\\', '/'), size = new FileInfo(path).Length, sha256 = TestLab.FileHash(path)
                }).ToArray();
                Write(source, "manifest.json", new KitManifest { schema = 1, suite_id = "scribble-stress-v1", files = files });
                var hash = TestLab.FileHash(Path.Combine(source, "manifest.json"));
                Rejected(() => TestLabSuite.SnapshotExternalKit(source, new string('b', 64), Path.Combine(root, "bad-snapshot"), CancellationToken.None),
                    "Incorrect external manifest pin accepted.");
                var snapshot = TestLabSuite.SnapshotExternalKit(source, hash, Path.Combine(root, "catalog"), CancellationToken.None);
                var projection = TestLabSuite.ProjectExternalCase(snapshot, Path.Combine(root, "EX60"), testCase, CancellationToken.None);
                Check(File.Exists(Path.Combine(projection, "inputs/excel/WB20.xlsx")) && File.Exists(Path.Combine(projection, testCase.oracle_ref)),
                    "Selected source or evaluator-only oracle missing.");
                Check(!File.Exists(Path.Combine(projection, "inputs/excel/WB19.xlsx")) && !File.Exists(Path.Combine(projection, "inputs/outlook/MAIL0001.eml")) &&
                    !File.Exists(Path.Combine(projection, "evaluator-only/unrelated.json")), "Projection cloned unrelated corpus inputs or answers.");
                Check(TestLab.VerifyKit(projection).parent_manifest_sha256 == hash, "Projection lost its pinned parent manifest identity.");
                File.AppendAllText(Path.Combine(source, "inputs/excel/WB20.xlsx"), "external change");
                TestLab.VerifyKit(projection);
                File.AppendAllText(Path.Combine(snapshot, "inputs/excel/WB20.xlsx"), "snapshot tamper");
                Rejected(() => TestLabSuite.ProjectExternalCase(snapshot, Path.Combine(root, "tampered"), testCase, CancellationToken.None),
                    "Tampered shared input was activated.");
                testCase.inputs = new[] { "evaluator-only/unrelated.json" };
                Rejected(() => TestLabSuite.ProjectExternalCase(snapshot, Path.Combine(root, "leak"), testCase, CancellationToken.None),
                    "Evaluator-only oracle accepted as a model-visible input.");
            }
            finally { if (Directory.Exists(root)) Directory.Delete(root, true); }
        }

        public static void PresentationMeasurementsDistinguishBodyAndTables()
        {
            var root = Path.Combine(Path.GetTempPath(), "scribble-stress-table-grade-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            try
            {
                var testCase = new LabCase { id = "PP01", host = "PowerPoint", prompt = "Create one editable synthetic review slide.",
                    inputs = new string[0], artifacts = new[] { "pptx" }, oracle_ref = "evaluator-only/cases/PP01.json" };
                Write(root, "operator/cases.json", new[] { testCase }); Write(root, "operator/mail-index.json", new MailFixture[0]);
                Write(root, "evaluator-only/theme.json", new { width = 960, height = 540, fonts = new[] { "Arial", "Arial Narrow" }, palette = new[] { "000000", "F2F2F2" } });
                Write(root, testCase.oracle_ref, new { id = "PP01", checks = new[] { new { kind = "presentation", slide_count = 1,
                    theme_ref = "evaluator-only/theme.json", required_facts = new[] { "46000" } } } });
                var files = Directory.GetFiles(root, "*", SearchOption.AllDirectories).Select(path => new KitFile {
                    path = path.Substring(root.Length + 1).Replace('\\', '/'), size = new FileInfo(path).Length, sha256 = TestLab.FileHash(path)
                }).ToArray();
                Write(root, "manifest.json", new KitManifest { schema = 1, suite_id = "scribble-stress-v1", files = files });
                var run = new LabRun { run_id = "measurement", case_id = "PP01", suite_id = "scribble-stress-v1", fixture_root = root,
                    manifest_sha256 = TestLab.FileHash(Path.Combine(root, "manifest.json")), input_paths = new string[0] };
                var chart = new StressChart { series = new[] { new StressSeries { name = "Revenue", fill_color = "#000000" } } };
                var slideCharts = new[] { chart };
                var outputText = "Profit EUR 46,000";
                Func<StressShape[], string[], TestLabCheck[]> evaluate = (shapes, warnings) => {
                    var capture = new StressReadback { run_id = run.run_id, native_readback = true, run_created_output = true, output_boundary = true,
                        artifact_extension = "pptx", text = outputText, stress_native = new StressNative { host = "PowerPoint", width = 960, height = 540,
                            measurement_warnings = warnings, slides = new[] { new StressSlide { number = 1, shapes = shapes, charts = slideCharts } } } };
                    using (var buffer = new MemoryStream())
                    {
                        using (var archive = new ZipArchive(buffer, ZipArchiveMode.Create, true))
                        using (var writer = new StreamWriter(archive.CreateEntry("artifacts/PowerPoint-final-output-1-readback.json").Open())) writer.Write(TestLab.Serialize(capture));
                        buffer.Position = 0;
                        using (var archive = new ZipArchive(buffer, ZipArchiveMode.Read)) return TestLabStressEvaluator.Evaluate(archive, run, testCase, new TestLabCheck[0]);
                    }
                };
                var body = new StressShape { name = "Body", text = "Profit EUR 46,000", font = "Arial", color = "#000000", font_size = 18,
                    x = 40, y = 140, width = 400, height = 60, available_width = 396, available_height = 58, bound_width = 200, bound_height = 22 };
                var cell = new StressShape { name = "Table R1C1", text = "Profit", font = "Arial Narrow", color = "#000000", font_size = 9,
                    x = 40, y = 240, width = 90, height = 30, available_width = 86, available_height = 28, bound_width = 30, bound_height = 12, is_table_cell = true };
                Check(evaluate(new[] { body, cell }, new string[0]).All(c => c.passed), "A readable9pt native table cell was treated as undersized body text.");
                slideCharts = new StressChart[0];
                body.text = outputText = "Revenue EUR 46,000; cost EUR 28,000; margin 39.13%.";
                Check(evaluate(new[] { body }, new string[0]).Any(c => c.hard && !c.passed),
                    "A numeric body-text dump with no visual hierarchy passed presentation grading.");
                var panels = new[] { 40d, 340d, 640d }.Select(x => new StressShape { name = "Card panel", fill_color = "#F2F2F2",
                    x = x, y = 215, width = 250, height = 195 }).ToArray();
                var cardLeads = new[] {
                    new StressShape { name = "Coverage lead", text = "144 source rows", font = "Arial", color = "#000000", font_size = 18,
                        x = 50, y = 275, width = 220, height = 45, available_width = 216, available_height = 43, bound_width = 170, bound_height = 24 },
                    new StressShape { name = "Integrity lead", text = "0 duplicate RowIDs", font = "Arial", color = "#000000", font_size = 18,
                        x = 350, y = 275, width = 220, height = 45, available_width = 216, available_height = 43, bound_width = 180, bound_height = 24 } };
                body.text = "Data quality 2026"; outputText = "Profit EUR 46,000";
                Check(evaluate(new[] { body }.Concat(panels).Concat(cardLeads).ToArray(), new string[0]).All(c => c.passed),
                    "Verified counts in distinct, readable cards were incorrectly forced into oversized KPI numerals.");
                cardLeads[1].text = "WB01-0033 / 50";
                Check(evaluate(new[] { body }.Concat(panels).Concat(cardLeads).ToArray(), new string[0]).Any(c => c.hard && !c.passed),
                    "An incidental source-row identifier satisfied the structured-card metric guardrail.");
                var explanatory = new[] {
                    new StressShape { name = "Scope heading", text = "Source scope", font = "Arial", color = "#000000", font_size = 18,
                        x = 50, y = 225, width = 220, height = 35, available_width = 216, available_height = 33, bound_width = 130, bound_height = 22 },
                    new StressShape { name = "Scope body", text = "January to June 2026", font = "Arial", color = "#000000", font_size = 16,
                        x = 50, y = 290, width = 220, height = 45, available_width = 216, available_height = 43, bound_width = 180, bound_height = 22 },
                    new StressShape { name = "Coverage heading", text = "Coverage status", font = "Arial", color = "#000000", font_size = 18,
                        x = 350, y = 225, width = 220, height = 35, available_width = 216, available_height = 33, bound_width = 150, bound_height = 22 },
                    new StressShape { name = "Coverage body", text = "Management view through May 2026", font = "Arial", color = "#000000", font_size = 16,
                        x = 350, y = 290, width = 220, height = 45, available_width = 216, available_height = 43, bound_width = 200, bound_height = 22 }
                };
                Check(evaluate(new[] { body }.Concat(panels).Concat(explanatory).ToArray(), new string[0]).All(c => c.passed),
                    "Distinct explanatory cards with readable headings and copy were forced to use oversized numeric KPIs.");
                var secondMetric = new StressShape { name = "Metric 2", text = "EUR 28,000 / 39.13%", font = "Arial", color = "#000000", font_size = 32,
                    x = 280, y = 140, width = 200, height = 60, available_width = 196, available_height = 58, bound_width = 150, bound_height = 36 };
                body.text = "EUR 46,000"; body.font_size = 32; body.width = 200; body.available_width = 196; body.bound_width = 150; body.bound_height = 36;
                Check(evaluate(new[] { body, secondMetric }, new string[0]).All(c => c.passed),
                    "Two prominent KPI values did not satisfy the visual-hierarchy guardrail.");
                body.text = outputText = "Profit EUR 46,000"; body.font_size = 18; body.width = 400; body.available_width = 396; body.bound_width = 200; body.bound_height = 22;
                slideCharts = new[] { chart };
                cell.fill_color = "#E91E63";
                Check(evaluate(new[] { body, cell }, new string[0]).Any(c => c.hard && !c.passed), "An incorrect pink table fill passed the Samsung palette check.");
                cell.fill_color = "#000000"; chart.series[0].fill_color = "#E91E63";
                Check(evaluate(new[] { body, cell }, new string[0]).Any(c => c.hard && !c.passed), "An incorrect pink chart series passed the Samsung palette check.");
                chart.series[0].fill_color = "#000000";
                cell.font_size = 7;
                Check(evaluate(new[] { body, cell }, new string[0]).Any(c => c.hard && !c.passed), "A table cell below Samsung7.5pt minimum passed.");
                cell.font_size = 9; cell.bound_height = 60;
                Check(evaluate(new[] { body, cell }, new string[0]).Any(c => c.hard && !c.passed), "Overflowing table-cell text passed.");
                cell.bound_height = 12;
                var pending = evaluate(new[] { body, cell }, new[] { "Group children require visual review." });
                Check(pending.Any(c => !c.hard && !c.passed) && pending.Where(c => c.hard).All(c => c.passed),
                    "Unknown group geometry was hidden or mislabeled as known bad output.");
                body.font_size = 8;
                Check(evaluate(new[] { body, cell }, new[] { "Group children require visual review." }).Any(c => c.hard && !c.passed),
                    "A review warning softened a known undersized-body failure.");
                body.font_size = 18;
                foreach (var text in new[] { "EUR46000", "USD46,000", "AED 46000.00", "GBP4.6e4" })
                {
                    outputText = text;
                    Check(evaluate(new[] { body, cell }, new string[0]).All(c => c.passed), "Valid currency notation was rejected: " + text);
                }
                foreach (var text in new[] { ".46000", "0.46000", "1.46000", "46000e-2", "12,46000", "ABC46000", "EUR460000" })
                {
                    outputText = text;
                    Check(evaluate(new[] { body, cell }, new string[0]).Any(c => c.hard && !c.passed), "A partial numeric token satisfied a whole-number oracle: " + text);
                }
            }
            finally { Directory.Delete(root, true); }
        }
        public static void NativeTableOracleKeepsOwnerDueDatePairs()
        {
            var method = typeof(TestLabStressEvaluator).GetMethod("RequiredNativeTables",
                System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic);
            Check(method != null, "Native table oracle is missing.");
            var serializer = new JavaScriptSerializer();
            var rule = serializer.Deserialize<Dictionary<string, object>>(serializer.Serialize(new {
                required_tables = new[] { new { slide = 1, headers = new[] { "Owner", "Due date" },
                    rows = new[] { new[] { "Mira Cole", "2026-07-10" }, new[] { "Leon Park", "2026-07-12" } } } }
            }));
            var cells = new[] { "Owner", "Due date", "Mira Cole", "2026-07-10", "Leon Park", "2026-07-12" }
                .Select((value, index) => new StressShape { name = "Table 1 R" + (index / 2 + 1) + "C" + (index % 2 + 1),
                    text = value, is_table_cell = true }).ToArray();
            var native = new StressNative { slides = new[] { new StressSlide { number = 1, shapes = cells } } };
            Func<bool> accepts = () => (bool)method.Invoke(null, new object[] { native, rule });
            Check(accepts(), "Exact owner and due-date pairs were rejected.");
            cells[3].text = "";
            Check(!accepts(), "A missing due-date column value passed native grading.");
            cells[3].text = "2026-07-12"; cells[5].text = "2026-07-10";
            Check(!accepts(), "Swapped owner and due-date pairs passed native grading.");
        }
    }
}
