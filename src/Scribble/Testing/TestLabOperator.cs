using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Scribble.Testing
{
    public static class TestLabStressSettings
    {
        // A disposable in-memory view. Never rewrite the user's saved settings
        // or expose their personal repositories, writing profile, or MCP tools.
        public static Scribble.Configuration.AppSettings Isolate(Scribble.Configuration.AppSettings saved)
        {
            if (saved == null) throw new ArgumentNullException(nameof(saved));
            var isolated = saved.ForModel(saved.Model);
            isolated.UseToneProfile = false; isolated.ToneProfile = ""; isolated.DraftRules = "";
            isolated.Topics = new System.Collections.Generic.List<Scribble.Configuration.TopicConfig>();
            isolated.McpServers = new System.Collections.Generic.List<Scribble.Configuration.McpServerConfig>();
            isolated.DiscoveredModels = new System.Collections.Generic.List<string> { saved.Model };
            isolated.SwitchToVisionModelForImages = false;
            isolated.UseGeminiSignIn = false; isolated.GeminiRefreshToken = ""; isolated.GeminiProject = "";
            return isolated;
        }
    }

    public static class TestLabOperatorReporting
    {
        public sealed class Manifest
        {
            public int schema { get; set; }
            public SuiteState suite { get; set; }
            public SuiteCaseResult[] cases { get; set; }
        }

        // Re-render captured results in this executable's own runtime. This path
        // never enables Test Lab capture, opens Office, or contacts a model.
        public static string Create(string manifestPath)
        {
            if (string.IsNullOrWhiteSpace(manifestPath) || !Path.IsPathRooted(manifestPath) ||
                (Path.GetPathRoot(manifestPath) ?? "").Length < 3 ||
                !string.Equals(Path.GetFileName(manifestPath), "suite.json", StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("Supply an absolute path to the captured suite.json file.");
            manifestPath = Path.GetFullPath(manifestPath);
            var folder = Path.GetDirectoryName(manifestPath);
            var manifest = TestLabSuite.Read<Manifest>(manifestPath);
            if (manifest == null || manifest.schema != 1 || manifest.suite == null || manifest.cases == null ||
                !string.Equals(Path.GetFullPath(manifest.suite.folder ?? "" ).TrimEnd(Path.DirectorySeparatorChar), folder, StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException("The suite manifest must describe its own containing results folder.");
            foreach (var result in manifest.cases)
                foreach (var path in new[] { result.evidence, result.evaluation }.Where(p => !string.IsNullOrEmpty(p)))
                    if (!Path.GetFullPath(path).StartsWith(folder + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
                        throw new InvalidDataException("Case evidence must remain inside the captured suite folder.");
            var pdf = TestLabSuiteReport.Create(manifest.suite, manifest.cases);
            if (!TestLabPdfWriter.IsValid(pdf)) throw new InvalidDataException("The standalone reporter did not produce a valid PDF.");
            return pdf;
        }
    }

    public static class TestLabOperatorExecution
    {
        // A failed diagnostic write must not replace the runner's terminal
        // receipt or skip mandatory report finalization. Never complete while
        // the actual runner task is still active.
        public static async Task<string> RunToTerminalAsync(Func<Task> start, Action<Exception> reportFailure)
        {
            if (start == null) throw new ArgumentNullException(nameof(start));
            try { await start(); return null; }
            catch (Exception error)
            {
                var failure = error.ToString();
                try { reportFailure?.Invoke(error); }
                catch (Exception loggingError) { failure += "\nDiagnostic logging also failed: " + loggingError; }
                return failure;
            }
        }
    }

    // Explicit operator entry point. These options are never inferred from a
    // browser native-message request or from merely opening the Test Lab window.
    public sealed class TestLabOperatorOptions
    {
        public string ResultPath { get; private set; }
        public string CaseId { get; private set; }
        public string KitPath { get; private set; }
        public string KitSha256 { get; private set; }
        public string[] RequestedCaseIds { get; private set; }
        public int RequestedCount => RequestedCaseIds?.Length ?? OfficeCaseIds.Length;
        public string StopPath => ResultPath + ".stop";
        internal string StopToken { get; } = Guid.NewGuid().ToString("N");
        public static string[] OfficeCaseIds => new[] {
            "EX01", "EX02", "EX03", "EX04", "PP01", "PP02", "PP03", "OL01",
            "XA01", "XA02", "XA04", "RB01", "RC01", "EX05", "PP04", "OL02" };

        public static TestLabOperatorOptions Parse(string[] args)
        {
            if (args == null || args.Length == 0 || args[0] != "--test-lab-run")
                throw new ArgumentException("Use --test-lab-run --result-json <new absolute path> [--case EX01].");
            string result = null, caseId = null, caseList = null, kit = null, kitHash = null;
            for (var i = 1; i < args.Length; i += 2)
            {
                if (i + 1 >= args.Length || string.IsNullOrWhiteSpace(args[i + 1])) throw new ArgumentException("Missing value for " + args[i] + ".");
                if (args[i] == "--result-json" && result == null) result = args[i + 1];
                else if (args[i] == "--case" && caseId == null) caseId = args[i + 1].ToUpperInvariant();
                else if (args[i] == "--cases" && caseList == null) caseList = args[i + 1].ToUpperInvariant();
                else if (args[i] == "--kit" && kit == null) kit = args[i + 1];
                else if (args[i] == "--kit-sha256" && kitHash == null) kitHash = args[i + 1].ToLowerInvariant();
                else throw new ArgumentException("Unknown or repeated operator option: " + args[i]);
            }
            if (string.IsNullOrWhiteSpace(result) || !Path.IsPathRooted(result) || (Path.GetPathRoot(result) ?? "").Length < 3 ||
                !string.Equals(Path.GetExtension(result), ".json", StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("--result-json must name a new absolute .json file.");
            if (caseId != null && caseList != null) throw new ArgumentException("Use either --case or --cases.");
            if ((kit == null) != (kitHash == null)) throw new ArgumentException("External kits require both --kit and --kit-sha256.");
            if (kit != null && (!Path.IsPathRooted(kit) || (Path.GetPathRoot(kit) ?? "").Length < 3 ||
                !System.Text.RegularExpressions.Regex.IsMatch(kitHash, "^[a-f0-9]{64}$")))
                throw new ArgumentException("--kit must be an absolute directory and --kit-sha256 must identify its manifest.json.");
            var requested = (caseList ?? caseId)?.Split(',').Select(id => id.Trim()).ToArray();
            if (requested != null && (requested.Length == 0 || requested.Distinct(StringComparer.OrdinalIgnoreCase).Count() != requested.Length ||
                requested.Any(id => !System.Text.RegularExpressions.Regex.IsMatch(id, "^[A-Z]{2}[0-9]{2}$"))))
                throw new ArgumentException("Case IDs must be unique comma-separated IDs such as EX01,OL02.");
            if (kit == null && requested != null && requested.Any(id => !OfficeCaseIds.Contains(id)))
                throw new ArgumentException("--case must identify one of the 16 Office cases: " + string.Join(", ", OfficeCaseIds));
            return new TestLabOperatorOptions { ResultPath = Path.GetFullPath(result), CaseId = caseId,
                RequestedCaseIds = requested, KitPath = kit == null ? null : Path.GetFullPath(kit), KitSha256 = kitHash };
        }

        internal void ValidateKitSelection()
        {
            if (KitPath == null) return;
            if (!Directory.Exists(KitPath) || !string.Equals(TestLab.FileHash(Path.Combine(KitPath, "manifest.json")), KitSha256, StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException("External kit manifest checksum mismatch.");
            TestLab.VerifyKit(KitPath);
            var selected = TestLabSuite.SelectRequestedCases(TestLabSuite.Read<LabCase[]>(Path.Combine(KitPath, "operator", "cases.json")), RequestedCaseIds);
            RequestedCaseIds = selected.Select(c => c.id).ToArray();
        }

        internal void ReserveResult()
        {
            Directory.CreateDirectory(Path.GetDirectoryName(ResultPath));
            if (File.Exists(StopPath)) throw new IOException("A stop file already exists at " + StopPath + "; choose a new result path.");
            using (new FileStream(ResultPath, FileMode.CreateNew, FileAccess.Write, FileShare.Read)) { }
        }
        internal void WriteResult(TestLabOperatorResult result)
        {
            result.stop_file = StopPath; result.stop_token = StopToken;
            var temporary = ResultPath + "." + Guid.NewGuid().ToString("N") + ".tmp";
            try
            {
                File.WriteAllText(temporary, TestLab.Serialize(result), new UTF8Encoding(false));
                for (var attempt = 0; ; attempt++)
                {
                    try { File.Replace(temporary, ResultPath, null); return; }
                    catch (IOException error) when (attempt < 39 && ((error.HResult & 0xffff) == 32 || (error.HResult & 0xffff) == 33))
                    { Thread.Sleep(25); }
                }
            }
            finally { if (File.Exists(temporary)) File.Delete(temporary); }
        }
    }

    public sealed class TestLabOperatorResult
    {
        public int schema = 1;
        public string execution_kind = "configured_model", status = "starting", correctness_status = "not_evaluated";
        public int exit_code = 2, requested_count, completed_count;
        public bool harness_complete, pdf_valid;
        public string case_id, suite_id, folder, pdf, configured_model, error, stop_file, stop_token;
        public SuiteCaseResult[] cases = new SuiteCaseResult[0];

        public static TestLabOperatorResult Classify(string caseId, SuiteState state, SuiteCaseResult[] results,
            string pdf, bool pdfValid, bool cancelled, string error)
        {
            var expected = string.IsNullOrEmpty(caseId) ? TestLabOperatorOptions.OfficeCaseIds : new[] { caseId };
            return ClassifyRequested(expected, caseId, state, results, pdf, pdfValid, cancelled, error);
        }

        public static TestLabOperatorResult ClassifyRequested(string[] expected, string caseId, SuiteState state, SuiteCaseResult[] results,
            string pdf, bool pdfValid, bool cancelled, string error)
        {
            expected = expected ?? new string[0];
            results = results ?? new SuiteCaseResult[0];
            var exactScope = expected.Length > 0 && expected.Distinct(StringComparer.OrdinalIgnoreCase).Count() == expected.Length &&
                expected.All(id => System.Text.RegularExpressions.Regex.IsMatch(id ?? "", "^[A-Z]{2}[0-9]{2}$")) && results.Length == expected.Length && results.Select(r => r.id).OrderBy(id => id, StringComparer.Ordinal)
                .SequenceEqual(expected.OrderBy(id => id, StringComparer.Ordinal));
            var complete = exactScope && pdfValid && string.IsNullOrEmpty(error) && !cancelled &&
                results.All(r => r.status == "needs_review" || r.status == "failed");
            var failed = results.Any(r => r.status == "failed" || r.evaluationStatus == "failed");
            return new TestLabOperatorResult {
                status = cancelled ? "stopped" : complete ? "completed" : "blocked",
                exit_code = cancelled ? 4 : !complete ? 2 : failed ? 3 : 0,
                correctness_status = failed ? "failed" : complete ? "needs_review" : "not_evaluated",
                requested_count = expected.Length,
                completed_count = results.Count(r => r.status == "needs_review" || r.status == "failed"),
                harness_complete = complete, pdf_valid = pdfValid, case_id = caseId,
                suite_id = state?.id, folder = state?.folder, pdf = pdf, error = error, cases = results
            };
        }
    }
}
