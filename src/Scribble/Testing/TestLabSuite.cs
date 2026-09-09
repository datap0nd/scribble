using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace Scribble.Testing
{
    public sealed class SuiteState
    {
        public string id, folder, commit, kitHash, caseId, host, runId, chromeToken, sourceUrl;
        public int pid;
        public long processStart;
        public DateTime expires;
    }
    public sealed class SuiteCaseResult
    {
        public string id, host, status = "not_run", error, started, finished, evidence;
    }
    public sealed class SuiteChromeCommand
    {
        public string id, action, prompt, sourceUrl, runId;
    }
    public static class TestLabSuite
    {
        private static string Descriptor => Path.Combine(TestLab.Root, "suite.bin");
        public static T Read<T>(string path) { return new JavaScriptSerializer { MaxJsonLength = int.MaxValue }.Deserialize<T>(File.ReadAllText(path)); }
        public static void Save(SuiteState state)
        {
            var temporary = Descriptor + "." + Guid.NewGuid().ToString("N");
            File.WriteAllBytes(temporary, ProtectedData.Protect(Encoding.UTF8.GetBytes(TestLab.Serialize(state)), null, DataProtectionScope.CurrentUser));
            if (File.Exists(Descriptor)) File.Replace(temporary, Descriptor, null); else File.Move(temporary, Descriptor);
        }
        public static SuiteState Active()
        {
            try {
                var state = new JavaScriptSerializer().Deserialize<SuiteState>(Encoding.UTF8.GetString(ProtectedData.Unprotect(File.ReadAllBytes(Descriptor), null, DataProtectionScope.CurrentUser)));
                using (var p = Process.GetProcessById(state.pid))
                    if (p.HasExited || p.StartTime.ToUniversalTime().Ticks != state.processStart || state.expires <= DateTime.UtcNow) return null;
                return state;
            } catch { return null; }
        }
        public static SuiteState Require(string id, string host)
        {
            var s = Active();
            if (s == null || s.id != id || s.host != host || s.runId == null || s.runId != TestLab.ActiveRunId())
                throw new InvalidOperationException("No matching active operator test suite.");
            return s;
        }
        public static LabCase CurrentCase(SuiteState s) { return TestLab.Cases().Single(c => c.id == s.caseId && c.host == s.host); }
        public static string Prompt(LabCase c, int phase)
        {
            if (phase == 0) return string.IsNullOrEmpty(c.prerequisite_prompt) ? c.prompt : c.prerequisite_prompt;
            if (phase == 1 && !string.IsNullOrEmpty(c.prerequisite_prompt)) return c.prompt;
            throw new InvalidOperationException("Invalid test phase.");
        }
        public static bool OwnsSource(string runId, string path)
        {
            var s = Active();
            if (s == null || s.runId != runId || string.IsNullOrEmpty(path)) return false;
            var run = TestLab.GetRun(runId);
            var root = TestLab.SafeChild(s.folder, "cases/" + s.caseId + "/scribble-test-kit-v1");
            return string.Equals(root, run.fixture_root, StringComparison.OrdinalIgnoreCase) &&
                (run.input_paths ?? new string[0]).Any(p => string.Equals(TestLab.SafeChild(root, p), path, StringComparison.OrdinalIgnoreCase));
        }
        // Native messaging uses a case-specific nonce and a single controller claim. Old tabs cannot submit a later case.
        public static string Chrome(string data)
        {
            var d = new JavaScriptSerializer().Deserialize<Dictionary<string, string>>(data ?? "{}");
            string id, token, controller, action;
            if (!d.TryGetValue("suite", out id) || !d.TryGetValue("token", out token) || !d.TryGetValue("controller", out controller) || !Regex.IsMatch(controller ?? "", "^[a-f0-9]{32}$")) throw new InvalidDataException("Invalid suite controller.");
            var s = Require(id, "Chrome");
            if (token != s.chromeToken) throw new InvalidOperationException("Expired Chrome test case.");
            var root = TestLab.SafeChild(s.folder, "cases/" + s.caseId);
            var claim = Path.Combine(root, "chrome-owner.txt");
            try { using (var w = new StreamWriter(new FileStream(claim, FileMode.CreateNew, FileAccess.Write, FileShare.Read))) w.Write(controller); }
            catch (IOException) { }
            if (File.ReadAllText(claim) != controller) return "null";
            d.TryGetValue("action", out action);
            var commandFile = Path.Combine(root, "chrome-command.json");
            if (!File.Exists(commandFile)) return "null";
            var command = Read<SuiteChromeCommand>(commandFile);
            if (action == "complete") {
                string commandId; d.TryGetValue("id", out commandId);
                if (commandId != command.id) return "null";
                string error; d.TryGetValue("error", out error);
                File.WriteAllText(Path.Combine(root, command.id + ".reply.json"), TestLab.Serialize(new SuiteReply { state = "done", error = error }));
                return "null";
            }
            return TestLab.Serialize(command);
        }
        public static string Extract(string zipPath, string destination)
        {
            Directory.CreateDirectory(destination);
            using (var zip = ZipFile.OpenRead(zipPath)) {
                long size = 0; var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                if (zip.Entries.Count > 5000) throw new InvalidDataException("Too many kit files.");
                foreach (var e in zip.Entries) {
                    size += e.Length;
                    if (size > 100 * 1024 * 1024 || !e.FullName.StartsWith("scribble-test-kit-v1/", StringComparison.Ordinal) || e.FullName.Contains(":") || e.FullName.Contains("\\")) throw new InvalidDataException("Invalid kit archive.");
                    var target = TestLab.SafeChild(destination, e.FullName);
                    if (!names.Add(target)) throw new InvalidDataException("Duplicate archive path.");
                    if (e.FullName.EndsWith("/", StringComparison.Ordinal)) { Directory.CreateDirectory(target); continue; }
                    Directory.CreateDirectory(Path.GetDirectoryName(target));
                    using (var source = e.Open()) using (var file = new FileStream(target, FileMode.CreateNew)) source.CopyTo(file);
                }
            }
            var root = Path.Combine(destination, "scribble-test-kit-v1"); TestLab.VerifyKit(root); return root;
        }
        private static byte[] Download(string url, int limit, CancellationToken cancel)
        {
            ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
            var request = (HttpWebRequest)WebRequest.Create(url); request.UserAgent = "Scribble-Test-Lab"; request.Timeout = 60000; request.ReadWriteTimeout = 30000;
            using (cancel.Register(request.Abort)) using (var response = request.GetResponse()) using (var input = response.GetResponseStream()) using (var bytes = new MemoryStream()) {
                var buffer = new byte[16384]; int n;
                while ((n = input.Read(buffer, 0, buffer.Length)) > 0) { cancel.ThrowIfCancellationRequested(); if (bytes.Length + n > limit) throw new InvalidDataException("Download exceeds kit size limit."); bytes.Write(buffer, 0, n); }
                return bytes.ToArray();
            }
        }
        public static string DownloadKit(SuiteState state, CancellationToken cancel)
        {
            var api = Encoding.UTF8.GetString(Download("https://api.github.com/repos/datap0nd/scribble/commits/main", 2 * 1024 * 1024, cancel));
            var commit = new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(api);
            state.commit = Convert.ToString(commit["sha"]);
            if (!Regex.IsMatch(state.commit, "^[a-f0-9]{40}$")) throw new InvalidDataException("Invalid main revision.");
            var root = "https://raw.githubusercontent.com/datap0nd/scribble/" + state.commit + "/tests/benchmarks/releases/scribble-test-kit-v1.zip";
            state.kitHash = Encoding.UTF8.GetString(Download(root + ".sha256", 4096, cancel)).Split(' ')[0].Trim();
            if (!Regex.IsMatch(state.kitHash, "^[a-f0-9]{64}$")) throw new InvalidDataException("Invalid kit checksum.");
            var bytes = Download(root, 25 * 1024 * 1024, cancel);
            if (TestLab.Hash(bytes) != state.kitHash) throw new InvalidDataException("Downloaded kit checksum mismatch.");
            var zip = Path.Combine(state.folder, "test-kit.zip"); File.WriteAllBytes(zip, bytes); return zip;
        }
    }

    internal sealed class TestLabSuiteRunner
    {
        internal readonly SuiteState State;
        internal readonly List<SuiteCaseResult> Results = new List<SuiteCaseResult>();
        private readonly Action<string> changed;
        private readonly CancellationToken cancel;
        private object application;
        private dynamic controller;
        private string preparation;
        private bool keepCaptureActive;
        internal TestLabSuiteRunner(string folder, Action<string> changed, CancellationToken cancel)
        {
            this.changed = changed; this.cancel = cancel;
            using (var p = Process.GetCurrentProcess()) State = new SuiteState { id = Guid.NewGuid().ToString("N"), folder = folder, pid = p.Id, processStart = p.StartTime.ToUniversalTime().Ticks, expires = DateTime.UtcNow.AddHours(8) };
        }
        internal void Log(string message)
        {
            var line = DateTime.UtcNow.ToString("O") + " " + message;
            File.AppendAllText(Path.Combine(State.folder, "suite.log"), line + Environment.NewLine); changed(line);
        }
        private void Journal() { File.WriteAllText(Path.Combine(State.folder, "suite.json"), TestLab.Serialize(new { schema = 1, suite = State, cases = Results })); }
        internal async Task Run()
        {
            Directory.CreateDirectory(TestLab.Root);
            using (var ownership = new FileStream(Path.Combine(TestLab.Root, "suite.lock"), FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None)) {
                if (TestLab.ActiveRunId() != null) throw new InvalidOperationException("Finish the existing manual test run before running the suite.");
                TestLabSuite.Save(State);
                try {
                    Log("Downloading the latest test kit from main...");
                    var zip = await Task.Run(() => TestLabSuite.DownloadKit(State, cancel));
                    var catalog = await Task.Run(() => TestLabSuite.Extract(zip, Path.Combine(State.folder, "catalog")));
                    var cases = TestLabSuite.Read<LabCase[]>(Path.Combine(catalog, "operator", "cases.json"));
                    Results.AddRange(cases.Select(c => new SuiteCaseResult { id = c.id, host = c.host }));
                    Log("Verified kit " + State.kitHash + " at main " + State.commit + "; " + cases.Length + " cases.");
                    foreach (var c in cases) {
                        cancel.ThrowIfCancellationRequested();
                        var result = Results.Single(r => r.id == c.id); result.started = DateTime.UtcNow.ToString("O"); result.status = "running";
                        State.caseId = c.id; State.host = c.host; State.runId = null; State.chromeToken = Guid.NewGuid().ToString("N"); State.sourceUrl = null;
                        TestLabSuite.Save(State); Journal();
                        var folder = TestLab.SafeChild(State.folder, "cases/" + c.id);
                        bool submitted = false, quiescent = true;
                        try {
                            Log(c.id + " / " + c.host + ": preparing visible apps and isolated files.");
                            var kit = await Task.Run(() => TestLabSuite.Extract(zip, folder)); cancel.ThrowIfCancellationRequested();
                            TestLab.Enable(kit);
                            await Prepare(c.id, folder);
                            State.runId = TestLab.Start(c.id, c.host, true).run_id; TestLabSuite.Save(State);
                            if (c.host == "Chrome") OpenChrome(); else ConnectOffice(c.host);
                            submitted = true; quiescent = false;
                            await Command("load", 0, 90, true);
                            if (c.id == "RC01") await CancellationCase(c);
                            else {
                                await Command("submit", 0, 600, true);
                                if (!string.IsNullOrEmpty(c.prerequisite_prompt)) {
                                    Log(c.id + ": prerequisite completed; capturing intermediate output.");
                                    Log(BenchmarkArtifactCollector.Capture(State.runId));
                                    if (c.id == "XA04") SaveSelectedDeck();
                                    await Command("submit", 1, 600, true);
                                }
                            }
                            quiescent = true;
                            result.status = "needs_review";
                            Log(c.id + ": model finished; saving native results and evidence (correctness needs review).");
                        } catch (Exception e) {
                            result.status = cancel.IsCancellationRequested ? "stopped" : "blocked"; result.error = e.ToString(); Log(c.id + " " + result.status.ToUpperInvariant() + ": " + e.Message);
                            if (submitted) {
                                try { await Command("stop", 0, 60, false); quiescent = true; }
                                catch (Exception stopError) { quiescent = false; Log("Could not confirm that the model stopped: " + stopError.Message); }
                            }
                        } finally {
                            if (submitted && !quiescent) keepCaptureActive = true;
                            TestLabPreparation.Stop(preparation); preparation = null;
                            if (State.runId != null) {
                                try { Log(BenchmarkArtifactCollector.Capture(State.runId)); } catch (Exception e) { Log("Output capture error: " + e.Message); result.error += "\nOutput capture: " + e; result.status = "blocked"; }
                                try {
                                    if (result.status != "needs_review") TestLab.MarkIncomplete(State.runId, result.error ?? result.status);
                                    if (quiescent) { TestLab.Finish(false); result.evidence = TestLab.Export(State.runId, folder); }
                                    else { keepCaptureActive = true; result.evidence = TestLab.ExportSnapshot(State.runId, folder); Log("Capture remains enabled because stopping the request was not confirmed. Stop it in the app before another suite."); }
                                    using (var evidence = ZipFile.OpenRead(result.evidence)) using (var reader = new StreamReader(evidence.GetEntry("run.json").Open())) {
                                        var exported = new JavaScriptSerializer().Deserialize<LabRun>(reader.ReadToEnd());
                                        if (result.status == "needs_review" && (exported.missing_artifacts.Length > 0 || !exported.trace_complete)) {
                                            result.status = "incomplete"; result.error = "Missing outputs: " + string.Join(", ", exported.missing_artifacts) + "; trace complete: " + exported.trace_complete;
                                        }
                                    }
                                    File.WriteAllText(Path.Combine(folder, "report.html"), TestLabReport.BuildHtml(result.evidence, Path.Combine(folder, "summary.txt")), new UTF8Encoding(false));
                                } catch (Exception e) { result.status = "blocked"; result.error += "\nEvidence export: " + e; Log("Evidence export error: " + e.Message); }
                            }
                            result.finished = DateTime.UtcNow.ToString("O"); Journal();
                            if (application != null && Marshal.IsComObject(application)) Marshal.ReleaseComObject(application);
                            application = null; controller = null;
                        }
                        if (!quiescent) throw new InvalidOperationException("Suite stopped because the previous request may still be running. Remaining cases were not submitted.");
                    }
                } catch (Exception e) { Log("Suite stopped: " + e); foreach (var r in Results.Where(r => r.status == "not_run")) r.error = "Not attempted: " + e.Message; }
                finally {
                    State.expires = DateTime.UtcNow; TestLabSuite.Save(State); Journal();
                    // Do not alter an unrelated session that replaced ours.
                    if (!keepCaptureActive && TestLab.Status()?.fixture_root.StartsWith(State.folder + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase) == true) TestLab.Disable();
                }
            }
        }
        private async Task Prepare(string caseId, string folder)
        {
            preparation = TestLabPreparation.Launch(caseId, true); var started = DateTime.UtcNow; string previous = "";
            while (true) {
                cancel.ThrowIfCancellationRequested();
                var report = TestLabPreparation.ReadReport(preparation); var log = TestLabPreparation.ReadLog(preparation);
                if (log.Length > previous.Length) Log(log.Substring(previous.Length).TrimEnd()); previous = log;
                File.WriteAllText(Path.Combine(folder, "preparation.log"), log);
                if (report != null) {
                    File.WriteAllText(Path.Combine(folder, "preparation.json"), TestLab.Serialize(report));
                    if (report.status == "failed") throw new InvalidOperationException("Preparation failed: " + string.Join("; ", report.remaining ?? new string[0]) + "\n" + log);
                    if (report.status == "finished") { Log(string.Join("\n", report.completed ?? new string[0])); return; }
                }
                if (DateTime.UtcNow - started > TimeSpan.FromSeconds(120)) throw new TimeoutException("Preparation exceeded 120 seconds. Check the visible Office startup/sign-in dialogs. " + log);
                await Task.Delay(500, cancel);
            }
        }
        private void ConnectOffice(string host)
        {
            application = Marshal.GetActiveObject(host + ".Application"); dynamic app = application;
            var progId = host == "Outlook" ? "Scribble.AddIn" : "Scribble." + host + "AddIn";
            controller = app.COMAddIns.Item(progId).Object;
            if (controller == null) throw new InvalidOperationException("The " + host + " add-in does not expose the suite runner. Install the latest Scribble and restart Office.");
        }
        private void SaveSelectedDeck()
        {
            dynamic app = application; object value = app.ActivePresentation; dynamic deck = value;
            if (!TestLab.IsRunOutput(value, State.runId) && !TestLabSuite.OwnsSource(State.runId, Convert.ToString(deck.FullName)))
                throw new InvalidOperationException("The active deck is not owned by this case.");
            var target = TestLab.SafeChild(State.folder, "cases/" + State.caseId + "/generated-deck.pptx");
            if (!TestLab.IsRunOutput(value, State.runId)) TestLab.RegisterOutput(value, "PowerPoint");
            deck.SaveAs(target, 24);
            TestLab.Collect(State.runId, target);
            Log("Saved and selected generated-deck.pptx for the email attachment.");
        }
        private void OpenChrome()
        {
            var session = TestLab.Status();
            var receipt = TestLabSuite.Read<Dictionary<string, object>>(Path.Combine(TestLab.Root, "fixture-server-" + session.session_id + ".json"));
            State.sourceUrl = "http://127.0.0.1:" + Convert.ToInt32(receipt["port"]) + "/operations.html"; TestLabSuite.Save(State);
            var chrome = new[] { Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) }.Select(p => Path.Combine(p, @"Google\Chrome\Application\chrome.exe")).First(File.Exists);
            Process.Start(new ProcessStartInfo(chrome, "--new-window \"chrome-extension://olkepladbgkfkhlglooilnmalckpdada/sidepanel.html?suite=" + State.id + "&token=" + State.chromeToken + "\"") { UseShellExecute = false });
        }
        private async Task Command(string action, int phase, int timeout, bool cancellable)
        {
            var id = Guid.NewGuid().ToString("N"); var started = DateTime.UtcNow; var progress = started; bool accepted = false;
            var folder = TestLab.SafeChild(State.folder, "cases/" + State.caseId);
            if (State.host == "Chrome") {
                var command = new SuiteChromeCommand { id = id, action = action, prompt = action == "submit" ? TestLabSuite.Prompt(TestLabSuite.CurrentCase(State), phase) : null, sourceUrl = State.sourceUrl, runId = State.runId };
                var file = Path.Combine(folder, "chrome-command.json"); var temp = file + ".tmp"; File.WriteAllText(temp, TestLab.Serialize(command)); if (File.Exists(file)) File.Replace(temp, file, null); else File.Move(temp, file);
            }
            Log(State.caseId + ": " + action + (action == "submit" ? " phase " + phase : ""));
            while (DateTime.UtcNow - started < TimeSpan.FromSeconds(timeout)) {
                if (cancellable) cancel.ThrowIfCancellationRequested();
                SuiteReply reply = null;
                if (State.host == "Chrome") {
                    var file = Path.Combine(folder, id + ".reply.json");
                    try { if (File.Exists(file)) reply = TestLabSuite.Read<SuiteReply>(file); } catch (IOException) { } catch (ArgumentException) { }
                } else {
                    try {
                        var json = Convert.ToString(controller.RunTestLabCommand(State.id, id, accepted && action != "stop" ? "status" : action, phase));
                        reply = new JavaScriptSerializer().Deserialize<SuiteReply>(json);
                        if (reply.state != "initializing") accepted = true;
                    } catch (COMException e) when ((uint)e.HResult == 0x80010001 || (uint)e.HResult == 0x8001010A) { /* Office is temporarily busy. */ }
                }
                if (reply != null) {
                    if (!string.IsNullOrEmpty(reply.error) && action != "stop") throw new InvalidOperationException(reply.error);
                    if (reply.state == "done") return;
                }
                if (DateTime.UtcNow - progress > TimeSpan.FromSeconds(15)) { Log(State.caseId + ": waiting for " + action + " (" + (int)(DateTime.UtcNow - started).TotalSeconds + " seconds)"); progress = DateTime.UtcNow; }
                await Task.Delay(750);
            }
            throw new TimeoutException(State.host + " did not complete " + action + " in " + timeout + " seconds. Check its visible Scribble pane and model connection.");
        }
        private async Task CancellationCase(LabCase c)
        {
            // The stop boundary must be observed, not simulated with a fixed delay.
            var pending = Command("submit", 0, 600, true);
            var reached = false;
            while (!pending.IsCompleted) {
                cancel.ThrowIfCancellationRequested();
                object excel = null;
                try {
                    excel = Marshal.GetActiveObject("Excel.Application"); dynamic app = excel;
                    for (int i = 1; i <= (int)app.Workbooks.Count; i++) if (TestLab.IsRunOutput((object)app.Workbooks.Item(i), State.runId)) reached = true;
                } catch (COMException) { }
                finally { if (excel != null && Marshal.IsComObject(excel)) Marshal.ReleaseComObject(excel); }
                if (reached) {
                    object ppt = null;
                    try { ppt = Marshal.GetActiveObject("PowerPoint.Application"); dynamic app = ppt;
                        for (int i = 1; i <= (int)app.Presentations.Count; i++) if (TestLab.IsRunOutput((object)app.Presentations.Item(i), State.runId) && (int)app.Presentations.Item(i).Slides.Count >= 6)
                            throw new InvalidOperationException("RC01 stop boundary was missed: the deck already has six slides.");
                    } catch (COMException) { } finally { if (ppt != null && Marshal.IsComObject(ppt)) Marshal.ReleaseComObject(ppt); }
                    TestLab.Marker("Automatic RC01 stop: run-owned workbook observed"); await Command("stop", 0, 60, false); break; }
                await Task.Delay(250);
            }
            try { await pending; } catch (Exception) when (reached) { }
            if (!reached) throw new InvalidOperationException("RC01 cancellation boundary was not observed before the request finished.");
            await Command("submit", 1, 600, true);
        }
    }
}
