using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;

namespace Scribble.Testing
{
    // Operator-only APIs. None of these methods is exposed in a model tool catalog.
    public static class TestLab
    {
        private static readonly JavaScriptSerializer Json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
        private static readonly string Instance = Guid.NewGuid().ToString("N");
        private static long sequence;
        private static readonly object Gate = new object();
        public static string Root { get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Scribble", "TestLab"); } }
        private static string Descriptor { get { return Path.Combine(Root, "session.bin"); } }
        public static string Hash(byte[] bytes) { using (var sha = SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-", "").ToLowerInvariant(); }
        public static string FileHash(string path) { return Hash(File.ReadAllBytes(path)); }
        public static string Serialize(object value) { lock (Gate) return Json.Serialize(value); }
        private static T Read<T>(string path) { lock (Gate) return Json.Deserialize<T>(File.ReadAllText(path)); }
        private static void Write(string path, object value) { File.WriteAllText(path, Serialize(value), new UTF8Encoding(false)); }
        private static void Id(string id) { if (!Regex.IsMatch(id ?? "", "^[a-f0-9]{32}$")) throw new InvalidDataException("Invalid run ID."); }
        public static string RunDirectory(string id) { Id(id); return Path.Combine(Root, "runs", id); }
        public static string SafeChild(string root, string relative)
        {
            if (string.IsNullOrWhiteSpace(relative) || Path.IsPathRooted(relative)) throw new InvalidDataException("Relative path required.");
            var parent = Path.GetFullPath(root).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
            var result = Path.GetFullPath(Path.Combine(parent, relative));
            if (!result.StartsWith(parent, StringComparison.OrdinalIgnoreCase)) throw new InvalidDataException("Path escapes root.");
            var cursor = result;
            while (cursor.Length >= parent.TrimEnd(Path.DirectorySeparatorChar).Length)
            {
                if ((File.Exists(cursor) || Directory.Exists(cursor)) && (File.GetAttributes(cursor) & FileAttributes.ReparsePoint) != 0)
                    throw new InvalidDataException("Links are not permitted in benchmark paths.");
                cursor = Path.GetDirectoryName(cursor);
                if (cursor == null) break;
            }
            return result;
        }
        private static IDisposable SessionLock()
        {
            Directory.CreateDirectory(Root);
            for (int i = 0; i < 100; i++)
            {
                try { return new FileStream(Path.Combine(Root, "operator.lock"), FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None); }
                catch (IOException) { System.Threading.Thread.Sleep(20); }
            }
            throw new IOException("Another test operation is running.");
        }
        private static void SaveSession(LabSession value)
        {
            var bytes = ProtectedData.Protect(Encoding.UTF8.GetBytes(Serialize(value)), null, DataProtectionScope.CurrentUser);
            var temporary = Descriptor + ".tmp";
            File.WriteAllBytes(temporary, bytes);
            if (File.Exists(Descriptor)) File.Replace(temporary, Descriptor, null); else File.Move(temporary, Descriptor);
        }
        public static LabSession Status()
        {
            try
            {
                if (!File.Exists(Descriptor)) return null;
                LabSession s;
                lock (Gate) s = Json.Deserialize<LabSession>(Encoding.UTF8.GetString(ProtectedData.Unprotect(File.ReadAllBytes(Descriptor), null, DataProtectionScope.CurrentUser)));
                if (s.schema != 1 || s.expires_utc <= DateTime.UtcNow || !Directory.Exists(s.fixture_root)) return null;
                Id(s.session_id);
                if (!string.IsNullOrEmpty(s.run_id)) Id(s.run_id);
                return s;
            }
            catch { return null; }
        }
        public static void Enable(string fixtureRoot)
        {
            fixtureRoot = Path.GetFullPath(fixtureRoot);
            var manifest = VerifyKit(fixtureRoot);
            using (SessionLock())
            {
                var current = Status();
                if (current != null && !string.IsNullOrEmpty(current.run_id)) throw new InvalidOperationException("Finish the active case before enabling another kit.");
                SaveSession(new LabSession { schema = 1, session_id = Guid.NewGuid().ToString("N"), fixture_root = fixtureRoot,
                    manifest_sha256 = FileHash(Path.Combine(fixtureRoot, "manifest.json")), expires_utc = DateTime.UtcNow.AddHours(8), suite_id = manifest.suite_id });
            }
        }
        public static void Disable()
        {
            using (SessionLock())
            {
                var s = Status();
                if (s != null && !string.IsNullOrEmpty(s.run_id)) MarkIncomplete(s.run_id, "Capture disabled before case finish.");
                if (File.Exists(Descriptor)) File.Delete(Descriptor);
            }
        }
        public static KitManifest VerifyKit(string root)
        {
            var manifest = Read<KitManifest>(SafeChild(root, "manifest.json"));
            if (manifest.schema != 1 || manifest.suite_id != "atlas-v1" || manifest.files == null || manifest.files.Length == 0)
                throw new InvalidDataException("Unsupported or empty fixture manifest.");
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var f in manifest.files)
            {
                if (!seen.Add(f.path)) throw new InvalidDataException("Duplicate manifest path.");
                var path = SafeChild(root, f.path);
                if (!File.Exists(path) || new FileInfo(path).Length != f.size || FileHash(path) != f.sha256)
                    throw new InvalidDataException("Fixture verification failed: " + f.path);
            }
            return manifest;
        }
        public static LabCase[] Cases()
        {
            var s = Status();
            return s == null ? new LabCase[0] : Read<LabCase[]>(SafeChild(s.fixture_root, "operator/cases.json"));
        }
        public static LabRun Start(string caseId, string host, bool fixtureOnlyConfirmed)
        {
            if (!fixtureOnlyConfirmed) throw new InvalidOperationException("Confirm the isolated synthetic context first.");
            using (SessionLock())
            {
                var s = Status();
                if (s == null) throw new InvalidOperationException("Test Lab is disabled or expired.");
                if (!string.IsNullOrEmpty(s.run_id)) throw new InvalidOperationException("Finish the active case first.");
                VerifyKit(s.fixture_root);
                if (FileHash(Path.Combine(s.fixture_root, "manifest.json")) != s.manifest_sha256) throw new InvalidDataException("Manifest changed since activation.");
                var c = Cases().Single(x => x.id == caseId);
                var allowedPaths = new HashSet<string>(c.inputs ?? new string[0], StringComparer.OrdinalIgnoreCase);
                foreach (var source in Read<MailFixture[]>(SafeChild(s.fixture_root, "operator/mail-index.json")))
                    if (allowedPaths.Contains(source.path)) foreach (var attachment in source.attachments ?? new string[0]) allowedPaths.Add(attachment);
                var settings = new Configuration.SettingsStore().Load();
                var run = new LabRun { schema = 1, run_id = Guid.NewGuid().ToString("N"), session_id = s.session_id,
                    suite_id = s.suite_id, case_id = c.id, host = host, started_utc = DateTime.UtcNow.ToString("O"),
                    status = "running", manifest_sha256 = s.manifest_sha256, fixture_root = s.fixture_root,
                    assembly_version = FileVersionInfo.GetVersionInfo(typeof(TestLab).Assembly.Location).FileVersion,
                    assembly_sha256 = FileHash(typeof(TestLab).Assembly.Location), os = Environment.OSVersion.ToString(),
                    process_bitness = IntPtr.Size * 8, locale = System.Globalization.CultureInfo.CurrentCulture.Name,
                    required_artifacts = c.artifacts ?? new string[0], input_paths = allowedPaths.ToArray(), prompt = c.prompt,
                    selected_model = settings.Model, writing_profile_enabled = settings.UseToneProfile,
                    writing_profile_hash = Hash(Encoding.UTF8.GetBytes(settings.ToneProfile ?? "")),
                    process_name = Process.GetCurrentProcess().ProcessName,
                    input_hashes = allowedPaths.Select(p => FileHash(SafeChild(s.fixture_root, p))).ToArray(),
                    trace_complete = true, evidence_status = "unreviewed", operator_context_attested = true };
                var folder = RunDirectory(run.run_id);
                Directory.CreateDirectory(Path.Combine(folder, "events"));
                Directory.CreateDirectory(Path.Combine(folder, "artifacts"));
                Write(Path.Combine(folder, "run.json"), run);
                s.run_id = run.run_id;
                SaveSession(s);
                Record(run.run_id, "operator", "case_started", new { case_id = c.id, host, prompt = c.prompt });
                return run;
            }
        }
        public static string ActiveRunId() { return Status()?.run_id; }
        public static string CaptureState()
        {
            var id = ActiveRunId();
            if (string.IsNullOrEmpty(id)) return "idle";
            return Directory.GetFiles(RunDirectory(id), "incomplete-*.json").Length > 0 ? "incomplete" : "capturing";
        }
        public static void CheckBrowserSource(string url)
        {
            var id = ActiveRunId(); if (string.IsNullOrEmpty(id)) return;
            Uri uri;
            if (!Uri.TryCreate(url, UriKind.Absolute, out uri) || uri.Scheme != "http" || !uri.IsLoopback ||
                !new[] { "/", "/index.html", "/operations.html", "/archive.html" }.Contains(uri.AbsolutePath) || uri.Query.Length > 0)
            { MarkIncomplete(id, "Browser source is outside the synthetic fixture site."); throw new InvalidOperationException("Test Lab requires the loopback synthetic fixture page."); }
        }
        public static void CheckOfficeSource(object application, string host)
        {
            var id = ActiveRunId(); if (string.IsNullOrEmpty(id) || application == null) return;
            dynamic app = application; dynamic document = null;
            try
            {
                switch ((host ?? "").ToLowerInvariant())
                {
                    case "excel": if (app.Workbooks.Count > 0) document = app.ActiveWorkbook; break;
                    case "word": if (app.Documents.Count > 0) document = app.ActiveDocument; break;
                    case "powerpoint": if (app.Presentations.Count > 0) document = app.ActivePresentation; break;
                }
                if (document == null) return;
                string path = Convert.ToString(document.Path);
                if (!IsRunOutput(document, id))
                {
                    if (string.IsNullOrEmpty(path)) { MarkIncomplete(id, "Unregistered unsaved document."); throw new InvalidOperationException("Open the saved case fixture before starting, or use a draft created during this run."); }
                    CheckInputFile(Convert.ToString(document.FullName));
                }
                Record(id, "source", "office_context", new { host, name = Convert.ToString(document.Name), app_version = Convert.ToString(app.Version), saved_source_checked = !string.IsNullOrEmpty(path) });
            }
            finally { if (document != null && System.Runtime.InteropServices.Marshal.IsComObject(document)) System.Runtime.InteropServices.Marshal.ReleaseComObject(document); }
        }
        public static bool IsRunOutput(object value, string runId)
        {
            try { dynamic document = value; return Convert.ToString(document.CustomDocumentProperties["ScribbleTestRunId"].Value) == runId; }
            catch { return false; }
        }
        public static void RegisterOutput(object value, string kind)
        {
            var id = ActiveRunId(); if (string.IsNullOrEmpty(id)) return;
            dynamic document = value;
            document.CustomDocumentProperties.Add("ScribbleTestRunId", false, 4, id);
            Record(id, "artifact", "output_registered", new { kind, name = Convert.ToString(document.Name) });
        }
        public static void RegisterMailOutput(object value)
        {
            var id = ActiveRunId(); if (string.IsNullOrEmpty(id)) return;
            dynamic mail = value;
            mail.UserProperties.Add("ScribbleTestRunId", 1, false).Value = id;
            Record(id, "artifact", "output_registered", new { kind = "msg", unsent = !Convert.ToBoolean(mail.Sent) });
        }
        public static void CheckMailSource(object item)
        {
            var id = ActiveRunId(); if (string.IsNullOrEmpty(id)) return;
            var run = GetRun(id); dynamic mail = item;
            var sources = Read<MailFixture[]>(SafeChild(run.fixture_root, "operator/mail-index.json"));
            string subject = Convert.ToString(mail.Subject), sender = Convert.ToString(mail.SenderEmailAddress);
            var source = sources.FirstOrDefault(x => run.input_paths.Contains(x.path) && x.subject == subject && x.sender == sender);
            string body = Convert.ToString(mail.Body) ?? "";
            if (source == null || Regex.Replace(body, @"\s+", " ").Trim() != Regex.Replace(source.body ?? "", @"\s+", " ").Trim())
            { MarkIncomplete(id, "Selected email does not match the case fixture."); throw new InvalidOperationException("Test Lab email mismatch. Select the imported case messages."); }
            Record(id, "source", "mail_verified", new { source_id = source.path, body_sha256 = Hash(Encoding.UTF8.GetBytes(body)) });
        }
        public static LabRun GetRun(string id) { return Read<LabRun>(Path.Combine(RunDirectory(id), "run.json")); }
        internal static void MarkIncomplete(string id, string reason)
        {
            // A separate marker avoids competing read-modify-write updates from Office processes.
            Write(Path.Combine(RunDirectory(id), "incomplete-" + Instance + ".json"), new { reason, utc = DateTime.UtcNow.ToString("O") });
        }
        public static string Redact(string text)
        {
            // Credentials are never passed by transport hooks. Also remove accidental reflected secrets.
            try { var key = new Configuration.SettingsStore().Load().ApiKey; if (!string.IsNullOrEmpty(key)) text = text.Replace(key, "[REDACTED]"); } catch { }
            return Regex.Replace(text ?? "", @"(?i)(Bearer\s+)[A-Za-z0-9._~+/=-]+", "$1[REDACTED]");
        }
        public static void Record(string runId, string taskId, string stage, object detail)
        {
            if (string.IsNullOrEmpty(runId)) return;
            var s = Status();
            if (s == null || s.run_id != runId) return;
            try
            {
                var folder = RunDirectory(runId);
                if (Directory.GetFiles(folder, "incomplete-*.json").Length != 0) return;
                var events = Path.Combine(folder, "events");
                if (Directory.GetFiles(events).Sum(p => new FileInfo(p).Length) > 250L * 1024 * 1024)
                { MarkIncomplete(runId, "Trace quota exceeded. Recording stopped."); return; }
                string payload;
                lock (Gate) payload = Redact(Json.Serialize(new { schema = 1, run_id = runId, task_id = taskId, stage,
                    instance_id = Instance, process_id = Process.GetCurrentProcess().Id, sequence = ++sequence,
                    utc = DateTime.UtcNow.ToString("O"), monotonic_ticks = Stopwatch.GetTimestamp(),
                    monotonic_frequency = Stopwatch.Frequency, detail }));
                var encrypted = ProtectedData.Protect(Encoding.UTF8.GetBytes(payload), null, DataProtectionScope.CurrentUser);
                var path = Path.Combine(events, Instance + "-" + Guid.NewGuid().ToString("N") + ".bin");
                var temporary = path + ".writing";
                // Exporters enumerate only committed .bin files. Publishing
                // after a durable, exclusive write prevents a snapshot from
                // attempting to decrypt a partially written DPAPI payload. A
                // failed .writing file remains as narrow forensic evidence.
                using (var stream = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                { stream.Write(encrypted, 0, encrypted.Length); stream.Flush(true); }
                File.Move(temporary, path);
                var activePath = Path.Combine(folder, "task-" + Hash(Encoding.UTF8.GetBytes(taskId ?? "")) + ".active");
                if (stage == "task_started" || stage == "task_resumed") File.WriteAllText(activePath, taskId);
                if ((stage == "task_completed" || stage == "task_paused") && File.Exists(activePath)) File.Delete(activePath);
            }
            catch (Exception e)
            {
                try { MarkIncomplete(runId, "Trace write failed: " + e.GetType().Name); } catch { }
                throw new IOException("Benchmark capture failed; stop this case and export the incomplete evidence.", e);
            }
        }
        public static void Marker(string text)
        { var id = ActiveRunId(); if (id == null) throw new InvalidOperationException("No active run."); Record(id, "operator", "video_marker", new { text }); }
        public static void CheckInputFile(string path)
        {
            var id = ActiveRunId(); if (string.IsNullOrEmpty(id)) return;
            var run = GetRun(id);
            var full = Path.GetFullPath(path);
            var hash = FileHash(full);
            var allowed = (run.input_hashes ?? new string[0]).Contains(hash);
            var generated = Directory.GetFiles(Path.Combine(RunDirectory(id), "artifacts"), "*.receipt.json")
                .Any(p => { var receipt = Read<Dictionary<string, object>>(p); return Convert.ToString(receipt["sha256"]) == hash; });
            if (!allowed && !generated)
            {
                MarkIncomplete(id, "Source mismatch; capture stopped.");
                throw new InvalidOperationException("Test Lab source mismatch. Use a verified synthetic input or finish this case.");
            }
            Record(id, "source", "input_verified", new { name = Path.GetFileName(path), sha256 = hash });
        }
        public static void Finish(bool assisted)
        {
            using (SessionLock())
            {
                var s = Status(); if (s == null || string.IsNullOrEmpty(s.run_id)) throw new InvalidOperationException("No active run.");
                var run = GetRun(s.run_id);
                if (Directory.GetFiles(RunDirectory(run.run_id), "*.active").Length > 0)
                    MarkIncomplete(run.run_id, "Case finished while a task had no completed or paused event.");
                Record(run.run_id, "operator", "case_finished", new { assisted });
                run.status = "finished"; run.finished_utc = DateTime.UtcNow.ToString("O"); run.assisted = assisted;
                run.trace_complete = Directory.GetFiles(RunDirectory(run.run_id), "incomplete-*.json").Length == 0;
                Write(Path.Combine(RunDirectory(run.run_id), "run.json"), run);
                s.run_id = null; SaveSession(s);
            }
        }
        public static string Collect(string runId, string sourcePath)
        {
            var run = GetRun(runId);
            var extension = Path.GetExtension(sourcePath).ToLowerInvariant();
            if (!new[] { ".xlsx", ".pptx", ".docx", ".pdf", ".png", ".msg", ".eml", ".html", ".json" }.Contains(extension))
                throw new InvalidDataException("Unsupported artifact type.");
            if (new FileInfo(sourcePath).Length > 100L * 1024 * 1024) throw new InvalidDataException("Artifact exceeds 100 MB.");
            var sourceHash = FileHash(sourcePath);
            var kit = Read<KitManifest>(SafeChild(run.fixture_root, "manifest.json"));
            if (kit.files.Any(f => f.sha256 == sourceHash)) throw new InvalidDataException("A fixture or reference file cannot be collected as a generated output.");
            var directory = Path.Combine(RunDirectory(runId), "artifacts");
            if (Directory.GetFiles(directory).Sum(p => new FileInfo(p).Length) > 500L * 1024 * 1024) throw new IOException("Artifact budget exceeded.");
            var target = SafeChild(directory, Guid.NewGuid().ToString("N").Substring(0, 8) + "-" + Path.GetFileName(sourcePath));
            File.Copy(sourcePath, target, false);
            Write(target + ".receipt.json", new { schema = 1, run_id = run.run_id, name = Path.GetFileName(target),
                sha256 = FileHash(target), size = new FileInfo(target).Length, captured_utc = DateTime.UtcNow.ToString("O"),
                provenance = "operator_selected_saved_output", native_readback_verified = false });
            return target;
        }
        public static string Export(string runId, string destinationDirectory) { return ExportCore(runId, destinationDirectory, false); }
        internal static string ExportSnapshot(string runId, string destinationDirectory) { return ExportCore(runId, destinationDirectory, true); }
        private static string ExportCore(string runId, string destinationDirectory, bool snapshot)
        {
            using (SessionLock())
            {
                var folder = RunDirectory(runId); var run = GetRun(runId);
                if (!snapshot && Status()?.run_id == runId) throw new InvalidOperationException("Finish the case before exporting.");
                var stage = Path.Combine(Root, "exports", Guid.NewGuid().ToString("N")); Directory.CreateDirectory(stage);
                var timeline = new List<string>();
                foreach (var file in Directory.GetFiles(Path.Combine(folder, "events"), "*.bin"))
                {
                    try { timeline.Add(Encoding.UTF8.GetString(ProtectedData.Unprotect(File.ReadAllBytes(file), null, DataProtectionScope.CurrentUser))); }
                    catch (Exception error) {
                        run.trace_complete = false;
                        Write(Path.Combine(stage, "incomplete-read-" + Path.GetFileNameWithoutExtension(file) + ".json"), new {
                            reason = "Recorded event could not be read: " + error.GetType().Name + ": " + error.Message,
                            file = Path.GetFileName(file), utc = DateTime.UtcNow.ToString("O") });
                    }
                }
                var ordered = timeline.Select(t => new { text = t, data = Json.Deserialize<Dictionary<string, object>>(t) })
                    .OrderBy(t => Convert.ToString(t.data["utc"])).ThenBy(t => Convert.ToString(t.data["instance_id"]))
                    .ThenBy(t => Convert.ToInt64(t.data["sequence"])).ToArray();
                File.WriteAllLines(Path.Combine(stage, "timeline.jsonl"), ordered.Select(t => t.text), new UTF8Encoding(false));
                File.WriteAllLines(Path.Combine(stage, "video-markers.csv"), new[] { "utc,marker" }.Concat(ordered.Where(t => Convert.ToString(t.data["stage"]) == "video_marker")
                    .Select(t => Convert.ToString(t.data["utc"]) + ",\"" + Serialize(t.data["detail"]).Replace("\"", "\"\"") + "\"")));
                Directory.CreateDirectory(Path.Combine(stage, "artifacts"));
                foreach (var file in Directory.GetFiles(Path.Combine(folder, "artifacts"))) File.Copy(file, Path.Combine(stage, "artifacts", Path.GetFileName(file)));
                var commands = Path.Combine(folder, "commands");
                if (Directory.Exists(commands)) {
                    Directory.CreateDirectory(Path.Combine(stage, "commands"));
                    foreach (var file in Directory.GetFiles(commands, "*.json")) File.Copy(file, Path.Combine(stage, "commands", Path.GetFileName(file)));
                }
                foreach (var file in Directory.GetFiles(folder, "incomplete-*.json")) File.Copy(file, Path.Combine(stage, Path.GetFileName(file)));
                var extensions = Directory.GetFiles(Path.Combine(stage, "artifacts")).Select(Path.GetExtension).ToArray();
                run.missing_artifacts = run.required_artifacts.Where(x => !extensions.Contains("." + x)).ToArray();
                run.trace_complete = run.trace_complete && run.status == "finished" && timeline.Count > 0 && Directory.GetFiles(folder, "incomplete-*.json").Length == 0 && Directory.GetFiles(folder, "*.active").Length == 0;
                Write(Path.Combine(stage, "run.json"), run);
                Write(Path.Combine(stage, "scorecard.json"), new { schema = 1, run_id = runId, status = "pending_evaluation", passed = false,
                    trace_complete = run.trace_complete, missing_artifacts = run.missing_artifacts, assisted = run.assisted,
                    note = "Artifact presence is not correctness. Run evaluate.py and complete native/visual review." });
                Write(Path.Combine(stage, "export-manifest.json"), new { schema = 1, files = Directory.GetFiles(stage, "*", SearchOption.AllDirectories).Select(p => new {
                    path = p.Substring(stage.Length + 1).Replace('\\', '/'), sha256 = FileHash(p), size = new FileInfo(p).Length }).ToArray() });
                Directory.CreateDirectory(destinationDirectory);
                var output = Path.Combine(destinationDirectory, "run-" + runId + "-" + DateTime.UtcNow.ToString("yyyyMMddHHmmssfff") + ".zip");
                ZipFile.CreateFromDirectory(stage, output, CompressionLevel.Optimal, false);
                return output;
            }
        }
    }
    public sealed class LabSession
    { public int schema { get; set; } public string session_id { get; set; } public string suite_id { get; set; } public string fixture_root { get; set; } public string manifest_sha256 { get; set; } public DateTime expires_utc { get; set; } public string run_id { get; set; } }
    public sealed class KitManifest
    { public int schema { get; set; } public string suite_id { get; set; } public KitFile[] files { get; set; } }
    public sealed class KitFile
    { public string path { get; set; } public string sha256 { get; set; } public long size { get; set; } public string role { get; set; } }
    public sealed class LabCase
    { public Dictionary<string, string> clarification_answers { get; set; } public string id { get; set; } public string host { get; set; } public string prompt { get; set; } public string prerequisite_prompt { get; set; } public string setup { get; set; } public string[] inputs { get; set; } public string[] artifacts { get; set; } public override string ToString() { return id + " / " + host; } }
    public sealed class LabRun
    {
        public int schema { get; set; } public string run_id { get; set; } public string session_id { get; set; } public string suite_id { get; set; }
        public string case_id { get; set; } public string host { get; set; } public string started_utc { get; set; } public string finished_utc { get; set; }
        public string status { get; set; } public string fixture_root { get; set; } public string manifest_sha256 { get; set; }
        public string assembly_version { get; set; } public string assembly_sha256 { get; set; } public string os { get; set; } public int process_bitness { get; set; }
        public string locale { get; set; } public string[] required_artifacts { get; set; } public string[] input_paths { get; set; } public string prompt { get; set; }
        public bool trace_complete { get; set; } public bool assisted { get; set; } public bool operator_context_attested { get; set; }
        public string evidence_status { get; set; } public string[] missing_artifacts { get; set; }
        public string[] input_hashes { get; set; } public string selected_model { get; set; } public bool writing_profile_enabled { get; set; }
        public string writing_profile_hash { get; set; } public string process_name { get; set; }
    }
    public sealed class MailFixture { public string path { get; set; } public string[] attachments { get; set; } public string subject { get; set; } public string sender { get; set; } public string body { get; set; } }
}
