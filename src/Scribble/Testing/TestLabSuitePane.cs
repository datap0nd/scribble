using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace Scribble.Testing
{
    // Runs the same visible pane entry points as a human. No model-tool entry point.
    internal sealed class TestLabSuitePane
    {
        private string commandId;
        private string state = "ready";
        private string error;
        private string runId;
        private string commandAction;
        private int commandPhase;
        private string commandHost;
        public void Observe(IDictionary<string, object> payload, Action<string> answer)
        {
            if (runId == null || runId != TestLab.ActiveRunId()) return;
            object type, value;
            payload.TryGetValue("type", out type);
            if (Convert.ToString(type) == "askUser") {
                var suite = TestLabSuite.Active();
                var question = payload.TryGetValue("question", out value) ? Convert.ToString(value) : "";
                var preset = suite?.runId == runId ? TestLabSuite.PresetAnswer(TestLabSuite.CurrentCase(suite), question) : null;
                if (preset == null) error = "The model requested an answer outside the kit presets. See the captured question.";
                else { TestLab.Record(runId, "suite", "preset_answer", new { question, answer = preset }); answer(preset); }
            }
            if (Convert.ToString(type) == "status" && payload.TryGetValue("error", out value) && Equals(value, true))
                error = payload.TryGetValue("text", out value) ? Convert.ToString(value) : "Model request failed.";
        }
        public string Command(string suiteId, string id, string action, int phase, string host,
            bool ready, Func<bool> busy, Action reset, Action<LabCase> load, Func<string, Task> send, Action stop)
        {
            var suite = TestLabSuite.Require(suiteId, host);
            if (action == "stop") { stop(); WriteReceipt(suite.runId, id, "stop", phase, host, "running", error); return TestLab.Serialize(new SuiteReply { state = busy() || state == "running" ? "running" : "done", error = error }); }
            if (!ready) return TestLab.Serialize(new SuiteReply { state = "initializing" });
            // The Office process owns the live request. Its memory is the
            // authority for retries and status polling; disk receipts exist
            // only for crash recovery and diagnostics. Reading the receipt
            // first allowed corporate DRM to replace JSON with ciphertext and
            // throw across the COM boundary on every status poll.
            if (runId == suite.runId && id == commandId) {
                if (action == "status" && commandAction == "load" && state == "running" && !busy()) {
                    state = "done";
                    WriteReceipt(runId, id, commandAction, commandPhase, commandHost, state, error);
                }
                return TestLab.Serialize(new SuiteReply { state = busy() || state == "running" ? "running" : state, error = error,
                    hostModule = typeof(TestLab).Assembly.ManifestModule.ModuleVersionId.ToString(), captureRoot = TestLab.Root });
            }
            var receipt = ReadReceipt(suite.runId, id);
            if (action == "status") {
                if (receipt != null) {
                    // Loading attachments is asynchronous in the visible pane.
                    // Do not acknowledge the load until that ordinary reader is
                    // idle, or the immediately following submit is rejected as
                    // an overlapping request.
                    if (receipt.action == "load" && receipt.state == "running" && !busy()) {
                        state = "done";
                        WriteReceipt(suite.runId, id, receipt.action, receipt.phase, receipt.host, state, error);
                        receipt.state = state; receipt.error = error;
                    }
                    return TestLab.Serialize(new SuiteReply { state = receipt.state, error = receipt.error,
                        hostModule = typeof(TestLab).Assembly.ManifestModule.ModuleVersionId.ToString(), captureRoot = TestLab.Root });
                }
                return TestLab.Serialize(new SuiteReply { state = busy() || state == "running" ? "running" : state, error = error });
            }
            if (receipt != null && receipt.action == "submit") return TestLab.Serialize(new SuiteReply { state = receipt.state, error = receipt.error });
            if (busy() || state == "running") throw new InvalidOperationException("The pane already has a request in progress.");
            var c = TestLabSuite.CurrentCase(suite);
            if (action != "load" && action != "submit") throw new InvalidOperationException("Unknown suite action.");
            var prompt = TestLabSuite.Prompt(c, phase);
            commandId = id; commandAction = action; commandPhase = phase; commandHost = host; runId = suite.runId; error = null;
            if (action == "load") {
                WriteReceipt(runId, id, action, phase, host, "running", null);
                TestLab.Record(runId, "suite", "host_connected", new { host,
                    loaded_module = typeof(TestLab).Assembly.ManifestModule.ModuleVersionId.ToString(),
                    assembly = typeof(TestLab).Assembly.Location, capture_root = TestLab.Root });
                reset(); load(c); state = busy() ? "running" : "done";
                WriteReceipt(runId, id, action, phase, host, state, null);
                return TestLab.Serialize(new SuiteReply { state = state, hostModule = typeof(TestLab).Assembly.ManifestModule.ModuleVersionId.ToString(), captureRoot = TestLab.Root });
            }
            else { state = "running"; WriteReceipt(runId, id, action, phase, host, state, null); Execute(send, prompt, id, phase, host); }
            return TestLab.Serialize(new SuiteReply { state = busy() || state == "running" ? "running" : state, error = error });
        }

        private async void Execute(Func<string, Task> send, string prompt, string id, int phase, string host)
        {
            try { await send(prompt); }
            catch (Exception e) { error = e.ToString(); }
            finally { state = "done"; WriteReceipt(runId, id, "submit", phase, host, state, error); }
        }

        private static string ReceiptPath(string runId, string id)
        {
            if (!System.Text.RegularExpressions.Regex.IsMatch(id ?? "", "^[a-f0-9]{32}$")) throw new InvalidDataException("Invalid suite command ID.");
            var directory = Path.Combine(TestLab.RunDirectory(runId), "commands"); Directory.CreateDirectory(directory);
            return Path.Combine(directory, id + ".json");
        }

        private static SuiteCommandReceipt ReadReceipt(string runId, string id)
        {
            try { var path = ReceiptPath(runId, id); return File.Exists(path) ? TestLabSuite.Read<SuiteCommandReceipt>(path) : null; }
            catch (Exception) { return null; }
        }

        private static void WriteReceipt(string runId, string id, string action, int phase, string host, string commandState, string commandError)
        {
            try
            {
                string payload;
                using (var process = System.Diagnostics.Process.GetCurrentProcess())
                {
                    payload = TestLab.Serialize(new SuiteCommandReceipt { schema = 1, run_id = runId,
                    command_id = id, action = action, phase = phase, host = host, state = commandState,
                    pid = process.Id, processStart = process.StartTime.ToUniversalTime().Ticks,
                    error = commandError, updated_utc = DateTime.UtcNow.ToString("O") });
                    var session = TestLab.Status();
                    if (session != null && !string.IsNullOrEmpty(session.transport_pipe) &&
                        (session.transport_pid != process.Id || session.transport_process_start != process.StartTime.ToUniversalTime().Ticks))
                        TestLabTransport.Send(session.transport_pipe, "receipt", runId, id, payload);
                    else PersistTransportedReceipt(runId, id, payload);
                }
            }
            catch (Exception) { /* A diagnostic receipt must never terminate its Office host. */ }
        }

        internal static void PersistTransportedReceipt(string runId, string id, string payload)
        {
            var path = ReceiptPath(runId, id);
            var receipt = new JavaScriptSerializer().Deserialize<SuiteCommandReceipt>(payload);
            if (receipt == null || receipt.run_id != runId || receipt.command_id != id)
                throw new InvalidDataException("Transported command receipt identity mismatch.");
            var temporary = path + "." + Guid.NewGuid().ToString("N");
            File.WriteAllText(temporary, payload, new UTF8Encoding(false));
            if (File.Exists(path)) File.Replace(temporary, path, null); else new FileInfo(temporary).MoveTo(path);
        }
    }
    public sealed class SuiteCommandReceipt
    {
        public int schema { get; set; }
        public string run_id { get; set; }
        public string command_id { get; set; }
        public string action { get; set; }
        public int phase { get; set; }
        public string host { get; set; }
        public int pid { get; set; }
        public long processStart { get; set; }
        public string state { get; set; }
        public string error { get; set; }
        public string updated_utc { get; set; }
    }
    public sealed class SuiteReply { public string hostModule { get; set; } public string captureRoot { get; set; } public string state { get; set; } public string error { get; set; } }
}
