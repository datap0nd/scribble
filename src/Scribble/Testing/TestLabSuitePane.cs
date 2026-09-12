using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace Scribble.Testing
{
    // Runs the same visible pane entry points as a human. No model-tool entry point.
    internal sealed class TestLabSuitePane
    {
        private string commandId;
        private string state = "ready";
        private string error;
        private string runId;
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
            var receipt = ReadReceipt(suite.runId, id);
            if (action == "stop") { stop(); WriteReceipt(suite.runId, id, "stop", phase, host, "running", error); return TestLab.Serialize(new SuiteReply { state = busy() || state == "running" ? "running" : "done", error = error }); }
            if (!ready) return TestLab.Serialize(new SuiteReply { state = "initializing" });
            if (action == "status") {
                if (receipt != null) return TestLab.Serialize(new SuiteReply { state = receipt.state, error = receipt.error,
                    hostModule = typeof(TestLab).Assembly.ManifestModule.ModuleVersionId.ToString(), captureRoot = TestLab.Root });
                return TestLab.Serialize(new SuiteReply { state = busy() || state == "running" ? "running" : state, error = error });
            }
            if (receipt != null && receipt.action == "submit") return TestLab.Serialize(new SuiteReply { state = receipt.state, error = receipt.error });
            if (id == commandId) return TestLab.Serialize(new SuiteReply { state = state, error = error, hostModule = typeof(TestLab).Assembly.ManifestModule.ModuleVersionId.ToString(), captureRoot = TestLab.Root });
            if (busy() || state == "running") throw new InvalidOperationException("The pane already has a request in progress.");
            var c = TestLabSuite.CurrentCase(suite);
            if (action != "load" && action != "submit") throw new InvalidOperationException("Unknown suite action.");
            var prompt = TestLabSuite.Prompt(c, phase);
            commandId = id; runId = suite.runId; error = null;
            if (action == "load") {
                WriteReceipt(runId, id, action, phase, host, "running", null);
                TestLab.Record(runId, "suite", "host_connected", new { host,
                    loaded_module = typeof(TestLab).Assembly.ManifestModule.ModuleVersionId.ToString(),
                    assembly = typeof(TestLab).Assembly.Location, capture_root = TestLab.Root });
                reset(); load(c); state = "done";
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
            catch (IOException) { return null; }
        }

        private static void WriteReceipt(string runId, string id, string action, int phase, string host, string commandState, string commandError)
        {
            var path = ReceiptPath(runId, id); var temporary = path + "." + Guid.NewGuid().ToString("N");
            using (var process = System.Diagnostics.Process.GetCurrentProcess())
                File.WriteAllText(temporary, TestLab.Serialize(new SuiteCommandReceipt { schema = 1, run_id = runId,
                    command_id = id, action = action, phase = phase, host = host, state = commandState,
                    pid = process.Id, processStart = process.StartTime.ToUniversalTime().Ticks,
                    error = commandError, updated_utc = DateTime.UtcNow.ToString("O") }), new UTF8Encoding(false));
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
