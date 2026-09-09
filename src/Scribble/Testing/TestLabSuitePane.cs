using System;
using System.Collections.Generic;
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
            if (action == "stop") { stop(); return TestLab.Serialize(new SuiteReply { state = busy() || state == "running" ? "running" : "done", error = error }); }
            if (!ready) return TestLab.Serialize(new SuiteReply { state = "initializing" });
            if (action == "status") return TestLab.Serialize(new SuiteReply { state = busy() || state == "running" ? "running" : state, error = error });
            if (id == commandId) return TestLab.Serialize(new SuiteReply { state = state, error = error, hostModule = typeof(TestLab).Assembly.ManifestModule.ModuleVersionId.ToString(), captureRoot = TestLab.Root });
            if (busy() || state == "running") throw new InvalidOperationException("The pane already has a request in progress.");
            var c = TestLabSuite.CurrentCase(suite);
            if (action != "load" && action != "submit") throw new InvalidOperationException("Unknown suite action.");
            var prompt = TestLabSuite.Prompt(c, phase);
            commandId = id; runId = suite.runId; error = null;
            if (action == "load") {
                TestLab.Record(runId, "suite", "host_connected", new { host,
                    loaded_module = typeof(TestLab).Assembly.ManifestModule.ModuleVersionId.ToString(),
                    assembly = typeof(TestLab).Assembly.Location, capture_root = TestLab.Root });
                reset(); load(c); state = "done";
                return TestLab.Serialize(new SuiteReply { state = state, hostModule = typeof(TestLab).Assembly.ManifestModule.ModuleVersionId.ToString(), captureRoot = TestLab.Root });
            }
            else { state = "running"; Execute(send, prompt); }
            return TestLab.Serialize(new SuiteReply { state = busy() || state == "running" ? "running" : state, error = error });
        }
        private async void Execute(Func<string, Task> send, string prompt)
        {
            try { await send(prompt); }
            catch (Exception e) { error = e.ToString(); }
            finally { state = "done"; }
        }
    }
    public sealed class SuiteReply { public string hostModule { get; set; } public string captureRoot { get; set; } public string state { get; set; } public string error { get; set; } }
}
