using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Web.Script.Serialization;

namespace Scribble.Chat
{
    // The extension transports only the current exchange. Native encrypted state
    // owns the complete request, pending calls, original answers and UI recovery.
    public static class BrowserTaskSession
    {
        public static string Id(string chat, string turn)
        {
            if (string.IsNullOrWhiteSpace(chat) || string.IsNullOrWhiteSpace(turn)) throw new ArgumentException("Browser task identity is required.");
            return TaskCheckpointStore.Fingerprint(chat + "\n" + turn).Substring(0, 32);
        }
        public static DurableTaskState Load(string chat, string turn)
        {
            try { return new TaskCheckpointStore().Load(Id(chat, turn)); }
            catch (FileNotFoundException) { return null; }
            catch (DirectoryNotFoundException) { return null; }
        }
        public static void SaveUi(string chat, string turn, string prompt, string data)
        {
            if (data == null || data.Length > 700000) throw new ArgumentException("The browser recovery page exceeds the transport budget.");
            var state = Load(chat, turn) ?? new DurableTaskState { Id = Id(chat, turn), Host = "chrome", Objective = prompt };
            if (state.Objective != prompt) throw new InvalidOperationException("The original browser instruction cannot change during a task.");
            state.HostData["recovery_input"] = "{}";
            state.HostData["browser_ui"] = data;
            var ui = new JavaScriptSerializer { MaxJsonLength = 700000 }.Deserialize<Dictionary<string, object>>(data);
            object paused;
            state.UserPaused = ui.TryGetValue("userPaused", out paused) && Convert.ToBoolean(paused);
            if (state.UserPaused) state.Lifecycle = TaskLifecycle.Paused;
            new TaskCheckpointStore().Save(state);
        }
        public static string RecoverUi()
        {
            var states = new TaskCheckpointStore().FindUnfinished("chrome");
            var json = new JavaScriptSerializer { MaxJsonLength = 900000 };
            var first = states.FirstOrDefault(s => s.HostData.ContainsKey("browser_ui"));
            return json.Serialize(new { available = first != null, unique = states.Count == 1,
                state = first == null ? null : first.HostData["browser_ui"] });
        }
        public static void Discard(string chat, string turn) { new TaskCheckpointStore().Discard(Id(chat, turn)); }
        public static void Pause(string chat, string turn)
        {
            var state = Load(chat, turn);
            if (state == null) return;
            state.UserPaused = true; state.Lifecycle = TaskLifecycle.Paused;
            state.Blocker = "Paused by user. Resume from the saved browser state.";
            new TaskCheckpointStore().Save(state);
        }
        public static void ThrowIfPaused(string chat, string turn)
        {
            if (Load(chat, turn)?.UserPaused == true) throw new OperationCanceledException("Browser task paused by user.");
        }
        public static string CoverageNote(DurableTaskState state)
        {
            string data;
            if (!state.HostData.TryGetValue("browser_ui", out data)) return "";
            var ui = new JavaScriptSerializer { MaxJsonLength = 900000 }.Deserialize<Dictionary<string, object>>(data);
            object conditions, evidence, complete, answers;
            ui.TryGetValue("conditions", out conditions); ui.TryGetValue("evidence", out evidence);
            ui.TryGetValue("conditionEnumerationComplete", out complete); ui.TryGetValue("answers", out answers);
            return "Host browser coverage receipt (original answers remain authoritative): " + new JavaScriptSerializer { MaxJsonLength = 900000 }.Serialize(new { conditions, evidence, enumeration_complete = complete, answers });
        }
    }
}
