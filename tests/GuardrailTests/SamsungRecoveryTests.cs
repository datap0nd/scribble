using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Web.Script.Serialization;
using Scribble.Chat;
using Scribble.Office;

namespace GuardrailTests
{
    internal static class SamsungRecoveryTests
    {
        private static readonly Assembly Assembly = typeof(SamsungAuthoringPolicy).Assembly;
        private static Type Type(string name) { return Assembly.GetType("Scribble.Office." + name, true); }
        private static object Invoke(Type type, string method, object instance, params object[] args)
        {
            try { return type.GetMethods(BindingFlags.Static | BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic).Single(m => m.Name == method && m.GetParameters().Length == args.Length).Invoke(instance, args); }
            catch (TargetInvocationException e) { throw e.InnerException; }
        }
        private static void Check(bool value, string message) { if (!value) throw new InvalidOperationException(message); }
        private static void Reject(Action action, string code)
        { try { action(); } catch (InvalidOperationException e) { Check(e.Message.Contains(code), e.Message); return; } throw new Exception("Expected " + code); }
        internal static void GenerationRecovery()
        {
            foreach (var mode in new[] { "resume", "user_edit", "uncertain" })
            {
                var root = Path.Combine(Path.GetTempPath(), "scribble-ppt-recovery-" + Guid.NewGuid().ToString("N"));
                var store = new TaskCheckpointStore(root); var events = new List<string>();
                dynamic app = new CrossAppFixture("powerpoint", events);
                var request = new ChatCompletionRequest { messages = new List<object>(), tools = new List<ChatToolDefinition> { PresentationToolCatalog.DraftDefinition() } };
                var task = new TaskContextManager(request, "powerpoint", "Create a deck", store);
                var json = new JavaScriptSerializer();
                var slides = new object[] { new { id = "a", title = "First", layout = "cover" }, new { id = "b", title = "Second", layout = "closing" } };
                var call = new ChatToolCall { id = "first", function = new ChatToolCallFunction { name = PresentationToolCatalog.AddDraftSlides, arguments = json.Serialize(new { plan = new[] { "a", "b" }, slides }) } };
                var journalType = Type("SamsungGenerationJournal"); var writer = Type("PresentationDraftWriter");
                var parsed = Invoke(writer, "ParseSlides", null, (object)json.Deserialize<object[]>(json.Serialize(slides)));
                task.BeforeTool(call, true);
                var journal = Activator.CreateInstance(journalType, BindingFlags.Instance | BindingFlags.NonPublic, null, new object[] { task, call }, null);
                var additions = 0;
                CrossAppFixture.BeforeNativeCall = name =>
                {
                    if (mode == "uncertain" && name.EndsWith(".Export")) throw new InvalidOperationException("injected interruption");
                    if (name.EndsWith(".Slides.Add") && ++additions == 2) throw new InvalidOperationException("injected interruption");
                };
                try { Reject(() => Invoke(writer, "AddDraftSlides", null, (object)app, parsed, null, true, null, null, journal), "injected interruption"); }
                finally { CrossAppFixture.BeforeNativeCall = null; }
                dynamic deck = app.Presentations[1];
                Check((int)deck.Slides.Count == 1, "The interruption fixture must retain one native slide.");
                if (mode == "user_edit") deck.Slides[1].Shapes[1].TextFrame.TextRange.Text = "A later user edit";
                var resumed = new TaskContextManager(new ChatCompletionRequest { messages = new List<object>(), tools = new List<ChatToolDefinition>() }, "powerpoint", "Create a deck", store, store.Load(task.State.Id));
                var retry = new ChatToolCall { id = "retry", function = call.function };
                resumed.BeforeTool(retry, true);
                var resumedJournal = Activator.CreateInstance(journalType, BindingFlags.Instance | BindingFlags.NonPublic, null, new object[] { resumed, retry }, null);
                if (mode == "resume")
                {
                    Invoke(writer, "AddDraftSlides", null, (object)app, parsed, null, false, null, (object)deck, resumedJournal);
                    Check((int)deck.Slides.Count == 2, "Resume duplicated surviving slides.");
                    Invoke(journalType, "Complete", resumedJournal);
                    Check(!resumed.State.HostData.ContainsKey("samsung_pending") && resumed.State.Writes.All(w => w.Status == "verified"), "Reconciled receipts did not clear uncertain writes.");
                }
                else
                {
                    Reject(() => Invoke(writer, "AddDraftSlides", null, (object)app, parsed, null, false, null, (object)deck, resumedJournal), mode == "user_edit" ? "SLIDE_RECOVERY_USER_EDIT" : "SLIDE_RECOVERY_UNCERTAIN");
                    Check((int)deck.Slides.Count == 1, "Uncertain resume wrote another slide.");
                }
                var changed = new ChatToolCall { id = "changed", function = new ChatToolCallFunction { name = call.function.name, arguments = call.function.arguments.Replace("First", "Different") } };
                if (mode != "resume") Reject(() => resumed.BeforeTool(changed, true), "uncertain");
            }
        }
        internal static void RevisionRecovery()
        {
            var events = new List<string>(); dynamic app = new CrossAppFixture("powerpoint", events);
            dynamic deck = app.Presentations.Add(0); deck.PageSetup.SlideWidth = 960; deck.PageSetup.SlideHeight = 540;
            dynamic slide = deck.Slides.Add(1, 12);
            dynamic backupDeck = app.Presentations.Add(0); backupDeck.PageSetup.SlideWidth = 960; backupDeck.PageSetup.SlideHeight = 540;
            backupDeck.Tags.Add("ScribbleRevisionRecovery", "batch"); dynamic backup = backupDeck.Slides.Add(1, 12);
            var fp = PresentationInspection.Fingerprint((object)slide);
            var content = Invoke(Type("PresentationInspection"), "ContentFingerprint", null, (object)slide);
            var json = new JavaScriptSerializer();
            var state = new { BatchId = "batch", Status = "applied", PresentationId = PresentationInspection.IdentityFor((object)deck),
                BeforeOrder = new[] { 1 }, AfterOrder = new[] { 1 }, LastOrder = new[] { 1 }, StructuralStarted = false,
                BeforeFingerprints = new Dictionary<string, string> { { "1", fp } }, UnrelatedContent = new Dictionary<string, string>(),
                Items = new[] { new { SlideId = 1, LiveId = 1, BackupId = 1, Index = 1, Before = fp, After = fp, LastKnown = content,
                    BackupFingerprint = PresentationInspection.Fingerprint((object)backup), Started = true, Applied = true, Deleted = false,
                    Operations = new object[0], InsertedIds = new int[0], AddedShapeIds = new int[0], InsertedFingerprints = new Dictionary<string, string>() } } };
            var transaction = Type("PresentationRevision"); var snapshot = json.Serialize(state);
            var recovered = Invoke(transaction, "Recover", null, (object)app, (object)deck, snapshot);
            var eventCount = events.Count;
            Check((string)Invoke(transaction, "Reconcile", recovered) == "applied", "A completed revision was not recognized.");
            Check(!events.Skip(eventCount).Any(e => e.Contains(".Add(") || e.Contains(".Delete(")), "A completed revision was replayed.");
            recovered = Invoke(transaction, "Recover", null, (object)app, (object)deck, snapshot.Replace("\"Status\":\"applied\"", "\"Status\":\"applying\""));
            Check((string)Invoke(transaction, "Reconcile", recovered) == "rolled_back", "An interrupted pre-mutation revision was not recovered.");
            recovered = Invoke(transaction, "Recover", null, (object)app, (object)deck, snapshot);
            slide.Shapes.AddTextbox(1, 10, 10, 100, 50).TextFrame.TextRange.Text = "User edit";
            Reject(() => Invoke(transaction, "Reconcile", recovered), "REVISION_RECOVERY_USER_EDIT");
            backup.Shapes.AddTextbox(1, 10, 10, 100, 50).TextFrame.TextRange.Text = "Changed backup";
            Reject(() => Invoke(transaction, "Recover", null, (object)app, (object)deck, snapshot), "REVISION_RECOVERY_ORIGINAL_CHANGED");
        }
        internal static void RepairScopeAndChartBindings()
        {
            var json = new JavaScriptSerializer(); var policy = Type("SamsungRepairPolicy");
            Func<string, object[]> parse = value => json.Deserialize<object[]>(value);
            var original = parse("[{\"kind\":\"replace_text\",\"slide_id\":42,\"shape_id\":9,\"fingerprint\":\"fixed\",\"before\":\"Sales 100 units\",\"text\":\"Sales increased to 125 units\"}]");
            var corrected = parse(json.Serialize(original).Replace("Sales increased to 125 units", "Sales: 125 units"));
            Invoke(policy, "ValidateScope", null, original, corrected);
            Reject(() => Invoke(policy, "ValidateScope", null, original, parse(json.Serialize(corrected).Replace("125", "120"))), "REVISION_REPAIR_VALUES");
            Reject(() => Invoke(policy, "ValidateScope", null, original, parse(json.Serialize(corrected).Replace("42", "43"))), "REVISION_REPAIR_SCOPE");
            Reject(() => Invoke(policy, "ValidateScope", null, original, new object[0]), "REVISION_REPAIR_SCOPE");
            var chart = Type("PresentationChartEdit");
            var range = Invoke(chart, "Resolve", null, "=SERIES(Sheet1!$B$1,Sheet1!$A$2:$A$4,Sheet1!$B$2:$B$4,1)", 2);
            Check((string)range.GetType().GetField("Cell", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(range) == "B3", "Chart point mapped to wrong cell.");
            range = Invoke(chart, "Resolve", null, "=SERIES(,,'O''Brien, Korea'!$B$5:$D$5,1)", 3);
            Check((string)range.GetType().GetField("Cell", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(range) == "D5", "Horizontal category mapped incorrectly.");
            foreach (var formula in new[] { "=SERIES(,,[external.xlsx]Sheet1!$B$2:$B$4,1)", "=SERIES(,,Sheet1!$B$2:$C$4,1)", "=SERIES(,,{100,120},1)" })
                Reject(() => Invoke(chart, "Resolve", null, formula, 1), "REVISION_CHART_SOURCE_UNSUPPORTED");
        }
        internal static void LegacyRenderer()
        {
            var json = new JavaScriptSerializer();
            var old = Type("LegacySamsung.SamsungPresentationReview");
            var slide = "[{\"title\":\"Legacy report\",\"subtitle\":\"Sales improved\",\"layout\":\"bullets\",\"bullets\":[\"Evidence retained\"]}]";
            var plan = Invoke(old, "InspectPlan", null, slide);
            Check(json.Serialize(plan).Contains("Legacy report"), "Frozen legacy renderer failed.");
            Check((string)Type("LegacySamsung.SamsungSlideDesign").GetField("Version").GetRawConstantValue() != SamsungSlideDesign.Version, "Legacy renderer points at v2 tokens.");
            Check(typeof(DocumentDraftHost).GetMethod("ExecuteLegacySamsungAsync", BindingFlags.Instance | BindingFlags.NonPublic) != null, "Legacy dispatch is missing.");
        }
    }
}
