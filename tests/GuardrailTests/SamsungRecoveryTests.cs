using System;
using System.Collections;
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
        internal static void FullyRolledBackWriteCanBeCorrected()
        {
            foreach (var mode in new[] { "rolled_back", "surviving_slide", "wrong_owner" })
            {
                var root = Path.Combine(Path.GetTempPath(), "scribble-ppt-rollback-" + Guid.NewGuid().ToString("N"));
                try
                {
                    var task = new TaskContextManager(new ChatCompletionRequest { messages = new List<object>(),
                        tools = new List<ChatToolDefinition> { PresentationToolCatalog.DraftDefinition() } },
                        "powerpoint", "Create a deck", new TaskCheckpointStore(root));
                    var call = new ChatToolCall { id = "first", function = new ChatToolCallFunction {
                        name = PresentationToolCatalog.AddDraftSlides, arguments = "{\"slides\":[{\"id\":\"a\",\"title\":\"First\"}]}" } };
                    task.BeforeTool(call, true);
                    var journalType = Type("SamsungGenerationJournal");
                    var journal = Activator.CreateInstance(journalType, BindingFlags.Instance | BindingFlags.NonPublic,
                        null, new object[] { task, call }, null);
                    dynamic app = new CrossAppFixture("powerpoint", new List<string>());
                    dynamic deck = app.Presentations.Add(0);
                    Invoke(journalType, "Bind", journal, (object)deck, 1, null);
                    Check(task.State.HostData.ContainsKey("samsung_pending"), "The write was not journaled before the native boundary.");
                    if (mode == "surviving_slide") deck.Slides.Add(1, 12);
                    if (mode == "wrong_owner") deck.Tags.Add("ScribbleTask", "another-task");
                    var released = (bool)Invoke(journalType, "ReleaseRolledBackWrite", journal);
                    Check(released == (mode == "rolled_back"), "An uncertain native write was incorrectly released or retained.");
                    Check(task.State.HostData.ContainsKey("samsung_pending") == !released,
                        "The write fence was not retained exactly when native state is uncertain.");
                }
                finally { if (Directory.Exists(root)) Directory.Delete(root, true); }
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
        internal static void GeometryRepairOwnershipBaseline()
        {
            const string geometry = "{\"approved\":false,\"findings\":[{" +
                "\"slide_id\":\"trend\",\"severity\":\"blocker\"," +
                "\"code\":\"CHART_HIGHLIGHT_BOUNDS\"," +
                "\"correction\":\"Move the highlight inside the chart\"}]}";
            var scoped = (string)Invoke(typeof(DocumentDraftHost), "ReviewFindingsForSlide", null,
                geometry, "trend");
            var directive = (string)Invoke(typeof(DocumentDraftHost), "SlideRepairDirective", null,
                "trend");
            Check(!SamsungAuthoringPolicy.Approved(scoped) &&
                scoped.Contains("CHART_HIGHLIGHT_BOUNDS") &&
                directive.Contains("complete corrected slide") &&
                directive.Contains("slide whose id is 'trend'"),
                "The geometry-review baseline changed; verify repair ownership before replacing this diagnostic.");
        }
        internal static void RepairScopeAndChartBindings()
        {
            var json = new JavaScriptSerializer(); var policy = Type("SamsungRepairPolicy");
            Check((bool)Invoke(Type("DocumentDraftHost"), "ShouldDraftRepairedDeck", null,
                "powerpoint", "Create a repaired, editable draft of every slide. Retain exactly 6 output slides. Preserve source slides.", 6),
                "A full-deck repair would append six drafts to the six source slides instead of creating a separate six-slide output.");
            Check(!(bool)Invoke(Type("DocumentDraftHost"), "ShouldDraftRepairedDeck", null,
                "powerpoint", "Add two draft slides after the current slide.", 2),
                "An ordinary in-deck slide request was redirected to a new presentation.");
            Check(SamsungSlideDesign.Takeaway.Width > 700f,
                "The takeaway band is too narrow for readable executive copy.");
            Check(SamsungAuthoringPolicy.AudienceTakeaway(
                "Single primary series on a zero-based axis: June revenue fell to 82,992 EUR.") ==
                "June revenue fell to 82,992 EUR.",
                "A chart-construction instruction leaked into the visible takeaway.");
            Check(SamsungAuthoringPolicy.AudienceTakeaway(
                "A single primary revenue series on a zero-based axis shows the 2,527 EUR dip.") ==
                "The 2,527 EUR dip.",
                "A measure-specific chart-construction clause leaked into the visible takeaway.");
            Check(SamsungAuthoringPolicy.AudienceUnit("not applicable") == "" &&
                SamsungAuthoringPolicy.AudienceUnit("N/A") == "" &&
                SamsungAuthoringPolicy.AudienceUnit(" EUR ") == "EUR",
                "A placeholder unit reached the visible slide or a real unit was removed.");
            var visualDraft = Invoke(Type("PresentationDraftWriter"), "ParseSlides", null,
                (object)json.Deserialize<object[]>("[{\"id\":\"headline\",\"layout\":\"scorecard\",\"title\":\"Results\",\"subtitle\":\"June revenue declined\",\"takeaway\":\"June revenue fell to 82,992 EUR.\",\"cards\":[{\"heading\":\"Revenue\",\"points\":[\"82,992\"]},{\"heading\":\"Cost\",\"points\":[\"36,714\"]}]}]"));
            var visualPage = ((IEnumerable)Invoke(Type("PresentationDraftWriter"), "ComposeSamsung", null, visualDraft))
                .Cast<object>().Single();
            var elements = ((IEnumerable)visualPage.GetType().GetField("Elements", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(visualPage))
                .Cast<object>();
            var takeawayElement = elements.Single(element => (string)element.GetType().GetField("Text", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(element) == "June revenue fell to 82,992 EUR.");
            Check((float)takeawayElement.GetType().GetField("Minimum", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(takeawayElement) >= 14f,
                "A factual takeaway was allowed to shrink below the native minimum body font.");
            var takeawayBox = (System.Drawing.RectangleF)takeawayElement.GetType().GetField("Box", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(takeawayElement);
            Check(Math.Abs(takeawayBox.Y - SamsungSlideDesign.ScorecardTakeaway.Y) < .1f &&
                SamsungSlideDesign.ScorecardTakeaway.Bottom < SamsungSlideDesign.Footer.Y,
                "A scorecard takeaway left a large empty gap or collided with the source footer.");
            var cardDraft = Invoke(Type("PresentationDraftWriter"), "ParseSlides", null,
                (object)json.Deserialize<object[]>("[{\"id\":\"quality\",\"layout\":\"cards\",\"title\":\"Data quality\",\"subtitle\":\"Complete observations\",\"cards\":[{\"heading\":\"Integrity\",\"points\":[\"Ledger A2:L145\",\"144 records verified\",\"No blanks\"]},{\"heading\":\"Coverage\",\"points\":[\"May 2026: 24 records\",\"June 2026: 24 records\"]}]}]"));
            var cardPage = ((IEnumerable)Invoke(Type("PresentationDraftWriter"), "ComposeSamsung", null, cardDraft))
                .Cast<object>().Single();
            var cardElements = ((IEnumerable)cardPage.GetType().GetField("Elements", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(cardPage))
                .Cast<object>().ToArray();
            var largeText = cardElements.Where(element =>
                (float)element.GetType().GetField("Size", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(element) >= 26f)
                .Select(element => (string)element.GetType().GetField("Text", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(element))
                .ToArray();
            Check(largeText.Contains("144") && largeText.Contains("24") && !largeText.Contains("2026") && !largeText.Contains("45,"),
                "A reporting year or fragment of a cell range was promoted as a hero metric.");
            var compactGridDraft = Invoke(Type("PresentationDraftWriter"), "ParseSlides", null,
                (object)json.Deserialize<object[]>("[{\"id\":\"data-quality\",\"layout\":\"cards\",\"title\":\"Data quality\",\"subtitle\":\"Complete figures\",\"cards\":[{\"heading\":\"Source coverage\",\"points\":[\"Workbook WB01 (Ledger)\",\"Jan–Jun 2026\",\"North / South / East / West\",\"Revenue & Cost EUR\"]},{\"heading\":\"Completeness\",\"points\":[\"24 obs in June\",\"0 blank Revenue\",\"0 blank Cost\",\"June measures complete\"]},{\"heading\":\"Method\",\"points\":[\"Additive SUMIF sums\",\"Blanks = unknown\",\"Aggregate-total rates\",\"No row-avg %\"]},{\"heading\":\"Evidence boundary\",\"points\":[\"Group compare: June only\",\"Blanks remain unknown\",\"Planned ≠ completed\",\"Financial ≠ operational\"]}]}]"));
            var compactGridPage = ((IEnumerable)Invoke(Type("PresentationDraftWriter"), "ComposeSamsung", null, compactGridDraft))
                .Cast<object>().Single();
            var compactGridElements = ((IEnumerable)compactGridPage.GetType().GetField("Elements", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(compactGridPage))
                .Cast<object>().ToArray();
            Check(compactGridElements.Any(element => ((string)element.GetType().GetField("Text", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(element)).Contains("Revenue & Cost EUR")),
                "A four-card source panel was dropped or overflowed instead of fitting at the native minimum size.");
            var numericGridDraft = Invoke(Type("PresentationDraftWriter"), "ParseSlides", null,
                (object)json.Deserialize<object[]>("[{\"id\":\"quality\",\"layout\":\"cards\",\"title\":\"Data-quality limits\",\"subtitle\":\"Ledger is complete\",\"cards\":[{\"heading\":\"Completeness\",\"points\":[\"144 unique RowIDs\",\"0 blank cells\",\"24 rows per month\"]},{\"heading\":\"Rate and method\",\"points\":[\"Rates use aggregate totals\",\"Never average row percentages\"]},{\"heading\":\"Incomplete observation\",\"points\":[\"History through May 2026\",\"June sourced from Ledger\"]},{\"heading\":\"Cross-check\",\"points\":[\"May reconciles Ledger and History\",\"Revenue EUR 85,519\",\"Cost EUR 36,702\"]}]}]"));
            var numericGridPage = ((IEnumerable)Invoke(Type("PresentationDraftWriter"), "ComposeSamsung", null, numericGridDraft))
                .Cast<object>().Single();
            var numericGridHeroes = ((IEnumerable)numericGridPage.GetType().GetField("Elements", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(numericGridPage))
                .Cast<object>().Where(element =>
                    (float)element.GetType().GetField("Size", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(element) >= 26f)
                .Select(element => (string)element.GetType().GetField("Text", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(element))
                .ToArray();
            Check(numericGridHeroes.Contains("144") && numericGridHeroes.Contains("85,519") && !numericGridHeroes.Contains("2026"),
                "A four-card numeric evidence slide lacked two relevant metric anchors or promoted a year as a metric.");
            var concentratedDraft = Invoke(Type("PresentationDraftWriter"), "ParseSlides", null,
                (object)json.Deserialize<object[]>("[{\"id\":\"quality\",\"layout\":\"cards\",\"title\":\"Data-quality limits and method\",\"subtitle\":\"144 records counted once\",\"cards\":[{\"heading\":\"Coverage\",\"points\":[\"Every RowID counted once — WB01-0001 to WB01-0144\",\"24 observations per month: 4 groups × 6 products\",\"May 2026 and June 2026 both fully populated\"]},{\"heading\":\"Integrity\",\"points\":[\"No blank or non-numeric cells\",\"Original worksheets preserved\"]},{\"heading\":\"Methodology\",\"points\":[\"Rates at aggregate level\",\"Full precision in formulas\"]}]}]"));
            var concentratedPage = ((IEnumerable)Invoke(Type("PresentationDraftWriter"), "ComposeSamsung", null, concentratedDraft))
                .Cast<object>().Single();
            var concentratedHeroes = ((IEnumerable)concentratedPage.GetType().GetField("Elements", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(concentratedPage))
                .Cast<object>().Where(element =>
                    (float)element.GetType().GetField("Size", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(element) >= 26f)
                .Select(element => (string)element.GetType().GetField("Text", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(element))
                .ToArray();
            Check(concentratedHeroes.Contains("24") && concentratedHeroes.Contains("4") &&
                !concentratedHeroes.Contains("0144") && !concentratedHeroes.Contains("2026"),
                "A single quantified card lost its distinct metric anchors or promoted an identifier/year.");
            var dualDraft = Invoke(Type("PresentationDraftWriter"), "ParseSlides", null,
                (object)json.Deserialize<object[]>("[{\"id\":\"bounds\",\"layout\":\"cards\",\"title\":\"Evidence bounds\",\"cards\":[{\"heading\":\"June figures\",\"points\":[\"Revenue EUR 82,992; Cost EUR 36,714\",\"24 ledger rows\"]},{\"heading\":\"Period scope\",\"points\":[\"June from Ledger\"]},{\"heading\":\"Method\",\"points\":[\"Blanks are unknown\"]}]}]"));
            var dualPage = ((IEnumerable)Invoke(Type("PresentationDraftWriter"), "ComposeSamsung", null, dualDraft))
                .Cast<object>().Single();
            var dualText = ((IEnumerable)dualPage.GetType().GetField("Elements", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(dualPage))
                .Cast<object>().Select(element => (string)element.GetType().GetField("Text", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(element)).ToArray();
            Check(dualText.Contains("82,992") && dualText.Contains("36,714") && dualText.Contains("Revenue EUR") &&
                dualText.Contains("Cost EUR") && dualText.Contains("24 ledger rows") &&
                !dualText.Any(value => value.Contains("Revenue EUR 82,992") || value.Contains("Cost EUR 36,714")),
                "The two hero metrics repeated their source lines instead of retaining labels and one readable copy of each number.");
            var highlightedDraft = Invoke(Type("PresentationDraftWriter"), "ParseSlides", null,
                (object)json.Deserialize<object[]>("[{\"id\":\"trend\",\"layout\":\"chart\",\"title\":\"Trend\",\"highlight_rows\":[1],\"chart\":{\"type\":\"column\",\"title\":\"Revenue EUR\",\"categories\":[\"2026-05\",\"2026-06\"],\"series\":[{\"name\":\"Revenue EUR\",\"values\":[85519,82992]}]}}]"));
            var highlightedPage = ((IEnumerable)Invoke(Type("PresentationDraftWriter"), "ComposeSamsung", null, highlightedDraft))
                .Cast<object>().Single();
            var highlights = ((IEnumerable)highlightedPage.GetType().GetField("Elements", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(highlightedPage))
                .Cast<object>().Where(element => element.GetType().GetField("HighlightChart", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(element) != null).ToArray();
            Check(highlights.Length == 1 && (int)highlights[0].GetType().GetField("HighlightSeries", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(highlights[0]) == 1,
                "Chart annotation surrounded an empty full-height category slot rather than the source bar.");
            Check((bool)Invoke(Type("PresentationDraftWriter"), "RetryableSamsungChartFailure", null,
                "write chart data: COMException 0x800A01A8 Exception from HRESULT: 0x800A01A8"),
                "A transient embedded Excel chart-grid failure should be retried inside the host call.");
            Check(!(bool)Invoke(Type("PresentationDraftWriter"), "RetryableSamsungChartFailure", null,
                "series readback: InvalidOperationException Chart values differ"),
                "A factual native chart mismatch must not be retried as a transient COM failure.");
            Func<string, object[]> parse = value => json.Deserialize<object[]>(value);
            var original = parse("[{\"kind\":\"replace_text\",\"slide_id\":42,\"shape_id\":9,\"fingerprint\":\"fixed\",\"before\":\"Sales 100 units\",\"text\":\"Sales increased to 125 units\"}]");
            var corrected = parse(json.Serialize(original).Replace("Sales increased to 125 units", "Sales: 125 units"));
            Invoke(policy, "ValidateScope", null, original, corrected);
            Reject(() => Invoke(policy, "ValidateScope", null, original, parse(json.Serialize(corrected).Replace("125", "120"))), "REVISION_REPAIR_VALUES");
            Reject(() => Invoke(policy, "ValidateScope", null, original, parse(json.Serialize(corrected).Replace("42", "43"))), "REVISION_REPAIR_SCOPE");
            Reject(() => Invoke(policy, "ValidateScope", null, original, new object[0]), "REVISION_REPAIR_SCOPE");
            var originalSlide = json.Deserialize<Dictionary<string, object>>("{\"id\":\"trend\",\"chart\":{\"type\":\"column\",\"title\":\"Revenue by month\",\"categories\":[\"2026-05\",\"2026-06\"],\"series\":[{\"name\":\"Revenue EUR\",\"values\":[85519,82992]}]}}");
            var retitledSlide = json.Deserialize<Dictionary<string, object>>("{\"id\":\"trend\",\"chart\":{\"type\":\"column\",\"title\":\"Revenue by month, EUR\",\"categories\":[\"2026-05\",\"2026-06\"],\"series\":[{\"name\":\"Revenue EUR\",\"values\":[85519,82992]}]}}");
            Invoke(typeof(DocumentDraftHost), "ValidateSlideRepairEvidence", null, originalSlide, retitledSlide);
            var changedData = json.Deserialize<Dictionary<string, object>>(json.Serialize(retitledSlide).Replace("82992", "82993"));
            Reject(() => Invoke(typeof(DocumentDraftHost), "ValidateSlideRepairEvidence", null, originalSlide, changedData), "SLIDE_REPAIR_EVIDENCE_CHANGED: chart");
            var sourcedSlide = json.Deserialize<Dictionary<string, object>>(
                "{\"id\":\"quality\",\"title\":\"Old title\",\"sources\":\"WB01 Ledger & Scribble Draft\",\"evidence\":\"verified cells\",\"source_spans\":[\"span-1\"]}");
            var visualReplacement = json.Deserialize<Dictionary<string, object>>(
                "{\"id\":\"quality\",\"title\":\"Clearer title\",\"sources\":\"\"}");
            Invoke(typeof(DocumentDraftHost), "RetainSlideRepairSources", null, sourcedSlide, visualReplacement);
            Check(SamsungAuthoringPolicy.Text(visualReplacement, "sources") == "WB01 Ledger & Scribble Draft" &&
                SamsungAuthoringPolicy.Text(visualReplacement, "evidence") == "verified cells" &&
                visualReplacement.ContainsKey("source_spans"),
                "A visual-only repair discarded the original visible citation or resolved evidence.");
            Invoke(typeof(DocumentDraftHost), "ValidateSlideRepairEvidence", null, sourcedSlide, visualReplacement);
            visualReplacement["sources"] = "Unrelated source";
            Reject(() => Invoke(typeof(DocumentDraftHost), "ValidateSlideRepairEvidence", null, sourcedSlide, visualReplacement),
                "SLIDE_REPAIR_EVIDENCE_CHANGED: sources");
            var repairWithExtra = json.Deserialize<Dictionary<string, object>>(
                "{\"slides\":[{\"id\":\"headline\",\"layout\":\"scorecard\"},{\"id\":\"period-comparison\",\"layout\":\"chart\"}]}");
            var selectedRepair = (object[])Invoke(typeof(DocumentDraftHost), "SelectSlideRepair", null,
                repairWithExtra, "headline");
            Check(selectedRepair.Length == 1 && SamsungAuthoringPolicy.Text((Dictionary<string, object>)selectedRepair[0], "id") == "headline",
                "A model-supplied extra planned slide was not excluded from single-slide repair.");
            Reject(() => Invoke(typeof(DocumentDraftHost), "SelectSlideRepair", null, repairWithExtra, "missing"),
                "SLIDE_REPAIR_COUNT");
            var ambiguousRepair = json.Deserialize<Dictionary<string, object>>(
                "{\"slides\":[{\"id\":\"headline\"},{\"id\":\"headline\"}]}");
            Reject(() => Invoke(typeof(DocumentDraftHost), "SelectSlideRepair", null, ambiguousRepair, "headline"),
                "SLIDE_REPAIR_COUNT");
            var visualReport = "{\"approved\":false,\"issues\":\"marker and real overflow\",\"findings\":[" +
                "{\"slide_id\":\"headline\",\"severity\":\"blocker\",\"correction\":\"Remove host draft marker\"}," +
                "{\"slide_id\":\"headline\",\"severity\":\"blocker\",\"correction\":\"Fix actual table overlap\"}]}";
            Func<IDictionary<string, object>, bool> refuteMarker = finding =>
                SamsungAuthoringPolicy.Text(finding, "correction").Contains("draft marker");
            var filtered = (string)Invoke(typeof(DocumentDraftHost), "FilterReviewFindings", null,
                visualReport, refuteMarker);
            var filteredMap = json.Deserialize<Dictionary<string, object>>(filtered);
            Check(!Convert.ToBoolean(filteredMap["approved"]) &&
                SamsungAuthoringPolicy.Array(filteredMap, "findings").Length == 1 &&
                filtered.Contains("actual table overlap") && !filtered.Contains("Remove host draft marker"),
                "A genuine visual defect was lost while filtering a native-refuted finding.");
            var allRefuted = (string)Invoke(typeof(DocumentDraftHost), "FilterReviewFindings", null,
                visualReport, new Func<IDictionary<string, object>, bool>(finding => true));
            Check(Convert.ToBoolean(json.Deserialize<Dictionary<string, object>>(allRefuted)["approved"]),
                "Fully native-refuted findings still blocked an unchanged slide.");
            var multiSlideReview = "{\"approved\":false,\"issues\":\"two slides\",\"findings\":[" +
                "{\"slide_id\":\"review\",\"severity\":\"blocker\",\"correction\":\"Remove repeated KPI copy\"}," +
                "{\"slide_id\":\"coverage\",\"severity\":\"blocker\",\"correction\":\"Tighten the footer\"}]}";
            var reviewOnly = (string)Invoke(typeof(DocumentDraftHost), "ReviewFindingsForSlide", null,
                multiSlideReview, "review");
            var reviewOnlyMap = json.Deserialize<Dictionary<string, object>>(reviewOnly);
            Check(SamsungAuthoringPolicy.Array(reviewOnlyMap, "findings").Length == 1 &&
                reviewOnly.Contains("Remove repeated KPI copy") && !reviewOnly.Contains("Tighten the footer"),
                "A single-slide repair received another slide's visual findings.");
            var briefConflict = "{\"approved\":false,\"issues\":\"Gross margin is 55.76%, but the brief explicitly requires 55.74%.\",\"findings\":[" +
                "{\"slide_id\":\"evidence\",\"severity\":\"blocker\",\"type\":\"facts\",\"correction\":\"Change gross margin from 55.76% to 55.74% as specified in the brief.\"}," +
                "{\"slide_id\":\"evidence\",\"severity\":\"warning\",\"type\":\"facts\",\"correction\":\"Keep the period caveat visible.\"}]}";
            var grounded = (string)Invoke(typeof(DocumentDraftHost), "FilterBriefRefutedReview", null,
                briefConflict, "{\"required_content\":[\"Gross margin 55.74%\"]}",
                "{\"calculations\":[{\"result\":55.76}],\"subtitle\":\"Gross margin 55.76%\"}");
            var groundedMap = json.Deserialize<Dictionary<string, object>>(grounded);
            Check(Convert.ToBoolean(groundedMap["approved"]) &&
                SamsungAuthoringPolicy.Array(groundedMap, "findings").Length == 1 &&
                grounded.Contains("period caveat") && !grounded.Contains("Change gross margin"),
                "A stale numeric brief overrode host-verified slide arithmetic or erased an unrelated warning.");
            var writerType = Type("PresentationDraftWriter");
            var outputType = writerType.GetNestedType("SamsungOutput", BindingFlags.NonPublic);
            var pageType = writerType.GetNestedType("SamsungPage", BindingFlags.NonPublic);
            var elementType = writerType.GetNestedType("SamsungElement", BindingFlags.NonPublic);
            var output = Activator.CreateInstance(outputType, true);
            var page = Activator.CreateInstance(pageType, true);
            var folio = Activator.CreateInstance(elementType, true);
            elementType.GetField("Text", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(folio, "- 6 -");
            pageType.GetField("PageNumber", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(page, folio);
            dynamic nativeSlide = new CrossAppFixture("native-slide", new List<string>());
            nativeSlide.SlideIndex = 6;
            dynamic blankShape = nativeSlide.Shapes.AddShape();
            blankShape.TextFrame.TextRange.Text = null;
            dynamic folioShape = nativeSlide.Shapes.AddShape();
            folioShape.TextFrame.TextRange.Text = "- 6 -";
            outputType.GetField("Slide", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(output, nativeSlide);
            outputType.GetField("Page", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(output, page);
            Check((bool)Invoke(typeof(DocumentDraftHost), "NativePageNumberMatches", null, output),
                "An empty native textbox before the correct folio hid the host-owned page number.");
            Check((bool)Invoke(typeof(DocumentDraftHost), "CanRetrySlideRepairShape", null,
                new InvalidOperationException("SLIDE_REPAIR_COUNT: ambiguous target"), 0) &&
                (bool)Invoke(typeof(DocumentDraftHost), "CanRetrySlideRepairShape", null,
                new InvalidOperationException("SLIDE_REPAIR_ID_CHANGED"), 0) &&
                (bool)Invoke(typeof(DocumentDraftHost), "CanRetrySlideRepairShape", null,
                new InvalidOperationException("SLIDE_REPAIR_COUNT: ambiguous target"), 1) &&
                !(bool)Invoke(typeof(DocumentDraftHost), "CanRetrySlideRepairShape", null,
                new InvalidOperationException("SLIDE_REPAIR_COUNT: ambiguous target"), 2) &&
                !(bool)Invoke(typeof(DocumentDraftHost), "CanRetrySlideRepairShape", null,
                new InvalidOperationException("SLIDE_REPAIR_EVIDENCE_CHANGED: chart"), 0),
                "A malformed visual repair was not offered two bounded, safe pre-mutation corrections.");
            Check((string)Invoke(Type("PresentationDraftWriter"), "StripHeroClauses", null,
                    "Revenue EUR 82,992 across 24 ledger records; Cost EUR 36,714", "82,992", "36,714") ==
                "24 ledger records",
                "Evidence-card KPI callouts repeated the same metrics in body copy.");
            var coverHeadline = new[] { json.Deserialize<Dictionary<string, object>>(
                "{\"id\":\"headline\",\"layout\":\"cover\",\"title\":\"June results\"}") };
            Reject(() => SamsungAuthoringPolicy.ValidateRequestedHeadlineLayout(
                "Create four slides: headline, period chart, groups and data quality.", coverHeadline),
                "SLIDE_HEADLINE_NOT_COVER");
            SamsungAuthoringPolicy.ValidateRequestedHeadlineLayout(
                "Create a cover and then a headline slide.", coverHeadline);
            var primaryOnlySlides = Invoke(Type("PresentationDraftWriter"), "ParseSlides", null, (object)json.Deserialize<object[]>(
                "[{\"id\":\"trend\",\"title\":\"Trend\",\"layout\":\"chart\",\"chart\":{\"type\":\"column\",\"categories\":[\"2026-05\",\"2026-06\"],\"series\":[{\"name\":\"Revenue EUR\",\"values\":[85519,82992]},{\"name\":\"Cost EUR\",\"values\":[36702,36714]}]}}]"));
            Reject(() => Invoke(typeof(DocumentDraftHost), "ValidatePromptChartConstraints", null,
                "The chart must use only primary values for May and June.", primaryOnlySlides), "SLIDE_PRIMARY_SERIES_ONLY");
            var splitCharts = Invoke(Type("PresentationDraftWriter"), "ParseSlides", null, (object)json.Deserialize<object[]>(
                "[{\"id\":\"trend\",\"title\":\"Trend\",\"layout\":\"annotated_chart\",\"chart\":{\"type\":\"column\",\"categories\":[\"2026-05\",\"2026-06\"],\"series\":[{\"name\":\"Revenue EUR\",\"values\":[85519,82992]}]},\"secondary_chart\":{\"type\":\"column\",\"categories\":[\"2026-05\",\"2026-06\"],\"series\":[{\"name\":\"Cost EUR\",\"values\":[36702,36714]}]}}]"));
            Reject(() => Invoke(typeof(DocumentDraftHost), "ValidatePromptChartConstraints", null,
                "The chart must use only primary values for May and June.", splitCharts), "SLIDE_PRIMARY_SERIES_ONLY");
            Invoke(typeof(DocumentDraftHost), "ValidatePromptChartConstraints", null,
                "Compare primary and secondary values in the chart.", primaryOnlySlides);
            var compliantChartSlides = Invoke(Type("PresentationDraftWriter"), "ParseSlides", null, (object)json.Deserialize<object[]>(
                "[{\"id\":\"trend\",\"title\":\"Trend\",\"layout\":\"chart\",\"chart\":{\"type\":\"column\",\"title\":\"Revenue EUR — May vs June\",\"categories\":[\"2026-05\",\"2026-06\"],\"series\":[{\"name\":\"Revenue EUR\",\"values\":[85519,82992]}]}}]"));
            Invoke(typeof(DocumentDraftHost), "ValidatePromptChartConstraints", null,
                "Use YYYY-MM categories and EUR in the title.", compliantChartSlides);
            var missingTitleUnit = Invoke(Type("PresentationDraftWriter"), "ParseSlides", null, (object)json.Deserialize<object[]>(
                "[{\"id\":\"trend\",\"title\":\"Trend\",\"layout\":\"chart\",\"chart\":{\"type\":\"column\",\"title\":\"Revenue — May vs June\",\"categories\":[\"2026-05\",\"2026-06\"],\"series\":[{\"name\":\"Revenue EUR\",\"values\":[85519,82992]}]}}]"));
            Reject(() => Invoke(typeof(DocumentDraftHost), "ValidatePromptChartConstraints", null,
                "Use YYYY-MM categories and EUR in the title.", missingTitleUnit), "SLIDE_CHART_TITLE_UNIT");
            var invalidCategories = Invoke(Type("PresentationDraftWriter"), "ParseSlides", null, (object)json.Deserialize<object[]>(
                "[{\"id\":\"trend\",\"title\":\"Trend\",\"layout\":\"chart\",\"chart\":{\"type\":\"column\",\"title\":\"Revenue EUR\",\"categories\":[\"May\",\"June\"],\"series\":[{\"name\":\"Revenue EUR\",\"values\":[85519,82992]}]}}]"));
            Reject(() => Invoke(typeof(DocumentDraftHost), "ValidatePromptChartConstraints", null,
                "Use YYYY-MM categories and EUR in the title.", invalidCategories), "SLIDE_CHART_CATEGORY_FORMAT");
            var groupCategories = Invoke(Type("PresentationDraftWriter"), "ParseSlides", null, (object)json.Deserialize<object[]>(
                "[{\"id\":\"groups\",\"title\":\"Group performance\",\"layout\":\"chart\",\"chart\":{\"type\":\"bar\",\"title\":\"Revenue EUR by group\",\"categories\":[\"North\",\"South\",\"East\",\"West\"],\"series\":[{\"name\":\"Revenue EUR\",\"values\":[19219,22675,19054,22044]}]}}]"));
            Invoke(typeof(DocumentDraftHost), "ValidatePromptChartConstraints", null,
                "The period chart must use only primary values with YYYY-MM categories and EUR in the title.", groupCategories);
            var mixedCategories = Invoke(Type("PresentationDraftWriter"), "ParseSlides", null, (object)json.Deserialize<object[]>(
                "[{\"id\":\"trend\",\"title\":\"Trend\",\"layout\":\"chart\",\"chart\":{\"type\":\"column\",\"title\":\"Revenue EUR\",\"categories\":[\"May\",\"2026-06\"],\"series\":[{\"name\":\"Revenue EUR\",\"values\":[85519,82992]}]}}]"));
            Reject(() => Invoke(typeof(DocumentDraftHost), "ValidatePromptChartConstraints", null,
                "Use YYYY-MM categories and EUR in the title.", mixedCategories), "SLIDE_CHART_CATEGORY_FORMAT");
            var fiveMonthChart = Invoke(Type("PresentationDraftWriter"), "ParseSlides", null, (object)json.Deserialize<object[]>(
                "[{\"id\":\"period\",\"title\":\"History\",\"layout\":\"chart\",\"chart\":{\"type\":\"column\",\"title\":\"Revenue and cost EUR\",\"categories\":[\"2026-01\",\"2026-02\",\"2026-03\",\"2026-04\",\"2026-05\"],\"series\":[{\"name\":\"Revenue EUR\",\"values\":[1,2,3,4,5]},{\"name\":\"Cost EUR\",\"values\":[1,2,3,4,5]}]}}]"));
            var sixMonthInstruction = "Recreate the monthly chart with exactly two series, all six YYYY-MM categories and EUR in its title.";
            Reject(() => Invoke(typeof(DocumentDraftHost), "ValidatePromptChartConstraints", null,
                sixMonthInstruction, fiveMonthChart), "SLIDE_CHART_PERIOD_COVERAGE");
            var sixMonthChart = Invoke(Type("PresentationDraftWriter"), "ParseSlides", null, (object)json.Deserialize<object[]>(
                "[{\"id\":\"period\",\"title\":\"History\",\"layout\":\"chart\",\"chart\":{\"type\":\"column\",\"title\":\"Revenue and cost EUR\",\"categories\":[\"2026-01\",\"2026-02\",\"2026-03\",\"2026-04\",\"2026-05\",\"2026-06\"],\"series\":[{\"name\":\"Revenue EUR\",\"values\":[1,2,3,4,5,6]},{\"name\":\"Cost EUR\",\"values\":[1,2,3,4,5,6]}]}}]"));
            Invoke(typeof(DocumentDraftHost), "ValidatePromptChartConstraints", null,
                sixMonthInstruction, sixMonthChart);
            var calculationFinding = json.Serialize(new { approved = false, findings = new[] {
                new { slide_id = "headline", object_id = "calculation:Gross margin % (May)", type = "facts",
                    correction = "Correct May margin from 57.08% to 57.26%." } } });
            Check((bool)Invoke(typeof(DocumentDraftHost), "OutlineReviewApprovedOrDeterministicallySatisfied", null,
                calculationFinding, "June headline", visualDraft),
                "A probabilistic outline arithmetic claim blocked a value that the exact source reviewer will calculate.");
            var satisfiedOutlineSlides = Invoke(Type("PresentationDraftWriter"), "ParseSlides", null,
                (object)json.Deserialize<object[]>("[{\"id\":\"period-trend\",\"title\":\"Revenue trend\",\"layout\":\"chart\",\"chart\":{\"type\":\"column\",\"title\":\"Revenue EUR — May vs June (zero-based)\",\"categories\":[\"2026-05\",\"2026-06\"],\"series\":[{\"name\":\"Revenue EUR\",\"values\":[85519,82992]}]}},{\"id\":\"june-groups\",\"title\":\"Groups\",\"layout\":\"table\",\"table\":{\"headers\":[\"Group\",\"Revenue\",\"Cost\"],\"rows\":[[\"North\",\"19,219\",\"8,082\"],[\"All groups\",\"82,992\",\"36,714\"]]}}]"));
            var satisfiedOutlineReview = json.Serialize(new { approved = false, findings = new object[] {
                new { slide_id = "period-trend", object_id = "chart.title", type = "facts",
                    correction = "Update chart title to 'Revenue EUR (May vs June 2026, zero-based)'." },
                new { slide_id = "june-groups", object_id = "table.rows", type = "coverage",
                    correction = "Add the 'All groups' row with values 82,992 and 36,714." } } });
            Check((bool)Invoke(typeof(DocumentDraftHost), "OutlineReviewApprovedOrDeterministicallySatisfied", null,
                satisfiedOutlineReview, "Use EUR in the chart title.", satisfiedOutlineSlides),
                "A reviewer missed an already-present chart title or source total row.");
            var falseRowReview = satisfiedOutlineReview.Replace("82,992", "82,993");
            Check(!(bool)Invoke(typeof(DocumentDraftHost), "OutlineReviewApprovedOrDeterministicallySatisfied", null,
                falseRowReview, "Use EUR in the chart title.", satisfiedOutlineSlides),
                "An incorrect source-total value was silently treated as a satisfied outline finding.");
            var falseTitleFinding = "{\"approved\":false,\"issues\":\"Chart title lacks EUR.\",\"findings\":[{\"slide_id\":\"trend\",\"object_id\":\"chart\",\"severity\":\"blocker\",\"type\":\"facts\",\"correction\":\"Include EUR explicitly in the chart title.\"}]}";
            Check((bool)Invoke(typeof(DocumentDraftHost), "ReviewApprovedOrSatisfiedPromptConstraint", null,
                falseTitleFinding, "Use EUR in the title.", ((IEnumerable)compliantChartSlides).Cast<object>().Single()),
                "A reviewer hallucination contradicted a host-verified literal chart-title token.");
            Check(!(bool)Invoke(typeof(DocumentDraftHost), "ReviewApprovedOrSatisfiedPromptConstraint", null,
                falseTitleFinding, "Use EUR in the title.", ((IEnumerable)missingTitleUnit).Cast<object>().Single()),
                "A genuinely missing chart-title token was ignored.");
            var mixedFinding = falseTitleFinding.Replace("]}", ",{\"slide_id\":\"trend\",\"object_id\":\"subtitle\",\"severity\":\"blocker\",\"type\":\"facts\",\"correction\":\"Remove the unsupported conclusion.\"}]}");
            Check(!(bool)Invoke(typeof(DocumentDraftHost), "ReviewApprovedOrSatisfiedPromptConstraint", null,
                mixedFinding, "Use EUR in the title.", ((IEnumerable)compliantChartSlides).Cast<object>().Single()),
                "A real factual blocker was hidden beside a satisfied title constraint.");
            var falseCostFinding = "{\"approved\":false,\"issues\":\"Cost series missing from chart.\",\"findings\":[{\"slide_id\":\"trend\",\"object_id\":\"chart\",\"severity\":\"blocker\",\"type\":\"coverage\",\"correction\":\"Add a second Cost EUR series to the native chart.\"}]}";
            Check((bool)Invoke(typeof(DocumentDraftHost), "ReviewApprovedOrSatisfiedPromptConstraint", null,
                falseCostFinding, "The chart must use only primary values for May and June.",
                ((IEnumerable)compliantChartSlides).Cast<object>().Single()),
                "A reviewer demanded the exact secondary chart series forbidden by the user.");
            Check(!(bool)Invoke(typeof(DocumentDraftHost), "ReviewApprovedOrSatisfiedPromptConstraint", null,
                falseCostFinding, "Compare revenue and cost in the chart.",
                ((IEnumerable)compliantChartSlides).Cast<object>().Single()),
                "A genuinely requested second series was ignored.");
            var cleanCoverSlides = Invoke(Type("PresentationDraftWriter"), "ParseSlides", null,
                (object)json.Deserialize<object[]>("[{\"id\":\"cover\",\"title\":\"Atlas Components: sales review\",\"subtitle\":\"January-June 2026, repaired June facts\",\"layout\":\"cover\",\"sources\":\"WB01\",\"footnote\":\"Fictional source\"}]"));
            var cleanCover = ((IEnumerable)cleanCoverSlides).Cast<object>().Single();
            var hallucinatedCoverCallouts = "{\"approved\":false,\"issues\":\"Cover has forbidden KPI callouts.\",\"findings\":[" +
                "{\"slide_id\":\"cover\",\"object_id\":\"Rectangle 4\",\"severity\":\"blocker\",\"type\":\"coverage\",\"correction\":\"Remove Revenue EUR 82,992, Cost EUR 36,714, and Gross margin 55.76% data callouts from the cover.\"}]}";
            Check((bool)Invoke(typeof(DocumentDraftHost), "ReviewApprovedOrSatisfiedPromptConstraint", null,
                hallucinatedCoverCallouts, "Repair all six slides.", cleanCover),
                "A reviewer hallucination about absent cover KPI callouts restarted the whole deck.");
            var demandedCoverMetrics = "{\"approved\":false,\"issues\":\"Headline figures are missing.\",\"findings\":[" +
                "{\"slide_id\":\"cover\",\"object_id\":\"subtitle\",\"severity\":\"blocker\",\"type\":\"coverage\",\"correction\":\"Add a scorecard to display Revenue, Cost, and Margin headline figures.\"}]}";
            Check((bool)Invoke(typeof(DocumentDraftHost), "OutlineReviewApprovedOrDeterministicallySatisfied", null,
                demandedCoverMetrics, "Repair all six slides.", cleanCoverSlides),
                "An outline reviewer invented a cover KPI requirement that the user never requested.");
            Check(!(bool)Invoke(typeof(DocumentDraftHost), "OutlineReviewApprovedOrDeterministicallySatisfied", null,
                demandedCoverMetrics, "The cover slide must include a scorecard showing revenue, cost, and margin.", cleanCoverSlides),
                "An explicit user requirement for cover metrics was incorrectly ignored.");
            var mixedCoverFinding = hallucinatedCoverCallouts.Replace("]}",
                ",{\"slide_id\":\"cover\",\"object_id\":\"subtitle\",\"severity\":\"blocker\",\"type\":\"facts\",\"correction\":\"Correct the unsupported period statement.\"}]}");
            Check(!(bool)Invoke(typeof(DocumentDraftHost), "ReviewApprovedOrSatisfiedPromptConstraint", null,
                mixedCoverFinding, "Repair all six slides.", cleanCover),
                "A real cover fact blocker was hidden beside a refuted KPI-callout hallucination.");
            var actualMetricCoverSlides = Invoke(Type("PresentationDraftWriter"), "ParseSlides", null,
                (object)json.Deserialize<object[]>("[{\"id\":\"cover\",\"title\":\"Atlas Components\",\"subtitle\":\"Revenue EUR 82,992\",\"layout\":\"cover\"}]"));
            Check(!(bool)Invoke(typeof(DocumentDraftHost), "ReviewApprovedOrSatisfiedPromptConstraint", null,
                hallucinatedCoverCallouts, "No data callouts on the cover.",
                ((IEnumerable)actualMetricCoverSlides).Cast<object>().Single()),
                "A real cover metric callout bypassed the reviewer.");
            var warningOnly = "{\"approved\":false,\"issues\":\"Use consistent casing.\",\"findings\":[" +
                "{\"slide_id\":\"period-trend\",\"object_id\":\"label\",\"severity\":\"warning\",\"type\":\"labels\",\"correction\":\"Use EUR instead of eur.\"}]}";
            Check((bool)Invoke(typeof(DocumentDraftHost), "ReviewApprovedOrSatisfiedPromptConstraint", null,
                warningOnly, "Use EUR.", ((IEnumerable)compliantChartSlides).Cast<object>().Single()),
                "A warning-only source review forced a full-deck retry.");
            var contradictoryBriefs = json.Deserialize<object[]>(
                "[{\"id\":\"trend\",\"layout\":\"chart\",\"required_content\":[\"Revenue EUR and Cost EUR series\"]}]");
            Reject(() => Invoke(typeof(DocumentDraftHost), "ValidatePromptChartBriefConstraints", null,
                "The chart must use only primary values for May and June.", contradictoryBriefs), "SLIDE_PRIMARY_SERIES_ONLY");
            var splitSeriesBrief = json.Deserialize<object[]>(
                "[{\"id\":\"trend\",\"layout\":\"chart\",\"required_content\":[\"Revenue series: 85519, 82992\",\"Cost series: 36702, 36714\"]}]");
            Reject(() => Invoke(typeof(DocumentDraftHost), "ValidatePromptChartBriefConstraints", null,
                "The chart must use only primary values for May and June.", splitSeriesBrief), "SLIDE_PRIMARY_SERIES_ONLY");
            var chartPurposeBrief = json.Deserialize<object[]>(
                "[{\"id\":\"trend\",\"layout\":\"chart\",\"purpose\":\"Show revenue and cost as a native editable chart\",\"required_content\":[\"Revenue series: 85519, 82992\"]}]");
            Reject(() => Invoke(typeof(DocumentDraftHost), "ValidatePromptChartBriefConstraints", null,
                "The chart must use only primary values for May and June.", chartPurposeBrief), "SLIDE_PRIMARY_SERIES_ONLY");
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
