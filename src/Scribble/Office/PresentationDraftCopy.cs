using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Web.Script.Serialization;

namespace Scribble.Office
{
    // Bounded six-page working copy for the development repair pilot. A source
    // slide is never used as a PresentationRevision target.
    internal sealed class PresentationDraftCopy
    {
        internal readonly object Source;
        internal readonly object Draft;
        private readonly int[] _sourceOrder;
        private readonly Dictionary<int, int> _slideIds =
            new Dictionary<int, int>();
        private readonly Dictionary<int, Dictionary<int, int>> _shapeIds =
            new Dictionary<int, Dictionary<int, int>>();
        private readonly Dictionary<int, string> _sourceContent =
            new Dictionary<int, string>();
        private readonly Dictionary<int, string> _sourceChartFingerprints =
            new Dictionary<int, string>();
        private readonly Dictionary<int, string> _draftFingerprints =
            new Dictionary<int, string>();
        private string _owner;
        private string _draftId;
        private string _sourceName;
        private string _sourceFullName;

        internal sealed class State
        {
            public int Version { get; set; } = 3;
            public string Owner { get; set; }
            public string DraftId { get; set; }
            public string SourceName { get; set; }
            public string SourceFullName { get; set; }
            public int[] SourceOrder { get; set; }
            public Dictionary<string, string> SourceContent { get; set; }
            public Dictionary<string, string> SourceChartFingerprints {
                get; set;
            }
            public Dictionary<string, int> SlideIds { get; set; }
            public Dictionary<string, Dictionary<string, int>> ShapeIds { get; set; }
            public Dictionary<string, string> DraftFingerprints { get; set; }
        }

        private PresentationDraftCopy(object source, object draft,
            int[] sourceOrder)
        { Source = source; Draft = draft; _sourceOrder = sourceOrder; }

        internal static PresentationDraftCopy Create(object application,
            object sourcePresentation, string owner)
        {
            dynamic source = sourcePresentation;
            dynamic app = application;
            if ((int)source.Slides.Count != 6 ||
                string.IsNullOrWhiteSpace(owner))
                throw new InvalidOperationException(
                    "REVISION_COPY_SCOPE: The pilot requires six source slides and a task owner.");
            var order = Enumerable.Range(1, 6).Select(index =>
                (int)source.Slides[index].SlideID).ToArray();
            dynamic draft = null;
            var stage = "new_draft";
            try
            {
                // PowerPoint's native chart engine requires a presentation
                // window, even when the application is driven through COM.
                draft = app.Presentations.Add(-1);
                draft.Tags.Add("ScribbleRevisionDraft", owner);
                var draftId = Guid.NewGuid().ToString("N");
                draft.Tags.Add("ScribblePresentationId", draftId);
                draft.PageSetup.SlideWidth = source.PageSetup.SlideWidth;
                draft.PageSetup.SlideHeight = source.PageSetup.SlideHeight;
                stage = "inspect_and_copy_source";
                var result = new PresentationDraftCopy(sourcePresentation,
                    (object)draft, order);
                result._owner = owner;
                result._draftId = draftId;
                result._sourceName = Convert.ToString(source.Name);
                result._sourceFullName = Convert.ToString(source.FullName);
                // PowerPoint's clipboard paste can terminate chart.dll when
                // a native chart slide is copied. A saved source can be read
                // into the new, unsaved draft without writing the source or
                // exporting its charts. Verify each inserted page below.
                var importSavedSource = (int)source.Saved != 0 &&
                    File.Exists(result._sourceFullName);
                if (importSavedSource)
                {
                    stage = "import_saved_source";
                    var inserted = (int)draft.Slides.InsertFromFile(
                        result._sourceFullName, 0, 1, 6);
                    if (inserted != 6 || (int)draft.Slides.Count != 6)
                        throw new InvalidOperationException(
                            "REVISION_COPY_INCOMPLETE: File import changed the page count.");
                }
                for (var index = 1; index <= 6; index++)
                {
                    stage = "inspect_source_slide_" + index;
                    dynamic original = source.Slides[index];
                    var originalId = (int)original.SlideID;
                    var fingerprint = PresentationInspection
                        .CopyContentFingerprint((object)original);
                    result._sourceContent[originalId] = fingerprint;
                    if (PresentationInspection.ContainsNativeChart(
                            (object)original))
                        result._sourceChartFingerprints[originalId] =
                            PresentationInspection.Fingerprint(
                                (object)original);
                    stage = "copy_source_slide_" + index;
                    dynamic copy = importSavedSource
                        ? draft.Slides[index]
                        : PresentationInspection.CopySlideTo(
                            (object)original, (object)draft);
                    if ((int)draft.Slides.Count != index)
                        throw new InvalidOperationException(
                            "REVISION_COPY_INCOMPLETE: Native paste changed the page count.");
                    if (PresentationInspection.CopyContentFingerprint(
                            (object)copy) != fingerprint)
                        throw new InvalidOperationException(
                            "REVISION_COPY_PRESERVATION: The copied page differs from the source.");
                    result._slideIds[originalId] = (int)copy.SlideID;
                    var shapes = new Dictionary<int, int>();
                    MapShapes((object)original.Shapes,
                        (object)copy.Shapes, shapes);
                    result._shapeIds[originalId] = shapes;
                }
                for (var index = 1; index <= 6; index++)
                {
                    stage = "fingerprint_draft_slide_" + index;
                    dynamic page = draft.Slides[index];
                    result._draftFingerprints[(int)page.SlideID] =
                        PresentationInspection.Fingerprint((object)page);
                }
                stage = "verify_copy";
                result.VerifySource();
                result.VerifyDraft();
                return result;
            }
            catch (Exception error)
            {
                if ((object)draft != null) try { draft.Close(); } catch { }
                throw new InvalidOperationException(
                    "REVISION_COPY_CREATE_FAILED at " + stage + ": " +
                    error.Message, error);
            }
        }

        internal string Snapshot()
        {
            VerifySource();
            VerifyDraft();
            return new JavaScriptSerializer { MaxJsonLength = 16000000 }
                .Serialize(new State
                {
                    Owner = _owner, DraftId = _draftId,
                    SourceName = _sourceName,
                    SourceFullName = _sourceFullName,
                    SourceOrder = _sourceOrder,
                    SourceContent = _sourceContent.ToDictionary(pair =>
                        pair.Key.ToString(), pair => pair.Value),
                    SourceChartFingerprints = _sourceChartFingerprints
                        .ToDictionary(pair => pair.Key.ToString(),
                            pair => pair.Value),
                    SlideIds = _slideIds.ToDictionary(pair =>
                        pair.Key.ToString(), pair => pair.Value),
                    ShapeIds = _shapeIds.ToDictionary(pair =>
                        pair.Key.ToString(), pair => pair.Value.ToDictionary(
                            shape => shape.Key.ToString(),
                            shape => shape.Value)),
                    DraftFingerprints = _draftFingerprints.ToDictionary(
                        pair => pair.Key.ToString(), pair => pair.Value)
                });
        }

        internal static PresentationDraftCopy Recover(object application,
            string snapshot)
        {
            State state;
            try
            {
                state = new JavaScriptSerializer { MaxJsonLength = 16000000 }
                    .Deserialize<State>(snapshot);
            }
            catch (ArgumentException)
            {
                throw new InvalidOperationException(
                    "REVISION_COPY_RECEIPT_INVALID");
            }
            if (state == null || state.Version != 3 ||
                string.IsNullOrWhiteSpace(state.Owner) ||
                string.IsNullOrWhiteSpace(state.DraftId) ||
                string.IsNullOrWhiteSpace(state.SourceFullName) ||
                state.SourceOrder == null || state.SourceOrder.Length != 6 ||
                state.SourceContent == null ||
                state.SourceContent.Count != 6 ||
                state.SourceChartFingerprints == null ||
                state.SourceChartFingerprints.Any(pair =>
                    !state.SourceContent.ContainsKey(pair.Key) ||
                    string.IsNullOrWhiteSpace(pair.Value)) ||
                state.SlideIds == null || state.SlideIds.Count != 6 ||
                state.ShapeIds == null || state.ShapeIds.Count != 6 ||
                state.DraftFingerprints == null ||
                state.DraftFingerprints.Count != 6 ||
                state.SourceOrder.Distinct().Count() != 6 ||
                state.SourceOrder.Any(id =>
                    !state.SourceContent.ContainsKey(id.ToString()) ||
                    !state.SlideIds.ContainsKey(id.ToString()) ||
                    !state.ShapeIds.ContainsKey(id.ToString())) ||
                state.ShapeIds.Values.Any(shapes => shapes == null) ||
                state.SlideIds.Values.Any(id =>
                    !state.DraftFingerprints.ContainsKey(id.ToString()) ||
                    string.IsNullOrWhiteSpace(
                        state.DraftFingerprints[id.ToString()])))
                throw new InvalidOperationException(
                    "REVISION_COPY_RECEIPT_INVALID");
            dynamic app = application;
            var sources = new List<object>();
            var drafts = new List<object>();
            foreach (dynamic candidate in app.Presentations)
            {
                if (Convert.ToString(candidate.Tags[
                        "ScribbleRevisionDraft"]) == state.Owner &&
                    Convert.ToString(candidate.Tags[
                        "ScribblePresentationId"]) == state.DraftId)
                    drafts.Add((object)candidate);
                if (Convert.ToString(candidate.Name) == state.SourceName &&
                    Convert.ToString(candidate.FullName) ==
                        state.SourceFullName &&
                    (int)candidate.Slides.Count == 6 &&
                    Enumerable.Range(1, 6).All(index =>
                        (int)candidate.Slides[index].SlideID ==
                            state.SourceOrder[index - 1] &&
                        PresentationInspection.CopyContentFingerprint(
                            (object)candidate.Slides[index]) ==
                            state.SourceContent[
                                state.SourceOrder[index - 1]
                                    .ToString()]))
                    sources.Add((object)candidate);
            }
            if (sources.Count != 1 || drafts.Count != 1 ||
                ReferenceEquals(sources[0], drafts[0]))
                throw new InvalidOperationException(
                    "REVISION_COPY_SESSION_UNAVAILABLE");
            dynamic draft = drafts[0];
            if ((int)draft.Slides.Count != 6 ||
                state.SlideIds.Values.Distinct().Count() != 6 ||
                state.SlideIds.Values.Any(id => !Enumerable.Range(1, 6)
                    .Any(index => (int)draft.Slides[index].SlideID == id)))
                throw new InvalidOperationException(
                    "REVISION_COPY_DRAFT_CHANGED");
            var result = new PresentationDraftCopy(sources[0], drafts[0],
                state.SourceOrder)
            {
                _owner = state.Owner, _draftId = state.DraftId,
                _sourceName = state.SourceName,
                _sourceFullName = state.SourceFullName
            };
            foreach (var pair in state.SourceContent)
                result._sourceContent.Add(int.Parse(pair.Key), pair.Value);
            foreach (var pair in state.SourceChartFingerprints)
                result._sourceChartFingerprints.Add(int.Parse(pair.Key),
                    pair.Value);
            foreach (var pair in state.SlideIds)
                result._slideIds.Add(int.Parse(pair.Key), pair.Value);
            foreach (var pair in state.ShapeIds)
                result._shapeIds.Add(int.Parse(pair.Key),
                    pair.Value.ToDictionary(shape => int.Parse(shape.Key),
                        shape => shape.Value));
            foreach (var pair in state.DraftFingerprints)
                result._draftFingerprints.Add(int.Parse(pair.Key),
                    pair.Value);
            result.VerifySource();
            result.VerifyDraft();
            return result;
        }

        internal object[] BindOperations(object[] operations)
        {
            VerifySource();
            VerifyDraft();
            if (operations == null || operations.Length == 0 ||
                operations.Length > 24)
                throw new InvalidOperationException(
                    "REVISION_COPY_OPERATIONS_INVALID");
            var bound = new List<object>();
            foreach (var raw in operations)
            {
                var sourceOperation = SamsungAuthoringPolicy.ReadMap(raw);
                var sourceId = Convert.ToInt32(
                    sourceOperation["slide_id"]);
                int draftId;
                if (!_slideIds.TryGetValue(sourceId, out draftId))
                    throw new InvalidOperationException(
                        "REVISION_COPY_SOURCE_SLIDE_CHANGED");
                var original = PresentationInspection.FindSlide(
                    Source, sourceId);
                if (SamsungAuthoringPolicy.Text(sourceOperation,
                        "fingerprint") != PresentationInspection
                            .Fingerprint(original))
                    throw new InvalidOperationException(
                        "REVISION_COPY_SOURCE_SLIDE_CHANGED");
                var mapped = new Dictionary<string, object>(
                    sourceOperation, StringComparer.Ordinal);
                mapped["slide_id"] = draftId;
                if (mapped.ContainsKey("shape_id"))
                {
                    var shapeId = Convert.ToInt32(
                        mapped["shape_id"]);
                    int newShapeId;
                    if (!_shapeIds[sourceId].TryGetValue(shapeId,
                            out newShapeId))
                        throw new InvalidOperationException(
                            "REVISION_COPY_SOURCE_SHAPE_CHANGED");
                    mapped["shape_id"] = newShapeId;
                }
                mapped["fingerprint"] = PresentationInspection
                    .Fingerprint(PresentationInspection.FindSlide(
                        Draft, draftId));
                bound.Add(mapped);
            }
            return bound.ToArray();
        }

        internal WorkbookMonthlyChartFacts.Result RecreateSalesChartFromWorkbook(
            int sourceSlideId, int sourceShapeId, string workbookPath,
            float left, float top, float width, float height)
        {
            VerifySource();
            VerifyDraft();
            var facts = WorkbookMonthlyChartFacts.ReadSalesLedger(
                workbookPath, CancellationToken.None);
            if (facts.Categories.Length != 6 ||
                facts.RevenueEur.Length != 6 || facts.CostEur.Length != 6)
                throw new InvalidOperationException(
                    "REVISION_CHART_FACTS_INCOMPLETE");
            int draftSlideId;
            int draftShapeId;
            if (!_slideIds.TryGetValue(sourceSlideId,
                    out draftSlideId) ||
                !_shapeIds[sourceSlideId].TryGetValue(sourceShapeId,
                    out draftShapeId))
                throw new InvalidOperationException(
                    "REVISION_CHART_SOURCE_CHANGED");
            dynamic slide = PresentationInspection.FindSlide(Draft,
                draftSlideId);
            dynamic oldChart = PresentationInspection.FindShape(
                (object)slide, draftShapeId);
            if ((int)oldChart.HasChart == 0 ||
                left < 0 || top < 0 || width < 100 || height < 100 ||
                left + width > (float)((dynamic)Draft).PageSetup.SlideWidth ||
                top + height > (float)((dynamic)Draft).PageSetup.SlideHeight)
                throw new InvalidOperationException(
                    "REVISION_CHART_REPLACEMENT_INVALID");
            var chart = new PresentationDraftWriter.DraftChart(
                DraftChartTypes.ColumnClustered,
                "Revenue EUR / Cost EUR (EUR)", facts.Categories,
                new[] {
                    new PresentationDraftWriter.DraftChartSeries(
                        "Revenue EUR", facts.RevenueEur.Select(value =>
                            (double?)value).ToArray()),
                    new PresentationDraftWriter.DraftChartSeries(
                        "Cost EUR", facts.CostEur.Select(value =>
                            (double?)value).ToArray())
                });
            var before = (int)slide.Shapes.Count;
            var created = false;
            for (var attempt = 0; attempt < 3; attempt++)
            {
                if (PresentationDraftWriter.AddChartToSlide(slide,
                        chart, left, top, width, height))
                { created = true; break; }
                while ((int)slide.Shapes.Count > before)
                    slide.Shapes[(int)slide.Shapes.Count].Delete();
                if (attempt == 2 || !PresentationDraftWriter
                        .RetryableSamsungChartFailure(
                            PresentationDraftWriter.LastChartFailure))
                    break;
                Thread.Sleep(350 * (attempt + 1));
            }
            if (!created || (int)slide.Shapes.Count != before + 1)
                throw new InvalidOperationException(
                    "REVISION_CHART_RECREATE_FAILED: " +
                    PresentationDraftWriter.LastChartFailure);
            dynamic replacement = slide.Shapes[before + 1];
            if ((int)replacement.HasChart == 0)
                throw new InvalidOperationException(
                    "REVISION_CHART_RECREATE_NOT_NATIVE");
            var replacementId = (int)replacement.Id;
            try { oldChart.Delete(); }
            catch
            {
                replacement.Delete();
                throw;
            }
            _shapeIds[sourceSlideId][sourceShapeId] =
                replacementId;
            // Office can finish updating the saved chart cache after the
            // embedded workbook closes. Seal only a state that stays the
            // same across several spaced package snapshots. A later change
            // still fails the ordinary exact draft verification.
            string stable = null;
            string previous = null;
            var consecutive = 0;
            for (var attempt = 0; attempt < 8; attempt++)
            {
                var current = PresentationInspection.Fingerprint(
                    (object)slide);
                consecutive = current == previous ? consecutive + 1 : 1;
                previous = current;
                if (attempt >= 4 && consecutive >= 3)
                { stable = current; break; }
                if (attempt < 7) Thread.Sleep(300);
            }
            if (stable == null)
                throw new InvalidOperationException(
                    "REVISION_CHART_FINGERPRINT_UNSTABLE");
            _draftFingerprints[draftSlideId] = stable;
            VerifySource();
            VerifyDraft();
            return facts;
        }

        internal void VerifyDraft()
        {
            dynamic draft = Draft;
            if (Convert.ToString(draft.Tags["ScribbleRevisionDraft"]) !=
                    _owner ||
                Convert.ToString(draft.Tags["ScribblePresentationId"]) !=
                    _draftId ||
                !string.IsNullOrEmpty(Convert.ToString(draft.Path)) ||
                (int)draft.Slides.Count != _sourceOrder.Length ||
                _draftFingerprints.Count != _sourceOrder.Length)
                throw new InvalidOperationException(
                    "REVISION_COPY_DRAFT_CHANGED");
            for (var index = 1; index <= _sourceOrder.Length; index++)
            {
                dynamic slide = draft.Slides[index];
                var id = (int)slide.SlideID;
                string expected;
                if (_slideIds[_sourceOrder[index - 1]] != id ||
                    !_draftFingerprints.TryGetValue(id, out expected) ||
                    PresentationInspection.Fingerprint((object)slide) !=
                        expected)
                    throw new InvalidOperationException(
                        "REVISION_COPY_DRAFT_CHANGED: slide " + id);
            }
        }

        internal void AcceptRevision(PresentationRevision revision)
        {
            VerifySource();
            if (revision == null ||
                !ReferenceEquals(revision.Presentation, Draft) ||
                revision.Status != "applied" ||
                revision.Items.Count == 0 ||
                revision.Items.Any(item => !item.Applied || item.Deleted ||
                    string.IsNullOrWhiteSpace(item.After) ||
                    !_draftFingerprints.ContainsKey(item.SlideId) ||
                    item.Operations.Any(operation => new[] { "insert",
                        "delete", "move" }.Contains(
                            SamsungAuthoringPolicy.Text(operation,
                                "kind")))))
                throw new InvalidOperationException(
                    "REVISION_COPY_RECEIPT_INVALID");
            dynamic draft = Draft;
            if ((int)draft.Slides.Count != _sourceOrder.Length ||
                Convert.ToString(draft.Tags["ScribbleRevisionDraft"]) !=
                    _owner ||
                Convert.ToString(draft.Tags["ScribblePresentationId"]) !=
                    _draftId)
                throw new InvalidOperationException(
                    "REVISION_COPY_DRAFT_CHANGED");
            var changed = revision.Items.ToDictionary(item =>
                item.SlideId, item => item);
            for (var index = 1; index <= _sourceOrder.Length; index++)
            {
                dynamic slide = draft.Slides[index];
                var id = (int)slide.SlideID;
                PresentationRevision.Item item;
                if (_slideIds[_sourceOrder[index - 1]] != id ||
                    PresentationInspection.Fingerprint((object)slide) !=
                        (changed.TryGetValue(id, out item) ? item.After :
                            _draftFingerprints[id]))
                    throw new InvalidOperationException(
                        "REVISION_COPY_DRAFT_CHANGED");
            }
            foreach (var item in revision.Items)
            {
                _draftFingerprints[item.SlideId] = item.After;
                if (item.Operations.Any(operation =>
                    SamsungAuthoringPolicy.Text(operation, "kind") ==
                        "replace_slide"))
                {
                    var sourceId = _slideIds.Single(pair =>
                        pair.Value == item.SlideId).Key;
                    _shapeIds[sourceId].Clear();
                }
            }
            VerifyDraft();
        }

        internal void VerifySource()
        {
            dynamic source = Source;
            if (Convert.ToString(source.Name) != _sourceName ||
                Convert.ToString(source.FullName) != _sourceFullName)
                throw new InvalidOperationException(
                    "REVISION_COPY_SOURCE_CHANGED");
            if ((int)source.Slides.Count != _sourceOrder.Length)
                throw new InvalidOperationException(
                    "REVISION_COPY_SOURCE_CHANGED");
            for (var index = 1; index <= _sourceOrder.Length; index++)
            {
                dynamic slide = source.Slides[index];
                var id = (int)slide.SlideID;
                if (id != _sourceOrder[index - 1] ||
                    PresentationInspection.CopyContentFingerprint(
                        (object)slide) != _sourceContent[id])
                    throw new InvalidOperationException(
                        "REVISION_COPY_SOURCE_CHANGED");
                string chartFingerprint;
                var hadChart = _sourceChartFingerprints.TryGetValue(id,
                    out chartFingerprint);
                if (PresentationInspection.ContainsNativeChart(
                        (object)slide) != hadChart ||
                    (hadChart && PresentationInspection.Fingerprint(
                        (object)slide) != chartFingerprint))
                    throw new InvalidOperationException(
                        "REVISION_COPY_SOURCE_CHANGED: chart " + id);
            }
        }

        private static void MapShapes(object sourceShapes,
            object copiedShapes, Dictionary<int, int> map)
        {
            dynamic source = sourceShapes;
            dynamic copied = copiedShapes;
            if ((int)source.Count != (int)copied.Count)
                throw new InvalidOperationException(
                    "REVISION_COPY_SHAPE_MAPPING_FAILED");
            for (var index = 1; index <= (int)source.Count; index++)
            {
                dynamic original = source[index];
                dynamic clone = copied[index];
                if ((int)original.Type != (int)clone.Type)
                    throw new InvalidOperationException(
                        "REVISION_COPY_SHAPE_MAPPING_FAILED");
                map.Add((int)original.Id, (int)clone.Id);
                if ((int)original.Type == 6)
                    MapShapes((object)original.GroupItems,
                        (object)clone.GroupItems, map);
            }
        }
    }
}
