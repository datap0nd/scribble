using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Web.Script.Serialization;
using Scribble.Chat;

namespace Scribble.Office
{
    // An untitled native copy is the working deck for the repair pilot. A
    // source slide is never used as a PresentationRevision target.
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

        internal sealed class MonthlyChartBinding
        {
            public int SourceSlideId;
            public int SourceShapeId;
            public float Left;
            public float Top;
            public float Width;
            public float Height;
            public WorkbookMonthlyChartFacts.BoundSeries Facts;
            public WorkbookMonthlyChartFacts.BoundSeries ContextFacts;
        }

        // Reading series names and category labels does not open the
        // embedded ChartData.Workbook. Reject ambiguous or unbound source
        // series before consuming write authorization.
        internal static MonthlyChartBinding[] BindMonthlyCharts(
            object sourcePresentation, string workbookPath,
            string trustedRequest, CancellationToken token)
        {
            dynamic source = sourcePresentation;
            var request = trustedRequest ?? string.Empty;
            var onlyPrimary = Regex.IsMatch(request,
                @"\bonly (?:the )?primary\b|\bsingle (?:primary )?series\b",
                RegexOptions.IgnoreCase);
            var primary = RequestedMeasure(request, "primary");
            var secondary = RequestedMeasure(request, "secondary");
            var bindings = new List<MonthlyChartBinding>();
            var cache = new Dictionary<string,
                WorkbookMonthlyChartFacts.BoundSeries>(
                    StringComparer.Ordinal);
            for (var slideIndex = 1; slideIndex <=
                (int)source.Slides.Count; slideIndex++)
            {
                dynamic slide = source.Slides[slideIndex];
                for (var shapeIndex = 1; shapeIndex <=
                    (int)slide.Shapes.Count; shapeIndex++)
                {
                    token.ThrowIfCancellationRequested();
                    dynamic shape = slide.Shapes[shapeIndex];
                    if ((int)shape.HasChart == 0) continue;
                    var count = (int)shape.Chart.SeriesCollection().Count;
                    if (count < 1 || count > 6)
                        throw new InvalidOperationException(
                            "REVISION_CHART_SERIES_AMBIGUOUS");
                    var names = new List<string>();
                    string[] categories = null;
                    for (var index = 1; index <= count; index++)
                    {
                        dynamic item = shape.Chart.SeriesCollection(index);
                        var labels = CategoryLabels((object)item.XValues);
                        if (categories == null) categories = labels;
                        else if (!categories.SequenceEqual(labels,
                            StringComparer.Ordinal))
                            throw new InvalidOperationException(
                                "REVISION_CHART_CATEGORIES_AMBIGUOUS");
                        names.Add(Convert.ToString(item.Name).Trim());
                    }
                    if (categories == null || categories.Length < 2 ||
                        categories.Any(label => !Regex.IsMatch(label,
                            @"^\d{4}-(?:0[1-9]|1[0-2])$")))
                        continue;
                    if (names.Any(string.IsNullOrWhiteSpace) ||
                        names.Distinct(StringComparer.OrdinalIgnoreCase)
                            .Count() != names.Count ||
                        (!string.IsNullOrEmpty(primary) &&
                         NormalizeChartName(primary) !=
                         NormalizeChartName(names[0])) ||
                        (!onlyPrimary && !string.IsNullOrEmpty(secondary) &&
                         (names.Count < 2 ||
                          NormalizeChartName(secondary) !=
                          NormalizeChartName(names[1]))))
                        throw new InvalidOperationException(
                            "REVISION_CHART_SERIES_AMBIGUOUS");
                    var chosen = onlyPrimary
                        ? names.Take(1).ToArray() : names.ToArray();
                    var key = string.Join("\0", chosen);
                    WorkbookMonthlyChartFacts.BoundSeries facts;
                    if (!cache.TryGetValue(key, out facts))
                    {
                        facts = WorkbookMonthlyChartFacts.ReadBoundSeries(
                            workbookPath, chosen, token);
                        cache.Add(key, facts);
                    }
                    var contextFacts = facts;
                    if (onlyPrimary &&
                        !string.IsNullOrWhiteSpace(secondary))
                    {
                        var contextNames = chosen.Concat(new[] {
                            secondary
                        }).ToArray();
                        var contextKey = string.Join("\0",
                            contextNames);
                        if (!cache.TryGetValue(contextKey,
                                out contextFacts))
                        {
                            contextFacts = WorkbookMonthlyChartFacts
                                .ReadBoundSeries(workbookPath,
                                    contextNames, token);
                            cache.Add(contextKey, contextFacts);
                        }
                    }
                    if (!string.Equals(contextFacts.SourceSha256,
                            facts.SourceSha256,
                            StringComparison.OrdinalIgnoreCase))
                        throw new InvalidOperationException(
                            "REVISION_CHART_SOURCE_CHANGED");
                    if (categories.Length != facts.Categories.Length)
                        throw new InvalidOperationException(
                            "REVISION_CHART_PERIOD_COVERAGE_INVALID");
                    bindings.Add(new MonthlyChartBinding {
                        SourceSlideId = (int)slide.SlideID,
                        SourceShapeId = (int)shape.Id,
                        Left = (float)shape.Left,
                        Top = (float)shape.Top,
                        Width = (float)shape.Width,
                        Height = (float)shape.Height,
                        Facts = facts,
                        ContextFacts = contextFacts
                    });
                }
            }
            if (bindings.Count == 0)
                throw new InvalidOperationException(
                    "REVISION_CHART_SOURCE_MISSING");
            return bindings.ToArray();
        }

        private static string[] CategoryLabels(object value)
        {
            if (value is string) return new[] { (string)value };
            var values = value as IEnumerable;
            return values == null ? new string[0] : values.Cast<object>()
                .Select(item => Convert.ToString(item).Trim()).ToArray();
        }

        private static string RequestedMeasure(string request,
            string role)
        {
            var match = Regex.Match(request,
                @"\b" + role +
                @"\s+(?:measure|metric|series)\s+(?:is|:|=)\s*(?<name>[^.;\r\n]+)",
                RegexOptions.IgnoreCase);
            return match.Success ? match.Groups["name"].Value.Trim() :
                string.Empty;
        }

        private static string NormalizeChartName(string value)
        {
            return Regex.Replace(value ?? string.Empty, @"[^a-z0-9]",
                string.Empty, RegexOptions.IgnoreCase).ToLowerInvariant();
        }

        private static bool MonthlyChartCategories(object chartShape)
        {
            dynamic shape = chartShape;
            if ((int)shape.HasChart == 0) return false;
            var count = (int)shape.Chart.SeriesCollection().Count;
            if (count < 1 || count > 6) return false;
            string[] first = null;
            for (var index = 1; index <= count; index++)
            {
                dynamic series = shape.Chart.SeriesCollection(index);
                var labels = CategoryLabels((object)series.XValues);
                if (first == null) first = labels;
                else if (!first.SequenceEqual(labels,
                    StringComparer.Ordinal)) return false;
            }
            return first != null && first.Length >= 2 &&
                first.All(label => Regex.IsMatch(label,
                    @"^\d{4}-(?:0[1-9]|1[0-2])$"));
        }

        internal static int[] MeasuredReplacementSlides(
            object sourcePresentation)
        {
            dynamic source = sourcePresentation;
            var width = (float)source.PageSetup.SlideWidth;
            var height = (float)source.PageSetup.SlideHeight;
            var damaged = new List<int>();
            for (var slideIndex = 1; slideIndex <=
                (int)source.Slides.Count; slideIndex++)
            {
                dynamic slide = source.Slides[slideIndex];
                var overflow = false;
                for (var shapeIndex = 1; shapeIndex <=
                    (int)slide.Shapes.Count; shapeIndex++)
                {
                    dynamic shape = slide.Shapes[shapeIndex];
                    if ((int)shape.HasChart != 0) continue;
                    if ((float)shape.Left < -.5f ||
                        (float)shape.Top < -.5f ||
                        (float)shape.Left + (float)shape.Width >
                            width + .5f ||
                        (float)shape.Top + (float)shape.Height >
                            height + .5f)
                    { overflow = true; break; }
                    if ((int)shape.HasTable != 0)
                    {
                        dynamic table = shape.Table;
                        for (var row = 1; row <= (int)table.Rows.Count &&
                            !overflow; row++)
                        for (var column = 1; column <=
                            (int)table.Columns.Count; column++)
                        {
                            dynamic cell = table.Cell(row, column).Shape;
                            dynamic range = cell.TextFrame.TextRange;
                            if (PresentationRevision.NativeTextOverflows(
                                    Convert.ToString(range.Text),
                                    (float)range.BoundHeight,
                                    (float)range.BoundWidth,
                                    (float)cell.Height,
                                    (float)cell.Width))
                            { overflow = true; break; }
                        }
                    }
                    else if ((int)shape.HasTextFrame != 0)
                    {
                        dynamic range = shape.TextFrame.TextRange;
                        if (PresentationRevision.NativeTextOverflows(
                                Convert.ToString(range.Text),
                                (float)range.BoundHeight,
                                (float)range.BoundWidth,
                                (float)shape.Height,
                                (float)shape.Width))
                        { overflow = true; break; }
                    }
                }
                if (overflow) damaged.Add((int)slide.SlideID);
            }
            return damaged.ToArray();
        }

        internal static PresentationDraftCopy Create(object application,
            object sourcePresentation, string owner)
        {
            dynamic source = sourcePresentation;
            dynamic app = application;
            var count = (int)source.Slides.Count;
            var sourcePath = Convert.ToString(source.FullName);
            if (count < 1 || string.IsNullOrWhiteSpace(owner) ||
                string.IsNullOrWhiteSpace(Convert.ToString(source.Path)) ||
                string.IsNullOrWhiteSpace(sourcePath) ||
                !File.Exists(sourcePath))
                throw new InvalidOperationException(
                    "REVISION_COPY_SCOPE: A saved source deck and task owner are required.");
            var order = Enumerable.Range(1, count).Select(index =>
                (int)source.Slides[index].SlideID).ToArray();
            dynamic draft = null;
            var stage = "copy_source_file";
            var temporary = Path.Combine(Path.GetTempPath(),
                "scribble-revision-copy-" + Guid.NewGuid().ToString("N") +
                ".pptx");
            try
            {
                var sourceHash = ExternalContextDocument.FingerprintFile(
                    sourcePath);
                File.Copy(sourcePath, temporary);
                if (!string.Equals(sourceHash,
                        ExternalContextDocument.FingerprintFile(temporary),
                        StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException(
                        "REVISION_COPY_SOURCE_CHANGED");
                stage = "open_untitled_copy";
                // Reopening sourcePath itself returns the existing saved
                // presentation. A separate disposable byte copy is required.
                draft = app.Presentations.Open(temporary, -1, -1, -1);
                if (!string.IsNullOrEmpty(Convert.ToString(draft.Path)) ||
                    (int)draft.Slides.Count != count)
                    throw new InvalidOperationException(
                        "REVISION_COPY_NOT_UNTITLED");
                draft.Tags.Add("ScribbleRevisionDraft", owner);
                var draftId = Guid.NewGuid().ToString("N");
                draft.Tags.Add("ScribblePresentationId", draftId);
                stage = "inspect_native_copy";
                var result = new PresentationDraftCopy(sourcePresentation,
                    (object)draft, order);
                result._owner = owner;
                result._draftId = draftId;
                result._sourceName = Convert.ToString(source.Name);
                result._sourceFullName = Convert.ToString(source.FullName);
                for (var index = 1; index <= count; index++)
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
                    stage = "map_copy_slide_" + index;
                    var shapes = new Dictionary<int, int>();
                    dynamic copy = draft.Slides[index];
                    if (PresentationInspection.CopyContentFingerprint(
                            (object)copy) != fingerprint)
                        throw new InvalidOperationException(
                            "REVISION_COPY_PRESERVATION: The native copy differs from the source.");
                    result._slideIds[originalId] = (int)copy.SlideID;
                    MapShapes((object)original.Shapes,
                        (object)copy.Shapes, shapes);
                    // Verify the native copy first, then remove only its
                    // monthly charts before patch staging. The source chart
                    // remains untouched and is rebuilt from bound facts.
                    for (var shapeIndex = 1; shapeIndex <=
                        (int)original.Shapes.Count; shapeIndex++)
                    {
                        dynamic sourceShape =
                            original.Shapes[shapeIndex];
                        if ((int)sourceShape.HasChart == 0 ||
                            !MonthlyChartCategories((object)sourceShape))
                            continue;
                        var sourceShapeId = (int)sourceShape.Id;
                        var draftShapeId = shapes[sourceShapeId];
                        dynamic draftShape = PresentationInspection
                            .FindShape((object)copy, draftShapeId);
                        draftShape.Delete();
                        shapes[sourceShapeId] = -1;
                    }
                    result._shapeIds[originalId] = shapes;
                }
                for (var index = 1; index <= count; index++)
                {
                    stage = "fingerprint_draft_slide_" + index;
                    dynamic page = draft.Slides[index];
                    result._draftFingerprints[(int)page.SlideID] =
                        PresentationInspection.Fingerprint((object)page);
                }
                stage = "verify_copy";
                if (!string.Equals(sourceHash,
                        ExternalContextDocument.FingerprintFile(sourcePath),
                        StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException(
                        "REVISION_COPY_SOURCE_CHANGED");
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
            finally
            {
                try { if (File.Exists(temporary)) File.Delete(temporary); }
                catch { }
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
                state.SourceOrder == null || state.SourceOrder.Length < 1 ||
                state.SourceContent == null ||
                state.SourceContent.Count != state.SourceOrder.Length ||
                state.SourceChartFingerprints == null ||
                state.SourceChartFingerprints.Any(pair =>
                    !state.SourceContent.ContainsKey(pair.Key) ||
                    string.IsNullOrWhiteSpace(pair.Value)) ||
                state.SlideIds == null ||
                state.SlideIds.Count != state.SourceOrder.Length ||
                state.ShapeIds == null ||
                state.ShapeIds.Count != state.SourceOrder.Length ||
                state.DraftFingerprints == null ||
                state.DraftFingerprints.Count !=
                    state.SourceOrder.Length ||
                state.SourceOrder.Distinct().Count() !=
                    state.SourceOrder.Length ||
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
                    (int)candidate.Slides.Count ==
                        state.SourceOrder.Length &&
                    Enumerable.Range(1, state.SourceOrder.Length).All(index =>
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
            if ((int)draft.Slides.Count != state.SourceOrder.Length ||
                state.SlideIds.Values.Distinct().Count() !=
                    state.SourceOrder.Length ||
                state.SlideIds.Values.Any(id => !Enumerable.Range(1,
                    state.SourceOrder.Length)
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

        internal void RecreateBoundChart(MonthlyChartBinding binding)
        {
            VerifySource();
            VerifyDraft();
            if (binding == null || binding.Facts == null ||
                binding.Facts.Names == null ||
                binding.Facts.Values == null ||
                binding.Facts.Names.Length != binding.Facts.Values.Length)
                throw new InvalidOperationException(
                    "REVISION_CHART_BINDING_INVALID");
            int draftSlideId;
            int draftShapeId;
            if (!_slideIds.TryGetValue(binding.SourceSlideId,
                    out draftSlideId) ||
                !_shapeIds[binding.SourceSlideId].TryGetValue(
                    binding.SourceShapeId, out draftShapeId))
                throw new InvalidOperationException(
                    "REVISION_CHART_SOURCE_CHANGED");
            dynamic slide = PresentationInspection.FindSlide(Draft,
                draftSlideId);
            dynamic oldChart = draftShapeId < 0 ? null :
                PresentationInspection.FindShape((object)slide,
                    draftShapeId);
            if (((object)oldChart == null &&
                    !_sourceChartFingerprints.ContainsKey(
                        binding.SourceSlideId)) ||
                ((object)oldChart != null &&
                    (int)oldChart.HasChart == 0))
                throw new InvalidOperationException(
                    "REVISION_CHART_SOURCE_CHANGED");
            var unit = ChartUnit(binding.Facts.Names);
            var title = string.Join(" / ", binding.Facts.Names) +
                " (" + unit + ")";
            var chart = new PresentationDraftWriter.DraftChart(
                DraftChartTypes.ColumnClustered, title,
                binding.Facts.Categories,
                binding.Facts.Names.Select((name, index) =>
                    new PresentationDraftWriter.DraftChartSeries(name,
                        binding.Facts.Values[index].Select(value =>
                            (double?)value).ToArray())).ToArray());
            var width = (float)((dynamic)Draft).PageSetup.SlideWidth;
            var height = (float)((dynamic)Draft).PageSetup.SlideHeight;
            var candidates = new[] {
                new[] { binding.Left, binding.Top,
                    binding.Width, binding.Height },
                new[] { width * .069f, height * .293f,
                    width * .859f, height * .515f },
                new[] { width * .52f, height * .28f,
                    width * .41f, height * .54f },
                new[] { width * .07f, height * .28f,
                    width * .41f, height * .54f }
            };
            var placement = candidates.FirstOrDefault(candidate =>
                ChartPlacementClear((object)slide, candidate,
                    width, height));
            if (placement == null)
                throw new InvalidOperationException(
                    "REVISION_CHART_PLACEMENT_AMBIGUOUS");
            var left = placement[0];
            var top = placement[1];
            var chartWidth = placement[2];
            var chartHeight = placement[3];
            var before = (int)slide.Shapes.Count;
            var created = false;
            for (var attempt = 0; attempt < 3; attempt++)
            {
                if (PresentationDraftWriter.AddChartToSlide(slide,
                        chart, left, top, chartWidth, chartHeight))
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
            if ((object)oldChart != null)
            {
                try { oldChart.Delete(); }
                catch
                {
                    replacement.Delete();
                    throw;
                }
            }
            _shapeIds[binding.SourceSlideId][binding.SourceShapeId] =
                replacementId;
            string stable = null;
            string previous = null;
            var consecutive = 0;
            for (var attempt = 0; attempt < 8; attempt++)
            {
                var current = PresentationInspection
                    .PackageSlideFingerprint((object)slide);
                consecutive = current == previous ? consecutive + 1 : 1;
                previous = current;
                if (attempt >= 4 && consecutive >= 3)
                { stable = current; break; }
                if (attempt < 7) Thread.Sleep(300);
            }
            if (stable == null)
                throw new InvalidOperationException(
                    "REVISION_CHART_FINGERPRINT_UNSTABLE");
            _draftFingerprints[draftSlideId] = "pkg:" + stable;
            VerifySource();
            VerifyDraft();
        }

        private static string ChartUnit(string[] names)
        {
            if (names.All(name => name.EndsWith("EUR",
                StringComparison.OrdinalIgnoreCase))) return "EUR";
            if (names.All(name => Regex.IsMatch(name, @"\bhours?\b",
                RegexOptions.IgnoreCase))) return "hours";
            if (names.All(name => Regex.IsMatch(name, @"\bunits?\b",
                RegexOptions.IgnoreCase))) return "units";
            throw new InvalidOperationException(
                "REVISION_CHART_UNITS_AMBIGUOUS");
        }

        private static bool ChartPlacementClear(object slide,
            float[] box, float pageWidth, float pageHeight)
        {
            if (box[0] < 0 || box[1] < 0 || box[2] < 100 ||
                box[3] < 100 || box[0] + box[2] > pageWidth ||
                box[1] + box[3] > pageHeight)
                return false;
            dynamic page = slide;
            for (var index = 1; index <= (int)page.Shapes.Count;
                index++)
            {
                dynamic shape = page.Shapes[index];
                var meaningful = (int)shape.HasTable != 0 ||
                    (int)shape.HasChart != 0 ||
                    ((int)shape.HasTextFrame != 0 &&
                     !string.IsNullOrWhiteSpace(Convert.ToString(
                         shape.TextFrame.TextRange.Text)));
                if (!meaningful) continue;
                var overlapWidth = Math.Min(box[0] + box[2],
                    (float)shape.Left + (float)shape.Width) -
                    Math.Max(box[0], (float)shape.Left);
                var overlapHeight = Math.Min(box[1] + box[3],
                    (float)shape.Top + (float)shape.Height) -
                    Math.Max(box[1], (float)shape.Top);
                if (overlapWidth > 2f && overlapHeight > 2f)
                    return false;
            }
            return true;
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
                    (expected.StartsWith("pkg:",
                        StringComparison.Ordinal)
                        ? "pkg:" + PresentationInspection
                            .PackageSlideFingerprint((object)slide)
                        : PresentationInspection.Fingerprint(
                            (object)slide)) != expected)
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

        internal void AcceptDirectRevision(
            IDictionary<int, string> changed,
            object[] operations)
        {
            VerifySource();
            if (changed == null || changed.Count == 0 ||
                operations == null || operations.Length == 0)
                throw new InvalidOperationException(
                    "REVISION_COPY_RECEIPT_INVALID");
            dynamic draft = Draft;
            if ((int)draft.Slides.Count != _sourceOrder.Length ||
                Convert.ToString(draft.Tags["ScribbleRevisionDraft"]) !=
                    _owner ||
                Convert.ToString(draft.Tags["ScribblePresentationId"]) !=
                    _draftId ||
                !string.IsNullOrEmpty(Convert.ToString(draft.Path)))
                throw new InvalidOperationException(
                    "REVISION_COPY_DRAFT_CHANGED");
            for (var index = 1; index <= _sourceOrder.Length; index++)
            {
                dynamic slide = draft.Slides[index];
                var id = (int)slide.SlideID;
                string expected;
                if (_slideIds[_sourceOrder[index - 1]] != id ||
                    !(changed.TryGetValue(id, out expected) ||
                      _draftFingerprints.TryGetValue(id, out expected)) ||
                    PresentationInspection.Fingerprint((object)slide) !=
                        expected)
                    throw new InvalidOperationException(
                        "REVISION_COPY_DRAFT_CHANGED");
            }
            foreach (var pair in changed)
                _draftFingerprints[pair.Key] = pair.Value;
            foreach (var raw in operations)
            {
                var operation = SamsungAuthoringPolicy.ReadMap(raw);
                if (SamsungAuthoringPolicy.Text(operation, "kind") !=
                        "replace_slide") continue;
                var draftId = Convert.ToInt32(operation["slide_id"]);
                var sourceId = _slideIds.Single(pair =>
                    pair.Value == draftId).Key;
                _shapeIds[sourceId].Clear();
            }
            VerifyDraft();
        }

        internal void DiscardOwnedDraft()
        {
            dynamic draft = Draft;
            if (Convert.ToString(draft.Tags["ScribbleRevisionDraft"]) !=
                    _owner ||
                Convert.ToString(draft.Tags["ScribblePresentationId"]) !=
                    _draftId ||
                !string.IsNullOrEmpty(Convert.ToString(draft.Path)))
                throw new InvalidOperationException(
                    "REVISION_COPY_DISCARD_UNSAFE");
            draft.Saved = -1;
            draft.Close();
            VerifySource();
        }

        // Native style repairs use measured geometry and table header rows.
        // The model never supplies a color or a font-size threshold.
        internal object[] MeasuredNativeStyleOperations(
            int[] replacedSourceSlideIds)
        {
            VerifySource();
            VerifyDraft();
            dynamic draft = Draft;
            var operations = new List<object>();
            var replaced = new HashSet<int>(
                replacedSourceSlideIds ?? new int[0]);
            var blue = MetoTheme.Rgb(SamsungSlideDesign.Blue);
            var pageHeight = (float)draft.PageSetup.SlideHeight;
            for (var index = 1; index <= _sourceOrder.Length; index++)
            {
                if (replaced.Contains(_sourceOrder[index - 1]))
                    continue;
                dynamic slide = draft.Slides[index];
                var fingerprint = PresentationInspection.Fingerprint(
                    (object)slide);
                var textShapes = new List<object>();
                for (var shapeIndex = 1; shapeIndex <=
                    (int)slide.Shapes.Count; shapeIndex++)
                {
                    dynamic shape = slide.Shapes[shapeIndex];
                    if ((int)shape.HasTable != 0)
                    {
                        dynamic table = shape.Table;
                        if ((int)table.Rows.Count < 2) continue;
                        var headers = Enumerable.Range(1,
                            (int)table.Columns.Count).Select(column =>
                            (string)Convert.ToString(table.Cell(1, column)
                                .Shape.TextFrame.TextRange.Text)).ToArray();
                        if (headers.Any(string.IsNullOrWhiteSpace) ||
                            headers.Distinct(StringComparer.OrdinalIgnoreCase)
                                .Count() != headers.Length)
                            continue;
                        for (var column = 1; column <=
                            (int)table.Columns.Count; column++)
                        {
                            var oldColor = (int)table.Cell(1, column)
                                .Shape.Fill.ForeColor.RGB;
                            if (oldColor == blue) continue;
                            operations.Add(new Dictionary<string, object>
                            {
                                { "kind", "table_cell_fill" },
                                { "slide_id", (int)slide.SlideID },
                                { "fingerprint", fingerprint },
                                { "shape_id", (int)shape.Id },
                                { "row", 1 }, { "column", column },
                                { "before_color", oldColor },
                                { "color", blue }
                            });
                        }
                    }
                    if ((int)shape.HasTextFrame == 0 ||
                        (int)shape.HasChart != 0 ||
                        (float)shape.Top >= pageHeight * .9f ||
                        string.IsNullOrWhiteSpace(Convert.ToString(
                            shape.TextFrame.TextRange.Text)))
                        continue;
                    textShapes.Add((object)shape);
                }
                var targets = new Dictionary<int, float>();
                foreach (dynamic shape in textShapes)
                {
                    var before = (float)shape.TextFrame.TextRange.Font.Size;
                    if (before > 0 && before < 14f)
                        targets[(int)shape.Id] = 14f;
                }
                foreach (dynamic shape in textShapes)
                {
                    var before = (float)shape.TextFrame.TextRange.Font.Size;
                    if (before < 14f || before >= 30f) continue;
                    var peers = textShapes.Cast<dynamic>().Where(peer =>
                        (int)peer.Id != (int)shape.Id &&
                        Math.Abs((float)peer.Top -
                            (float)shape.Top) <= 2f &&
                        Math.Abs((float)peer.Height -
                            (float)shape.Height) <= 5f)
                        .Select(peer => (float)peer.TextFrame.TextRange
                            .Font.Size).ToArray();
                    if (peers.Length == 0) continue;
                    var target = peers.Max();
                    if (target - before < 2f || target > 30f) continue;
                    dynamic range = shape.TextFrame.TextRange;
                    if ((float)range.BoundHeight * target / before >
                            (float)shape.Height + 1f ||
                        (float)range.BoundWidth * target / before >
                            (float)shape.Width + 1f)
                        continue;
                    targets[(int)shape.Id] = target;
                }
                foreach (dynamic shape in textShapes)
                {
                    float target;
                    if (!targets.TryGetValue((int)shape.Id, out target))
                        continue;
                    operations.Add(new Dictionary<string, object>
                    {
                        { "kind", "shape_font_size" },
                        { "slide_id", (int)slide.SlideID },
                        { "fingerprint", fingerprint },
                        { "shape_id", (int)shape.Id },
                        { "before_size", (float)shape.TextFrame.TextRange
                            .Font.Size },
                        { "size", target }
                    });
                }
            }
            return operations.ToArray();
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
