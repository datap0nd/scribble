using System;
using System.Collections.Generic;
using System.Linq;

namespace Scribble.Office
{
    // The source deck defines the repair's scope and structure. The model may
    // supply wording for these named slots; it never supplies an operation,
    // shape, layout, chart, position, colour, or source fingerprint.
    internal sealed class PilotRepairSkeleton
    {
        private sealed class Slot
        {
            internal string Name, Original, Field;
            internal int SlideId, ShapeId;
            internal bool Replacement;
        }

        private sealed class Page
        {
            internal int SlideId;
            internal Dictionary<string, object> Content;
        }

        private readonly List<Slot> _slots = new List<Slot>();
        private readonly List<Page> _pages = new List<Page>();
        private readonly object _source;

        private PilotRepairSkeleton(object source) { _source = source; }

        internal string[] SlotNames => _slots.Select(slot => slot.Name)
            .ToArray();

        internal object[] PublicSlots() => _slots.Select(slot => (object)new {
            name = slot.Name,
            current_wording = slot.Original,
            role = slot.Field,
            max_characters = slot.Field == "subtitle"
                ? PresentationDraftWriter.MaxSubtitleCharacters
                : PresentationDraftWriter.MaxTitleCharacters
        }).ToArray();

        internal static PilotRepairSkeleton Build(object source,
            IEnumerable<int> measuredReplacementSlideIds,
            IEnumerable<PresentationDraftCopy.MonthlyChartBinding> charts)
        {
            if (source == null || measuredReplacementSlideIds == null ||
                charts == null)
                throw new InvalidOperationException(
                    "PILOT_SKELETON_CONTEXT_MISSING");
            var measured = new HashSet<int>(
                measuredReplacementSlideIds);
            var chartSlides = new HashSet<int>(charts.Select(chart =>
                chart.SourceSlideId));
            if (measured.Overlaps(chartSlides))
                throw new InvalidOperationException(
                    "PILOT_SKELETON_CHART_SLIDE_UNSUPPORTED");
            var result = new PilotRepairSkeleton(source);
            dynamic deck = source;
            var pageHeight = (float)deck.PageSetup.SlideHeight;
            for (var index = 1; index <= (int)deck.Slides.Count; index++)
            {
                dynamic slide = deck.Slides[index];
                var id = (int)slide.SlideID;
                var replacement = measured.Contains(id);
                var text = new List<Tuple<int, float, float, string>>();
                var tables = new List<Dictionary<string, object>>();
                for (var shapeIndex = 1; shapeIndex <=
                    (int)slide.Shapes.Count; shapeIndex++)
                {
                    dynamic shape = slide.Shapes[shapeIndex];
                    if ((int)shape.HasChart != 0) continue;
                    if ((int)shape.HasTable != 0 && replacement)
                    {
                        tables.Add(ReadTable(shape));
                        continue;
                    }
                    if ((int)shape.HasTextFrame == 0 ||
                        (int)shape.TextFrame.HasText == 0 ||
                        (float)shape.Top >= pageHeight * .90f)
                        continue;
                    var wording = Convert.ToString(
                        shape.TextFrame.TextRange.Text)?.Trim();
                    if (string.IsNullOrWhiteSpace(wording)) continue;
                    text.Add(Tuple.Create((int)shape.Id,
                        (float)shape.Top, (float)shape.Left, wording));
                }
                text = text.OrderBy(item => item.Item2)
                    .ThenBy(item => item.Item3).ToList();
                if (replacement)
                {
                    var content = BuildReplacement(text, tables);
                    result._pages.Add(new Page {
                        SlideId = id, Content = content
                    });
                }
                for (var textIndex = 0; textIndex < Math.Min(2,
                    text.Count); textIndex++)
                {
                    var role = textIndex == 0 ? "title" : "subtitle";
                    result._slots.Add(new Slot {
                        Name = "slide_" + id + "_" + role,
                        Original = text[textIndex].Item4,
                        SlideId = id,
                        ShapeId = text[textIndex].Item1,
                        Replacement = replacement,
                        Field = role
                    });
                }
            }
            if (result._slots.Count == 0 && result._pages.Count == 0)
                throw new InvalidOperationException(
                    "PILOT_SKELETON_NO_EDITABLE_CONTENT");
            return result;
        }

        // Existing content is the safe default. A long or unbound model value
        // is discarded, so an unfixable validator proposal cannot consume the
        // request budget or introduce an unsupported data claim.
        internal object[] Operations(IDictionary<string, object> supplied,
            RevisionFactCatalog facts, out int ignoredSlotCount)
        {
            if (facts == null)
                throw new InvalidOperationException(
                    "PILOT_SKELETON_FACTS_MISSING");
            var values = supplied ?? new Dictionary<string, object>();
            var available = new HashSet<string>(SlotNames,
                StringComparer.Ordinal);
            ignoredSlotCount = values.Keys.Count(key =>
                !available.Contains(key));
            var edits = new Dictionary<string, string>(
                StringComparer.Ordinal);
            foreach (var slot in _slots)
            {
                object candidate;
                if (!values.TryGetValue(slot.Name, out candidate) ||
                    !(candidate is string) ||
                    string.IsNullOrWhiteSpace((string)candidate))
                    continue;
                try
                {
                    var rendered = facts.Render((string)candidate).Trim();
                    var limit = slot.Field == "subtitle"
                        ? PresentationDraftWriter.MaxSubtitleCharacters
                        : PresentationDraftWriter.MaxTitleCharacters;
                    if (rendered.Length > 0 && rendered.Length <= limit)
                        edits[slot.Name] = rendered;
                    else ignoredSlotCount++;
                }
                catch (InvalidOperationException error) when (
                    error.Message.StartsWith("REVISION_FACT_",
                        StringComparison.Ordinal))
                {
                    ignoredSlotCount++;
                }
            }
            var operations = new List<object>();
            foreach (var page in _pages)
            {
                var slideContent = new Dictionary<string, object>(
                    page.Content, StringComparer.Ordinal);
                foreach (var slot in _slots.Where(slot =>
                    slot.Replacement && slot.SlideId == page.SlideId))
                {
                    string value;
                    if (edits.TryGetValue(slot.Name, out value))
                        slideContent[slot.Field] = value;
                }
                var original = PresentationInspection.FindSlide(_source,
                    page.SlideId);
                operations.Add(new Dictionary<string, object> {
                    { "kind", "replace_slide" },
                    { "slide_id", page.SlideId },
                    { "fingerprint", PresentationInspection.Fingerprint(
                        original) },
                    { "slide", slideContent }
                });
            }
            foreach (var slot in _slots.Where(slot =>
                !slot.Replacement))
            {
                string value;
                if (!edits.TryGetValue(slot.Name, out value)) continue;
                var original = PresentationInspection.FindSlide(_source,
                    slot.SlideId);
                operations.Add(new Dictionary<string, object> {
                    { "kind", "replace_text" },
                    { "slide_id", slot.SlideId },
                    { "fingerprint", PresentationInspection.Fingerprint(
                        original) },
                    { "shape_id", slot.ShapeId },
                    { "before", slot.Original },
                    { "text", value }
                });
            }
            if (operations.Count == 0)
            {
                var slot = _slots.FirstOrDefault(item =>
                    !item.Replacement);
                if (slot == null)
                    throw new InvalidOperationException(
                        "PILOT_SKELETON_NO_OPERATION");
                var original = PresentationInspection.FindSlide(_source,
                    slot.SlideId);
                operations.Add(new Dictionary<string, object> {
                    { "kind", "replace_text" },
                    { "slide_id", slot.SlideId },
                    { "fingerprint", PresentationInspection.Fingerprint(
                        original) },
                    { "shape_id", slot.ShapeId },
                    { "before", slot.Original },
                    { "text", slot.Original }
                });
            }
            return operations.ToArray();
        }

        private static Dictionary<string, object> BuildReplacement(
            List<Tuple<int, float, float, string>> text,
            List<Dictionary<string, object>> tables)
        {
            var title = text.Count > 0 ? text[0].Item4 :
                "Source review";
            var subtitle = text.Count > 1 ? text[1].Item4 :
                string.Empty;
            if (title.Length > PresentationDraftWriter
                    .MaxTitleCharacters ||
                subtitle.Length > PresentationDraftWriter
                    .MaxSubtitleCharacters || tables.Count > 2)
                throw new InvalidOperationException(
                    "PILOT_SKELETON_CAPACITY");
            var result = new Dictionary<string, object> {
                { "title", title }, { "subtitle", subtitle }
            };
            if (tables.Count > 0)
            {
                result["layout"] = tables.Count == 1
                    ? "table" : "two_pane";
                result["table"] = tables[0];
                if (tables.Count == 2)
                    result["secondary_table"] = tables[1];
                var body = text.Skip(2).Select(item => item.Item4)
                    .ToArray();
                if (body.Length > PresentationDraftWriter
                        .MaxBulletsPerSlide ||
                    body.Any(line => line.Length > PresentationDraftWriter
                        .MaxBulletCharacters))
                    throw new InvalidOperationException(
                        "PILOT_SKELETON_CAPACITY");
                result["bullets"] = body;
            }
            else
            {
                var body = text.Skip(2).Select(item => item.Item4)
                    .ToArray();
                if (body.Length > PresentationDraftWriter.MaxCards *
                        PresentationDraftWriter.MaxCardPoints ||
                    body.Any(line => line.Length > PresentationDraftWriter
                        .MaxBulletCharacters))
                    throw new InvalidOperationException(
                        "PILOT_SKELETON_CAPACITY");
                result["layout"] = "cards";
                var groups = Math.Min(PresentationDraftWriter.MaxCards,
                    Math.Max(1, (body.Length + 3) / 4));
                var cards = new List<object>();
                for (var group = 0; group < groups; group++)
                {
                    var points = body.Where((line, index) =>
                        index % groups == group).ToArray();
                    cards.Add(new Dictionary<string, object> {
                        { "heading", "Source content " + (group + 1) },
                        { "points", points }
                    });
                }
                result["cards"] = cards.ToArray();
            }
            return result;
        }

        private static Dictionary<string, object> ReadTable(dynamic shape)
        {
            dynamic table = shape.Table;
            var rows = (int)table.Rows.Count;
            var columns = (int)table.Columns.Count;
            if (rows < 1 || rows > PresentationDraftWriter
                    .MaxTableRows + 1 ||
                columns < 1 || columns > PresentationDraftWriter
                    .MaxTableColumns)
                throw new InvalidOperationException(
                    "PILOT_SKELETON_TABLE_CAPACITY");
            var data = new List<string[]>();
            for (var row = 1; row <= rows; row++)
            {
                var cells = new List<string>();
                for (var column = 1; column <= columns; column++)
                {
                    var value = Convert.ToString(table.Cell(row,
                        column).Shape.TextFrame.TextRange.Text) ?? "";
                    if (value.Length > PresentationDraftWriter
                            .MaxCellCharacters)
                        throw new InvalidOperationException(
                            "PILOT_SKELETON_TABLE_CAPACITY");
                    cells.Add(value);
                }
                data.Add(cells.ToArray());
            }
            return new Dictionary<string, object> {
                { "headers", data[0] },
                { "rows", data.Skip(1).ToArray() }
            };
        }
    }
}
