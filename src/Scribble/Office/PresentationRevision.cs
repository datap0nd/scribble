using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Web.Script.Serialization;

namespace Scribble.Office
{
    // Native-only transaction. No disk deck, save, arbitrary code or external data refresh.
    internal sealed partial class PresentationRevision
    {
        internal sealed class Item
        {
            internal int SlideId, Index;
            internal string Before, After;
            internal object Original, Staged, Backup;
            internal List<Dictionary<string, object>> Operations = new List<Dictionary<string, object>>();
            internal bool Started, Applied, Deleted;
            internal string LastKnownContent;
            internal string BackupFingerprint;
            internal readonly List<int> InsertedIds = new List<int>();
            internal readonly List<int> AddedShapeIds = new List<int>();
            internal readonly List<object> StagedInserts = new List<object>();
            internal readonly List<PresentationDraftWriter.SamsungOutput> StagedInsertOutputs = new List<PresentationDraftWriter.SamsungOutput>();
            internal readonly Dictionary<int, string> InsertedFingerprints = new Dictionary<int, string>();
        }
        private static readonly ConditionalWeakTable<object, PresentationRevision> Latest = new ConditionalWeakTable<object, PresentationRevision>();
        internal readonly object Presentation;
        internal readonly List<Item> Items = new List<Item>();
        internal object Working, Recovery;
        internal string Status = "preparing";
        private int[] _beforeOrder, _afterOrder;
        private Dictionary<int, string> _beforeFingerprints;
        private bool _structuralStarted;
        private int[] _lastKnownOrder;
        private int[] Order()
        {
            dynamic deck = Presentation; var ids = new List<int>();
            for (var i = 1; i <= (int)deck.Slides.Count; i++) ids.Add((int)deck.Slides[i].SlideID);
            return ids.ToArray();
        }
        internal PresentationRevision(object presentation) { Presentation = presentation; }

        internal void Stage(object application, object[] operations)
        {
            if (operations.Length == 0 || operations.Length > 24) throw new InvalidOperationException("A revision batch needs 1 to 24 operations.");
            dynamic deck = Presentation;
            _beforeOrder = Order(); _lastKnownOrder = _beforeOrder;
            _beforeFingerprints = _beforeOrder.ToDictionary(id => id, id => PresentationInspection.Fingerprint(PresentationInspection.FindSlide(Presentation, id)));
            foreach (var raw in operations)
            {
                var operation = SamsungAuthoringPolicy.ReadMap(raw);
                var id = Convert.ToInt32(operation["slide_id"]);
                var item = Items.FirstOrDefault(i => i.SlideId == id);
                if (item == null)
                {
                    var slide = PresentationInspection.FindSlide(Presentation, id);
                    item = new Item { SlideId = id, Index = (int)((dynamic)slide).SlideIndex, Original = slide,
                        Before = PresentationInspection.Fingerprint(slide), LastKnownContent = PresentationInspection.ContentFingerprint(slide) };
                    Items.Add(item);
                }
                if (item.Before != SamsungAuthoringPolicy.Text(operation, "fingerprint"))
                    throw new InvalidOperationException("SLIDE_CHANGED: Read the original slide again before editing.");
                ValidateOperation(item.Original, operation);
                if (item.Operations.Any(o => SamsungAuthoringPolicy.Text(o, "kind") == "delete") ||
                    (SamsungAuthoringPolicy.Text(operation, "kind") == "delete" && item.Operations.Count > 0))
                    throw new InvalidOperationException("REVISION_DELETE_CONFLICT: Delete cannot be combined with edits to the same slide.");
                if (SamsungAuthoringPolicy.Text(operation, "kind") == "move" &&
                    (Convert.ToInt32(operation["new_index"]) < 1 || Convert.ToInt32(operation["new_index"]) > (int)deck.Slides.Count))
                    throw new InvalidOperationException("REVISION_POSITION_INVALID");
                if (item.Operations.Count > 0 && (SamsungAuthoringPolicy.Text(operation, "kind") == "replace_slide" ||
                    item.Operations.Any(o => SamsungAuthoringPolicy.Text(o, "kind") == "replace_slide")))
                    throw new InvalidOperationException("REVISION_RECOMPOSE_CONFLICT: Recompose a slide in a separate batch from targeted edits.");
                item.Operations.Add(operation);
            }
            _unrelatedContent = _beforeOrder.Where(id => Items.All(i => i.SlideId != id)).ToDictionary(id => id, id => PresentationInspection.ContentFingerprint(PresentationInspection.FindSlide(Presentation, id)));
            dynamic app = application;
            Working = app.Presentations.Add(0);
            Recovery = app.Presentations.Add(0);
            ((dynamic)Recovery).Tags.Add("ScribbleRevisionRecovery", BatchId);
            foreach (var target in new[] { Working, Recovery })
            { dynamic temp = target; temp.PageSetup.SlideWidth = deck.PageSetup.SlideWidth; temp.PageSetup.SlideHeight = deck.PageSetup.SlideHeight; }
            foreach (var item in Items)
            {
                item.Staged = CopySlide(item.Original, Working);
                item.Backup = CopySlide(item.Original, Recovery);
                item.BackupFingerprint = PresentationInspection.Fingerprint(item.Backup);
                foreach (var operation in item.Operations)
                {
                    if (SamsungAuthoringPolicy.Text(operation, "kind") == "insert")
                    {
                        dynamic working = Working; dynamic inserted = working.Slides.Add((int)working.Slides.Count + 1, 12);
                        item.StagedInsertOutputs.Add(DrawReplacement((object)inserted, operation, false)); ValidateNativeGeometry((object)inserted);
                        item.StagedInserts.Add((object)inserted);
                    }
                    else Apply(item.Original, item.Staged, operation);
                }
                ValidateNativeGeometry(item.Staged);
                var serializer = new JavaScriptSerializer();
                if (serializer.Serialize(PresentationInspection.Hyperlinks(item.Original)) != serializer.Serialize(PresentationInspection.Hyperlinks(item.Staged)))
                    throw new InvalidOperationException("REVISION_PRESERVATION: The proposed edit changed existing hyperlinks.");
            }
            var proposedOrder = ReviewedSlides();
            foreach (var item in Items)
                foreach (var output in item.StagedInsertOutputs) PresentationDraftWriter.SetSamsungPageNumber(output, proposedOrder.IndexOf(output.Slide) + 1);
            Status = "staged";
        }
        internal List<object> ReviewedSlides()
        {
            var proposed = new List<object>();
            foreach (var id in _beforeOrder)
            {
                var item = Items.FirstOrDefault(v => v.SlideId == id);
                proposed.Add(item == null ? PresentationInspection.FindSlide(Presentation, id) : item.Staged);
            }
            foreach (var item in Items)
            {
                var inserted = 0;
                foreach (var operation in item.Operations)
                {
                    var kind = SamsungAuthoringPolicy.Text(operation, "kind");
                    if (kind == "move")
                    {
                        proposed.Remove(item.Staged);
                        proposed.Insert(Convert.ToInt32(operation["new_index"]) - 1, item.Staged);
                    }
                    if (kind == "insert")
                    {
                        proposed.Insert(proposed.IndexOf(item.Staged) + inserted + 1, item.StagedInserts[inserted]);
                        inserted++;
                    }
                    if (kind == "delete") proposed.Remove(item.Staged);
                }
            }
            return proposed;
        }
        private static object CopySlide(object slide, object target)
        {
            dynamic original = slide; dynamic deck = target;
            original.Copy();
            dynamic pasted = deck.Slides.Paste((int)deck.Slides.Count + 1);
            object copy = pasted[1];
            if (PresentationInspection.ContentFingerprint(slide) != PresentationInspection.ContentFingerprint(copy))
                throw new InvalidOperationException("REVISION_COPY_PRESERVATION: Native staging did not preserve the source content and formatting.");
            return copy;
        }
        private static object CorrespondingShape(object original, object target, int id)
        {
            if (ReferenceEquals(original, target)) return PresentationInspection.FindShape(original, id);
            dynamic source = original; dynamic destination = target;
            var mapped = MapShape((object)source.Shapes, (object)destination.Shapes, id);
            return mapped ?? throw new InvalidOperationException("SLIDE_SHAPE_MAPPING_FAILED");
        }
        private static object MapShape(object source, object target, int id)
        {
            dynamic from = source; dynamic to = target;
            if ((int)from.Count > (int)to.Count) throw new InvalidOperationException("SLIDE_SHAPE_MAPPING_CHANGED");
            for (var i = 1; i <= (int)from.Count; i++)
            {
                dynamic a = from[i]; dynamic b = to[i];
                if ((int)a.Type != (int)b.Type) throw new InvalidOperationException("SLIDE_SHAPE_MAPPING_CHANGED");
                if ((int)a.Id == id) return b;
                if ((int)a.Type == 6)
                { var result = MapShape((object)a.GroupItems, (object)b.GroupItems, id); if (result != null) return result; }
            }
            return null;
        }
        private static void ValidateOperation(object slide, Dictionary<string, object> operation)
        {
            var kind = SamsungAuthoringPolicy.Text(operation, "kind");
            if (!new[] { "replace_text", "table_cell", "chart_point", "move", "delete", "replace_slide", "insert", "annotate", "notes_append" }.Contains(kind))
                throw new InvalidOperationException("REVISION_UNSUPPORTED: Use a supported targeted operation.");
            if (kind == "move" || kind == "notes_append") return;
            if (kind == "delete")
            {
                dynamic page = slide; dynamic deck = page.Parent;
                if ((int)deck.SlideShowSettings.NamedSlideShows.Count > 0)
                    throw new InvalidOperationException("REVISION_PRESERVATION: Slide deletion needs explicit handling of existing custom slide shows.");
                for (var i = 1; i <= (int)deck.Slides.Count; i++)
                {
                    dynamic links = deck.Slides[i].Hyperlinks;
                    for (var n = 1; n <= (int)links.Count; n++)
                    {
                        var subaddress = Convert.ToString(links[n].SubAddress) ?? "";
                        if (subaddress.Split(',')[0] == Convert.ToString(page.SlideID))
                            throw new InvalidOperationException("REVISION_PRESERVATION: Another slide links to the requested deletion target. Preserve or explicitly revise that link first.");
                    }
                }
                return;
            }
            if (kind == "replace_slide" || kind == "insert")
            {
                var content = SamsungAuthoringPolicy.ReadMap(operation["slide"]);
                var parsed = PresentationDraftWriter.ParseSlides(new object[] { content });
                if (PresentationDraftWriter.ComposeSamsung(parsed).Count != 1) throw new InvalidOperationException("REVISION_SLIDE_COUNT: Supply exactly one page per operation.");
                dynamic page = slide; dynamic deck = page.Parent;
                if (Math.Abs((double)deck.PageSetup.SlideWidth / (double)deck.PageSetup.SlideHeight - 16.0 / 9) > .01)
                    throw new InvalidOperationException("REVISION_CANVAS: Recomposition needs a separate 16:9 deck; targeted edits preserve this canvas.");
                if (kind == "replace_slide")
                {
                    if ((int)page.Background.Fill.Type != 1)
                        throw new InvalidOperationException("REVISION_PRESERVATION: Reconstruction supports solid backgrounds only. Use targeted edits to preserve this background.");
                    if ((int)page.TimeLine.MainSequence.Count > 0 || (int)page.TimeLine.InteractiveSequences.Count > 0 || (int)page.Hyperlinks.Count > 0)
                        throw new InvalidOperationException("REVISION_PRESERVATION: Cannot reconstruct a slide with animations.");
                    for (var i = 1; i <= (int)page.Shapes.Count; i++)
                    {
                        dynamic existing = page.Shapes[i];
                        if (!new[] { 1, 14, 17, 19, 3 }.Contains((int)existing.Type) || HasActions((object)existing))
                            throw new InvalidOperationException("REVISION_PRESERVATION: Reconstruction would lose artwork, groups, media or actions. Use targeted edits.");
                    }
                }
                return;
            }
            dynamic shape = PresentationInspection.FindShape(slide, Convert.ToInt32(operation["shape_id"]));
            if (kind == "replace_text")
            {
                if ((int)shape.HasTextFrame == 0 || string.IsNullOrEmpty(SamsungAuthoringPolicy.Text(operation, "before")))
                    throw new InvalidOperationException("REVISION_TEXT_REQUIRED");
                var text = Convert.ToString(shape.TextFrame.TextRange.Text);
                var before = SamsungAuthoringPolicy.Text(operation, "before");
                if (text.IndexOf(before, StringComparison.Ordinal) < 0 || text.IndexOf(before, StringComparison.Ordinal) != text.LastIndexOf(before, StringComparison.Ordinal))
                    throw new InvalidOperationException("REVISION_TEXT_AMBIGUOUS: Select a unique exact text span.");
            }
            if (kind == "table_cell" && (int)shape.HasTable == 0) throw new InvalidOperationException("REVISION_TABLE_REQUIRED");
            if (kind == "annotate" && (int)shape.HasTable == 0 && (int)shape.HasChart == 0)
                throw new InvalidOperationException("REVISION_ANNOTATION_TARGET: Select a table or chart.");
            if (kind == "chart_point" && ((int)shape.HasChart == 0 || (bool)shape.Chart.ChartData.IsLinked))
                throw new InvalidOperationException("REVISION_CHART_UNSUPPORTED: Linked charts cannot be refreshed or changed.");
        }
        private static bool HasActions(object shape)
        {
            dynamic target = shape;
            return (int)target.ActionSettings[1].Action != 0 || (int)target.ActionSettings[2].Action != 0;
        }
        private static PresentationDraftWriter.SamsungOutput DrawReplacement(object target, Dictionary<string, object> operation, bool clear, int? displaySlideNumber = null)
        {
            dynamic page = target;
            if (clear) for (var i = (int)page.Shapes.Count; i >= 1; i--) page.Shapes[i].Delete();
            var slide = PresentationDraftWriter.ParseSlides(new object[] { operation["slide"] });
            var composed = PresentationDraftWriter.ComposeSamsung(slide);
            dynamic deck = page.Parent;
            var scale = (float)deck.PageSetup.SlideWidth / SamsungSlideDesign.Width;
            if (Math.Abs(scale - 1) > .001) PresentationDraftWriter.ScaleSamsungPage(composed[0], scale);
            return PresentationDraftWriter.DrawSamsungPage(target, composed[0], Guid.NewGuid().ToString("N"), displaySlideNumber);
        }
        private static void Apply(object original, object target, Dictionary<string, object> operation)
        {
            var kind = SamsungAuthoringPolicy.Text(operation, "kind");
            if (kind == "move" || kind == "delete" || kind == "insert") return; // Structural commit occurs after content review.
            if (kind == "replace_slide") { DrawReplacement(target, operation, true, (int)((dynamic)original).SlideIndex); return; }
            if (kind == "notes_append")
            {
                dynamic page = target; dynamic notes = page.NotesPage.Shapes.Placeholders[2].TextFrame.TextRange;
                notes.InsertAfter("\n" + SamsungAuthoringPolicy.Text(operation, "notes")); return;
            }
            dynamic shape = CorrespondingShape(original, target, Convert.ToInt32(operation["shape_id"]));
            if (kind == "annotate")
            {
                dynamic page = target;
                float x = shape.Left, y = shape.Top, width = shape.Width, height = shape.Height;
                var row = Convert.ToInt32(operation["row"]);
                if ((int)shape.HasTable != 0)
                {
                    var col = operation.ContainsKey("column") ? Convert.ToInt32(operation["column"]) : 1;
                    dynamic cell = shape.Table.Cell(row, col).Shape;
                    x = cell.Left; y = cell.Top; height = cell.Height;
                    if (operation.ContainsKey("column")) width = cell.Width;
                }
                else
                {
                    var series = operation.ContainsKey("series") ? Convert.ToInt32(operation["series"]) : 1;
                    dynamic point = shape.Chart.SeriesCollection(series).Points(row);
                    x += (float)point.Left; y += (float)point.Top; width = Math.Max(4, (float)point.Width); height = Math.Max(4, (float)point.Height);
                }
                dynamic frame = page.Shapes.AddShape(1, x, y, width, height);
                frame.Fill.Visible = 0; frame.Line.ForeColor.RGB = MetoTheme.Rgb(SamsungSlideDesign.Red); frame.Line.Weight = 1;
                return;
            }
            if (kind == "replace_text")
            {
                dynamic range = shape.TextFrame.TextRange;
                var text = Convert.ToString(range.Text);
                var before = SamsungAuthoringPolicy.Text(operation, "before");
                var start = text.IndexOf(before, StringComparison.Ordinal);
                if (start < 0) throw new InvalidOperationException("REVISION_TEXT_CHANGED");
                range.Characters(start + 1, before.Length).Text = SamsungAuthoringPolicy.Text(operation, "text");
            }
            else if (kind == "table_cell")
            {
                dynamic range = shape.Table.Cell(Convert.ToInt32(operation["row"]), Convert.ToInt32(operation["column"])).Shape.TextFrame.TextRange;
                if (Convert.ToString(range.Text) != SamsungAuthoringPolicy.Text(operation, "before")) throw new InvalidOperationException("REVISION_CELL_CHANGED");
                range.Text = SamsungAuthoringPolicy.Text(operation, "text");
            }
            else if (kind == "chart_point")
            {
                PresentationChartEdit.SetPoint((object)shape.Chart, Convert.ToInt32(operation["series"]), Convert.ToInt32(operation["category"]),
                    Convert.ToDouble(operation["before_value"]), Convert.ToDouble(operation["value"]));
            }
        }
        internal static void ValidateNativeGeometry(object slide)
        {
            dynamic page = slide; dynamic deck = page.Parent;
            for (var i = 1; i <= (int)page.Shapes.Count; i++)
            {
                dynamic shape = page.Shapes[i];
                if ((float)shape.Left < -.5 || (float)shape.Top < -.5 || (float)shape.Left + (float)shape.Width > (float)deck.PageSetup.SlideWidth + .5 ||
                    (float)shape.Top + (float)shape.Height > (float)deck.PageSetup.SlideHeight + .5)
                    throw new InvalidOperationException("SLIDE_NATIVE_OVERFLOW: Shape exceeds canvas.");
                if ((int)shape.HasTable != 0)
                {
                    dynamic table = shape.Table;
                    for (var row = 1; row <= (int)table.Rows.Count; row++)
                    for (var col = 1; col <= (int)table.Columns.Count; col++)
                    {
                        dynamic cell = table.Cell(row, col).Shape; dynamic text = cell.TextFrame.TextRange;
                        if ((float)text.BoundHeight > (float)cell.Height + 1 || (float)text.BoundWidth > (float)cell.Width + 1)
                            throw new InvalidOperationException("SLIDE_NATIVE_OVERFLOW: Table cell text does not fit.");
                    }
                }
                if ((int)shape.HasChart != 0)
                {
                    dynamic chart = shape.Chart;
                    if ((bool)chart.HasTitle) CheckChartLabel((object)chart.ChartTitle, (float)shape.Width, (float)shape.Height);
                    for (var series = 1; series <= (int)chart.SeriesCollection().Count; series++)
                    {
                        dynamic points = chart.SeriesCollection(series).Points();
                        for (var n = 1; n <= (int)points.Count; n++)
                            if ((bool)points[n].HasDataLabel) CheckChartLabel((object)points[n].DataLabel, (float)shape.Width, (float)shape.Height);
                    }
                }
                if ((int)shape.HasTextFrame != 0 && (int)shape.HasTable == 0 && (int)shape.HasChart == 0)
                {
                    dynamic range = shape.TextFrame.TextRange;
                    if ((float)range.BoundHeight > (float)shape.Height + 1 || (float)range.BoundWidth > (float)shape.Width + 1)
                        throw new InvalidOperationException("SLIDE_NATIVE_OVERFLOW: Text does not fit.");
                }
            }
        }
        private static void CheckChartLabel(object value, float width, float height)
        {
            dynamic label = value;
            if ((float)label.Left < -1 || (float)label.Top < -1 || (float)label.Left + (float)label.Width > width + 1 ||
                (float)label.Top + (float)label.Height > height + 1)
                throw new InvalidOperationException("SLIDE_NATIVE_OVERFLOW: Chart label exceeds the chart frame.");
        }
        private static void RestoreContent(Item item)
        {
            foreach (var id in item.AddedShapeIds) { dynamic added = PresentationInspection.FindShape(item.Original, id); added.Delete(); }
            item.AddedShapeIds.Clear();
            foreach (var operation in item.Operations)
            {
                var kind = SamsungAuthoringPolicy.Text(operation, "kind");
                if (kind == "move" || kind == "delete" || kind == "insert" || kind == "annotate") continue;
                if (kind == "replace_slide")
                {
                    dynamic original = item.Original; dynamic backup = item.Backup;
                    for (var i = (int)original.Shapes.Count; i >= 1; i--) original.Shapes[i].Delete();
                    if ((int)backup.Shapes.Count > 0) { backup.Shapes.Range().Copy(); original.Shapes.Paste(); }
                    original.FollowMasterBackground = 0; original.Background.Fill.Solid();
                    original.Background.Fill.ForeColor.RGB = backup.Background.Fill.ForeColor.RGB;
                    original.Background.Fill.Transparency = backup.Background.Fill.Transparency;
                    original.FollowMasterBackground = backup.FollowMasterBackground;
                    backup.NotesPage.Shapes.Placeholders[2].TextFrame.TextRange.Copy(); original.NotesPage.Shapes.Placeholders[2].TextFrame.TextRange.PasteSpecial(9);
                    continue;
                }
                if (kind == "notes_append")
                {
                    dynamic original = item.Original; dynamic backup = item.Backup;
                    backup.NotesPage.Shapes.Placeholders[2].TextFrame.TextRange.Copy(); original.NotesPage.Shapes.Placeholders[2].TextFrame.TextRange.PasteSpecial(9); continue;
                }
                var id = Convert.ToInt32(operation["shape_id"]);
                dynamic source = CorrespondingShape(item.Original, item.Backup, id);
                dynamic target = PresentationInspection.FindShape(item.Original, id);
                if (kind == "replace_text")
                {
                    source.TextFrame.TextRange.Copy(); target.TextFrame.TextRange.PasteSpecial(9);
                }
                else if (kind == "table_cell")
                {
                    var row = Convert.ToInt32(operation["row"]); var column = Convert.ToInt32(operation["column"]);
                    source.Table.Cell(row, column).Shape.TextFrame.TextRange.Copy();
                    target.Table.Cell(row, column).Shape.TextFrame.TextRange.PasteSpecial(9);
                }
                else if (kind == "chart_point")
                {
                    var series = Convert.ToInt32(operation["series"]); var category = Convert.ToInt32(operation["category"]);
                    var before = ((System.Collections.IEnumerable)source.Chart.SeriesCollection(series).Values).Cast<object>().ToArray();
                    var current = ((System.Collections.IEnumerable)target.Chart.SeriesCollection(series).Values).Cast<object>().ToArray();
                    PresentationChartEdit.SetPoint((object)target.Chart, series, category, Convert.ToDouble(current[category - 1]), Convert.ToDouble(before[category - 1]));
                }
            }
        }
        internal void Commit(Action<string> journal)
        {
            foreach (var item in Items)
                if (PresentationInspection.Fingerprint(item.Original) != item.Before) throw new InvalidOperationException("SLIDE_CHANGED_DURING_REVIEW");
            if (!Order().SequenceEqual(_beforeOrder)) throw new InvalidOperationException("SLIDE_ORDER_CHANGED_DURING_REVIEW");
            VerifyUnrelated(); VerifyRecoveryOriginals();
            Status = "applying"; journal(Status);
            try
            {
                foreach (var item in Items)
                {
                    item.Started = true; journal("applying:" + item.SlideId);
                    foreach (var operation in item.Operations)
                    {
                        dynamic live = item.Original;
                        var oldCount = (int)live.Shapes.Count;
                        Apply(item.Original, item.Original, operation);
                        item.LastKnownContent = PresentationInspection.ContentFingerprint(item.Original);
                        if (SamsungAuthoringPolicy.Text(operation, "kind") == "annotate")
                            for (var n = oldCount + 1; n <= (int)live.Shapes.Count; n++) item.AddedShapeIds.Add((int)live.Shapes[n].Id);
                        journal("operation_applied:" + item.SlideId);
                    }
                    ValidateNativeGeometry(item.Original);
                    if (PresentationInspection.ContentFingerprint(item.Original) != PresentationInspection.ContentFingerprint(item.Staged))
                        throw new InvalidOperationException("REVISION_LIVE_MISMATCH: The live result differs from the reviewed staging slide.");
                    item.After = PresentationInspection.Fingerprint(item.Original); item.Applied = true;
                    journal("applied:" + item.SlideId);
                }
                foreach (var item in Items)
                    foreach (var operation in item.Operations)
                    {
                        var kind = SamsungAuthoringPolicy.Text(operation, "kind");
                        if (kind == "move" || kind == "insert" || kind == "delete") _structuralStarted = true;
                        if (kind == "move") { dynamic slide = item.Original; slide.MoveTo(Convert.ToInt32(operation["new_index"])); }
                        if (kind == "insert")
                        {
                            dynamic anchor = item.Original;
                            var stagedIndex = item.InsertedIds.Count;
                            dynamic inserted = CopySlide(item.StagedInserts[stagedIndex], Presentation);
                            item.InsertedIds.Add((int)inserted.SlideID);
                            inserted.MoveTo((int)anchor.SlideIndex + item.InsertedIds.Count);
                            ValidateNativeGeometry((object)inserted);
                            item.InsertedFingerprints[(int)inserted.SlideID] = PresentationInspection.Fingerprint((object)inserted);
                        }
                        if (kind == "delete") { dynamic slide = item.Original; slide.Delete(); item.Deleted = true; }
                        _lastKnownOrder = Order();
                        journal("structure_applied:" + item.SlideId);
                    }
                foreach (var item in Items.Where(i => !i.Operations.Any(o => SamsungAuthoringPolicy.Text(o, "kind") == "delete")))
                    item.After = PresentationInspection.Fingerprint(item.Original);
                foreach (var item in Items)
                    foreach (var id in item.InsertedIds) item.InsertedFingerprints[id] = PresentationInspection.Fingerprint(PresentationInspection.FindSlide(Presentation, id));
                _afterOrder = Order();
                VerifyUnrelated();
                Status = "applied"; journal(Status);
                Latest.Remove(Presentation); Latest.Add(Presentation, this);
            }
            catch
            {
                try
                {
                    VerifyRecoveryOriginals();
                    if (!Order().SequenceEqual(_lastKnownOrder)) throw new InvalidOperationException("REVISION_ROLLBACK_ORDER_CONFLICT");
                    foreach (var item in Items.Where(i => i.Started && !i.Deleted))
                        if (PresentationInspection.ContentFingerprint(item.Original) != item.LastKnownContent)
                            throw new InvalidOperationException("REVISION_ROLLBACK_CONFLICT");
                    if (_structuralStarted) RestoreStructure();
                    foreach (var item in Items.Where(i => i.Started).Reverse()) RestoreContent(item);
                    VerifyRestored();
                    Status = "rolled_back"; journal(Status);
                }
                catch { Status = "recovery_required"; journal(Status); }
                throw;
            }
        }
        private void RestoreStructure()
        {
            foreach (var item in Items)
            {
                foreach (var id in item.InsertedIds)
                {
                    string expected;
                    if (!item.InsertedFingerprints.TryGetValue(id, out expected) || PresentationInspection.Fingerprint(PresentationInspection.FindSlide(Presentation, id)) != expected)
                        throw new InvalidOperationException("REVISION_INSERT_RECOVERY_CONFLICT");
                }
            }
            foreach (var item in Items)
            {
                foreach (var id in item.InsertedIds) { dynamic inserted = PresentationInspection.FindSlide(Presentation, id); inserted.Delete(); }
                item.InsertedIds.Clear(); item.InsertedFingerprints.Clear();
                if (item.Deleted) { item.Original = CopySlide(item.Backup, Presentation); item.Deleted = false; }
            }
            for (var i = 0; i < _beforeOrder.Length; i++)
            {
                var item = Items.FirstOrDefault(v => v.SlideId == _beforeOrder[i]);
                dynamic original = item == null ? PresentationInspection.FindSlide(Presentation, _beforeOrder[i]) : item.Original;
                original.MoveTo(i + 1);
            }
        }
        private void VerifyRestored()
        {
            foreach (var item in Items)
                if (PresentationInspection.ContentFingerprint(item.Original) != PresentationInspection.ContentFingerprint(item.Backup))
                    throw new InvalidOperationException("REVISION_RECOVERY_INCOMPLETE: The restored slide differs from its native original.");
        }
        private void VerifyUnrelated()
        {
            if (_beforeFingerprints == null) throw new InvalidOperationException("REVISION_DECK_RECEIPT_MISSING");
            foreach (var pair in _beforeFingerprints.Where(p => Items.All(i => i.SlideId != p.Key)))
            {
                // Slide indexes can change in an authorized structural operation;
                // compare native content at the original index-independent snapshot.
                var current = PresentationInspection.FindSlide(Presentation, pair.Key);
                var actual = PresentationInspection.ContentFingerprint(current);
                string expected;
                if (!_unrelatedContent.TryGetValue(pair.Key, out expected) || actual != expected) throw new InvalidOperationException("REVISION_UNRELATED_SLIDE_CHANGED: User edits were preserved; re-review the deck.");
            }
        }
        private Dictionary<int, string> _unrelatedContent = new Dictionary<int, string>();
        internal void CloseStaging(bool keepRecovery)
        {
            if (Working != null) { dynamic working = Working; working.Close(); Working = null; }
            if (!keepRecovery && Recovery != null) { dynamic recovery = Recovery; recovery.Close(); Recovery = null; }
            if (Status == "recovery_required" && Recovery != null) { dynamic recovery = Recovery; recovery.NewWindow(); }
        }
        internal static PresentationRevision Last(object presentation)
        { PresentationRevision value; return Latest.TryGetValue(presentation, out value) ? value : null; }
        private void VerifyRecoveryOriginals()
        {
            foreach (var item in Items)
                if (PresentationInspection.Fingerprint(item.Backup) != item.BackupFingerprint) throw new InvalidOperationException("REVISION_RECOVERY_ORIGINAL_CHANGED");
        }
        internal void Revert() { RevertWithJournal(null); }
        internal void RevertWithJournal(Action<string> journal)
        {
            VerifyRecoveryOriginals();
            if (Status != "applied") throw new InvalidOperationException("REVISION_RECOVERY_REQUIRED: Inspect the unsaved recovery presentation.");
            if (!Order().SequenceEqual(_afterOrder)) throw new InvalidOperationException("REVISION_REVERT_CONFLICT: The deck order changed after this revision.");
            foreach (var item in Items)
            {
                foreach (var inserted in item.InsertedFingerprints)
                    if (PresentationInspection.Fingerprint(PresentationInspection.FindSlide(Presentation, inserted.Key)) != inserted.Value)
                        throw new InvalidOperationException("REVISION_REVERT_CONFLICT: A user edited an inserted slide.");
                if (item.Operations.Any(o => SamsungAuthoringPolicy.Text(o, "kind") == "delete")) continue;
                if (PresentationInspection.Fingerprint(item.Original) != item.After) throw new InvalidOperationException("REVISION_REVERT_CONFLICT: Later user edits were preserved.");
            }
            try
            {
                Status = "reverting"; journal?.Invoke(Status);
                RestoreStructure(); _lastKnownOrder = Order();
                foreach (var item in Items) item.LastKnownContent = PresentationInspection.ContentFingerprint(item.Original);
                journal?.Invoke("structure_restored");
                foreach (var item in Items)
                {
                    RestoreContent(item); item.LastKnownContent = PresentationInspection.ContentFingerprint(item.Original);
                    journal?.Invoke("content_restored:" + item.SlideId);
                }
                VerifyRestored();
                Status = "reverted"; journal?.Invoke(Status);
            }
            catch { Status = "recovery_required"; throw; }
        }
    }
}
