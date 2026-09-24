using System;
using System.Collections.Generic;
using System.Linq;

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
            try
            {
                draft = app.Presentations.Add(0);
                draft.Tags.Add("ScribbleRevisionDraft", owner);
                draft.PageSetup.SlideWidth = source.PageSetup.SlideWidth;
                draft.PageSetup.SlideHeight = source.PageSetup.SlideHeight;
                var result = new PresentationDraftCopy(sourcePresentation,
                    (object)draft, order);
                for (var index = 1; index <= 6; index++)
                {
                    dynamic original = source.Slides[index];
                    var originalId = (int)original.SlideID;
                    var fingerprint = PresentationInspection
                        .CopyContentFingerprint((object)original);
                    result._sourceContent[originalId] = fingerprint;
                    original.Copy();
                    draft.Slides.Paste(index);
                    if ((int)draft.Slides.Count != index)
                        throw new InvalidOperationException(
                            "REVISION_COPY_INCOMPLETE: Native paste changed the page count.");
                    dynamic copy = draft.Slides[index];
                    PresentationInspection.RestoreCopiedBackground(
                        (object)original, (object)copy);
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
                result.VerifySource();
                return result;
            }
            catch
            {
                if (draft != null) try { draft.Close(); } catch { }
                throw;
            }
        }

        internal object[] BindOperations(object[] operations)
        {
            VerifySource();
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

        internal void VerifySource()
        {
            dynamic source = Source;
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
