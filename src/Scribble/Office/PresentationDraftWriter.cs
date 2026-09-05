using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Scribble.Security;

namespace Scribble.Office
{
    // The single PowerPoint write surface of the suite. Draft slides
    // are added to the presentation, every added slide carries the
    // [Scribble draft] marker, existing slides are never modified, and
    // the file is never saved - saving stays a human action.
    //
    // Every slide is built on a blank layout and painted entirely
    // from MetoTheme: the model supplies content (title, subtitle,
    // bullets, cards, table, chart, footnote) and the writer decides
    // the layout template, fonts, sizes, colors, and geometry. That
    // split is deliberate - a small local model cannot be trusted to
    // apply a corporate design system, but it can be trusted to
    // write the words.
    internal static partial class PresentationDraftWriter
    {
        internal const string DraftMarker = "[Scribble draft]";
        internal const int MaxDraftSlides = 10;
        internal const int MaxBulletsPerSlide = 12;
        internal const int MaxTitleCharacters = 200;
        internal const int MaxBulletCharacters = 300;
        internal const int MaxSubtitleCharacters = 180;
        internal const int MaxFootnoteCharacters = 200;
        internal const int MaxUnitCharacters = 40;

        internal const int MaxChartCategories = 24;
        internal const int MaxChartSeries = 5;

        internal const int MaxCards = 4;
        internal const int MaxCardPoints = 6;
        internal const int MaxTableRows = 200;
        internal const int MaxTableColumns = 8;
        internal const int MaxCellCharacters = 400;

        // PowerPoint enumeration values used through late binding.
        private const int PpLayoutBlank = 12;
        private const int MsoShapeRectangle = 1;
        private const int MsoTextOrientationHorizontal = 1;
        private const int PpAlignLeft = 1;
        private const int PpAlignCenter = 2;
        private const int PpAlignRight = 3;
        private const int MsoTrue = -1;
        private const int MsoFalse = 0;

        internal sealed class DraftSlide
        {
            internal DraftSlide(
                string title,
                string subtitle,
                string footnote,
                string unit,
                string layout,
                IReadOnlyList<DraftBullet> bullets,
                IReadOnlyList<DraftCard> cards,
                DraftTable table,
                DraftChart chart)
            {
                Title = title ?? string.Empty;
                Subtitle = subtitle ?? string.Empty;
                Footnote = footnote ?? string.Empty;
                Unit = unit ?? string.Empty;
                Bullets = bullets ?? new DraftBullet[0];
                Cards = cards ?? new DraftCard[0];
                Table = table;
                Chart = chart;
                Layout = MetoTheme.ResolveLayout(
                    layout,
                    Bullets.Count > 0,
                    Cards.Count > 0,
                    table != null,
                    chart != null);
            }

            internal string Title { get; }

            internal string Subtitle { get; }

            internal string Footnote { get; }

            internal string Unit { get; }

            internal string Layout { get; }

            internal IReadOnlyList<DraftBullet> Bullets { get; }

            internal IReadOnlyList<DraftCard> Cards { get; }

            internal DraftTable Table { get; }

            internal DraftChart Chart { get; }

            internal string Takeaway { get; set; } = "";
            internal string Caption { get; set; } = "";
            internal string Sources { get; set; } = "";
            internal string Evidence { get; set; } = "";
            internal DraftTable SecondaryTable { get; set; }
            internal DraftChart SecondaryChart { get; set; }
            internal IReadOnlyList<int> HighlightRows { get; set; } = new int[0];
        }

        // One card of the three-column strategy grid.
        internal sealed class DraftCard
        {
            internal DraftCard(
                string heading,
                IReadOnlyList<string> points)
            {
                Heading = heading ?? string.Empty;
                Points = points ?? new string[0];
            }

            internal string Heading { get; }

            internal IReadOnlyList<string> Points { get; }
        }

        // A dense performance matrix: header row plus data rows.
        internal sealed class DraftTable
        {
            internal DraftTable(
                IReadOnlyList<string> headers,
                IReadOnlyList<IReadOnlyList<string>> rows)
            {
                Headers = headers ?? new string[0];
                Rows = rows ??
                    new IReadOnlyList<string>[0];
            }

            internal IReadOnlyList<string> Headers { get; }

            internal IReadOnlyList<IReadOnlyList<string>> Rows
            {
                get;
            }
        }

        // A native chart drawn on the slide from bounded data the
        // model supplies: categories down the side, one to five
        // named series of numbers.
        internal sealed class DraftChart
        {
            internal DraftChart(
                int typeCode,
                string title,
                IReadOnlyList<string> categories,
                IReadOnlyList<DraftChartSeries> series)
            {
                TypeCode = typeCode;
                Title = title ?? string.Empty;
                Categories = categories ?? new string[0];
                Series = series ?? new DraftChartSeries[0];
            }

            internal int TypeCode { get; }

            internal string Title { get; }

            internal IReadOnlyList<string> Categories { get; }

            internal IReadOnlyList<DraftChartSeries> Series
            {
                get;
            }
        }

        internal sealed class DraftChartSeries
        {
            internal DraftChartSeries(
                string name,
                IReadOnlyList<double> values)
            {
                Name = name ?? string.Empty;
                Values = values ?? new double[0];
            }

            internal string Name { get; }

            internal IReadOnlyList<double> Values { get; }
        }

        // A bullet line with its outline level (1-5). Sub-bullets
        // are written with two leading spaces per level.
        internal sealed class DraftBullet
        {
            internal DraftBullet(string text, int level)
            {
                Text = text ?? string.Empty;
                Level = level < 1 ? 1 : (level > 5 ? 5 : level);
            }

            internal string Text { get; }

            internal int Level { get; }
        }

        internal static string AddDraftSlides(
            object powerPointApplication,
            IReadOnlyList<DraftSlide> slides)
        {
            return AddDraftSlides(
                powerPointApplication,
                slides,
                null);
        }

        internal static string AddDraftSlides(
            object powerPointApplication,
            IReadOnlyList<DraftSlide> slides,
            int? afterSlide)
        {
            return AddDraftSlides(
                powerPointApplication,
                slides,
                afterSlide,
                false);
        }

        // afterSlide places the new slides after that slide number
        // (0 = at the very start); null keeps the default append at
        // the end. Existing slides are still never modified - only
        // the insertion point moves. inNewPresentation forces a
        // brand-new unsaved deck - the cross-app send tools use it
        // so content handed over from another app never lands
        // quietly inside whatever deck happens to be open.
        internal static string AddDraftSlides(
            object powerPointApplication,
            IReadOnlyList<DraftSlide> slides,
            int? afterSlide,
            bool inNewPresentation,
            Action<SamsungOutput> onRendered = null)
        {
            if (slides == null || slides.Count == 0)
            {
                throw new InvalidOperationException(
                    "At least one draft slide is required.");
            }

            dynamic application = powerPointApplication;
            var planned = ComposeSamsung(slides);
            dynamic presentation = null;
            if (!inNewPresentation)
            {
                try
                {
                    presentation = application.ActivePresentation;
                }
                catch
                {
                }
            }

            if (presentation == null)
            {
                // msoTrue window so the new unsaved deck is visible.
                presentation = application.Presentations.Add(-1);
            }

            var existing = (int)presentation.Slides.Count;
            if (existing == 0)
            {
                presentation.PageSetup.SlideWidth = SamsungSlideDesign.Width;
                presentation.PageSetup.SlideHeight = SamsungSlideDesign.Height;
            }
            else if (Math.Abs((double)presentation.PageSetup.SlideWidth - SamsungSlideDesign.Width) > .1 ||
                     Math.Abs((double)presentation.PageSetup.SlideHeight - SamsungSlideDesign.Height) > .1)
                throw new InvalidOperationException("SAMSUNG_CANVAS_MISMATCH: Use a new 960 x 540 presentation. Existing slides were not resized.");
            var anchor = existing;
            if (afterSlide.HasValue)
            {
                anchor = afterSlide.Value;
                if (anchor < 0)
                {
                    anchor = 0;
                }

                if (anchor > existing)
                {
                    anchor = existing;
                }
            }

            var inserted = anchor < existing;
            var added = 0;
            var charts = 0;
            var tables = 0;
            var owner = Guid.NewGuid().ToString("N");
            foreach (var page in planned)
            {
                var index = anchor + added + 1;
                dynamic created = presentation.Slides.Add(
                    index,
                    PpLayoutBlank);
                var output = DrawSamsungPage((object)created, page, owner);
                output.Image = ExportSamsung(output);
                onRendered?.Invoke(output);
                var drawn = (page.Elements.Any(e => e.Chart != null) ? 1 : 0) | (page.Elements.Any(e => e.Table != null) ? 2 : 0);
                if ((drawn & 1) != 0)
                {
                    charts++;
                }

                if ((drawn & 2) != 0)
                {
                    tables++;
                }

                added++;
            }

            var position = inNewPresentation
                ? "in a new unsaved draft presentation"
                : (inserted
                    ? (anchor == 0
                        ? "at the start of the presentation"
                        : "after slide " + anchor)
                    : "at the end of the presentation");
            var contents = string.Empty;
            if (charts > 0 || tables > 0)
            {
                contents = " with " +
                    (charts > 0
                        ? charts +
                          (charts == 1
                              ? " native chart"
                              : " native charts")
                        : string.Empty) +
                    (charts > 0 && tables > 0
                        ? " and "
                        : string.Empty) +
                    (tables > 0
                        ? tables +
                          (tables == 1
                              ? " data table"
                              : " data tables")
                        : string.Empty);
            }

            return "Added " + added + " marked draft slides" +
                contents + " " + position + ", styled with the " +
                MetoTheme.ThemeName +
                " corporate theme. Nothing was saved.";
        }

        // Paints one slide from the theme. Returns bit 1 when a
        // chart was drawn and bit 2 when a table was drawn.
        private static int RenderSlide(
            dynamic slide,
            dynamic presentation,
            DraftSlide draft)
        {
            double width = presentation.PageSetup.SlideWidth;
            double height = presentation.PageSetup.SlideHeight;
            if (string.Equals(
                draft.Layout,
                MetoTheme.LayoutCover,
                StringComparison.Ordinal))
            {
                RenderCover(slide, draft, width, height);
                return 0;
            }

            PaintBackground(
                slide,
                MetoTheme.Rgb(MetoTheme.WhiteHex));
            AddDraftTag(
                slide,
                width,
                height,
                MetoTheme.Rgb(MetoTheme.LightGrayHex));
            var top = RenderHeader(slide, draft, width, height);
            var left = width * 0.055;
            var contentWidth = width * 0.89;
            var bottom = draft.Footnote.Length > 0
                ? height * 0.875
                : height * 0.925;
            var contentHeight = bottom - top;
            if (contentHeight < height * 0.2)
            {
                // A very long title and subtitle must still leave a
                // usable content band.
                contentHeight = height * 0.2;
            }

            var drawn = 0;

            if (string.Equals(
                draft.Layout,
                MetoTheme.LayoutAgenda,
                StringComparison.Ordinal))
            {
                RenderAgenda(
                    slide,
                    draft,
                    left,
                    top,
                    contentWidth,
                    contentHeight);
            }
            else if (string.Equals(
                draft.Layout,
                MetoTheme.LayoutCards,
                StringComparison.Ordinal))
            {
                RenderCards(
                    slide,
                    draft,
                    left,
                    top,
                    contentWidth,
                    contentHeight,
                    width);
            }
            else
            {
                drawn = RenderComposedContent(
                    slide,
                    draft,
                    left,
                    top,
                    contentWidth,
                    contentHeight,
                    width);
            }

            RenderFootnote(slide, draft, width, height);
            return drawn;
        }

        // Lays out whatever content blocks the slide carries. An
        // executive slide is dense: a table and a chart and the
        // takeaways share one slide instead of being spread thin
        // over three. Returns bit 1 for a chart and bit 2 for a
        // table.
        private static int RenderComposedContent(
            dynamic slide,
            DraftSlide draft,
            double left,
            double top,
            double width,
            double height,
            double slideWidth)
        {
            var hasTable = draft.Table != null;
            var hasChart = draft.Chart != null;
            var hasBullets = draft.Bullets.Count > 0;
            var gap = slideWidth * 0.02;
            var drawn = 0;

            if (hasTable && (hasChart || hasBullets))
            {
                // Table leads on the left; the chart and takeaways
                // stack down the right column.
                var tableWidth = hasChart
                    ? width * 0.56
                    : width * 0.62;
                var rightLeft = left + tableWidth + gap;
                var rightWidth = width - tableWidth - gap;
                if (RenderTable(
                    slide,
                    draft.Table,
                    left,
                    top,
                    tableWidth,
                    height))
                {
                    drawn |= 2;
                }

                if (hasChart && hasBullets)
                {
                    var chartHeight = height * 0.6;
                    if (AddChartToSlide(
                        slide,
                        draft.Chart,
                        rightLeft,
                        top,
                        rightWidth,
                        chartHeight))
                    {
                        drawn |= 1;
                    }

                    RenderBullets(
                        slide,
                        draft.Bullets,
                        rightLeft,
                        top + chartHeight + gap,
                        rightWidth,
                        height - chartHeight - gap);
                }
                else if (hasChart)
                {
                    if (AddChartToSlide(
                        slide,
                        draft.Chart,
                        rightLeft,
                        top,
                        rightWidth,
                        height))
                    {
                        drawn |= 1;
                    }
                }
                else
                {
                    RenderBullets(
                        slide,
                        draft.Bullets,
                        rightLeft,
                        top,
                        rightWidth,
                        height);
                }

                return drawn;
            }

            if (hasTable)
            {
                if (RenderTable(
                    slide,
                    draft.Table,
                    left,
                    top,
                    width,
                    height))
                {
                    drawn |= 2;
                }

                return drawn;
            }

            if (hasChart && hasBullets)
            {
                var textWidth = width * 0.44;
                RenderBullets(
                    slide,
                    draft.Bullets,
                    left,
                    top,
                    textWidth,
                    height);
                if (AddChartToSlide(
                    slide,
                    draft.Chart,
                    left + textWidth + gap,
                    top,
                    width - textWidth - gap,
                    height))
                {
                    drawn |= 1;
                }

                return drawn;
            }

            if (hasChart)
            {
                if (AddChartToSlide(
                    slide,
                    draft.Chart,
                    left,
                    top,
                    width,
                    height))
                {
                    drawn |= 1;
                }

                return drawn;
            }

            if (hasBullets)
            {
                RenderBullets(
                    slide,
                    draft.Bullets,
                    left,
                    top,
                    width,
                    height);
            }

            return drawn;
        }

        // The C-suite cover: deep royal blue field, oversized brand
        // title, and metadata in the lower right.
        private static void RenderCover(
            dynamic slide,
            DraftSlide draft,
            double width,
            double height)
        {
            PaintBackground(
                slide,
                MetoTheme.Rgb(MetoTheme.BrandBlueHex));
            AddDraftTag(
                slide,
                width,
                height,
                MetoTheme.Rgb(MetoTheme.CardSlateHex));
            var left = width * 0.08;
            var boxWidth = width * 0.84;
            var title = AddTextBox(
                slide,
                left,
                height * 0.33,
                boxWidth,
                height * 0.24);
            SetText(
                title,
                draft.Title,
                MetoTheme.TitleFont,
                MetoTheme.FitTitleSize(
                    draft.Title,
                    MetoTheme.CoverTitleSize,
                    30),
                true,
                MetoTheme.Rgb(MetoTheme.WhiteHex),
                PpAlignLeft);

            if (draft.Subtitle.Length > 0)
            {
                var subtitle = AddTextBox(
                    slide,
                    left,
                    height * 0.59,
                    boxWidth,
                    height * 0.08);
                SetText(
                    subtitle,
                    draft.Subtitle,
                    MetoTheme.SubtitleFont,
                    MetoTheme.CoverSubtitleSize,
                    true,
                    MetoTheme.Rgb(MetoTheme.CardSlateHex),
                    PpAlignLeft);
            }

            // A short accent rule separates title from metadata.
            AddRectangle(
                slide,
                left,
                height * 0.27,
                width * 0.12,
                4,
                MetoTheme.Rgb(MetoTheme.WhiteHex),
                false,
                0);

            if (draft.Footnote.Length > 0)
            {
                var meta = AddTextBox(
                    slide,
                    width * 0.5,
                    height * 0.82,
                    width * 0.42,
                    height * 0.08);
                SetText(
                    meta,
                    draft.Footnote,
                    MetoTheme.SubtitleFont,
                    MetoTheme.CoverSubtitleSize,
                    true,
                    MetoTheme.Rgb(MetoTheme.WhiteHex),
                    PpAlignRight);
            }
        }

        // Title, optional scope subtitle, the brand rule beneath
        // them, and the optional unit indicator. Returns the top of
        // the content band.
        private static double RenderHeader(
            dynamic slide,
            DraftSlide draft,
            double width,
            double height)
        {
            var left = width * 0.055;
            var contentWidth = width * 0.89;
            var isAgenda = string.Equals(
                draft.Layout,
                MetoTheme.LayoutAgenda,
                StringComparison.Ordinal);
            var titleSize = isAgenda
                ? MetoTheme.SlideTitleSize
                : MetoTheme.FitTitleSize(
                    draft.Title,
                    MetoTheme.SlideTitleSize,
                    46);
            var title = AddTextBox(
                slide,
                left,
                height * 0.055,
                contentWidth,
                height * 0.11);
            SetText(
                title,
                draft.Title,
                MetoTheme.TitleFont,
                titleSize,
                true,
                MetoTheme.Rgb(MetoTheme.CharcoalHex),
                isAgenda ? PpAlignCenter : PpAlignLeft);

            var ruleTop = height * 0.185;
            if (draft.Subtitle.Length > 0)
            {
                var subtitle = AddTextBox(
                    slide,
                    left,
                    height * 0.175,
                    contentWidth,
                    height * 0.055);
                SetText(
                    subtitle,
                    draft.Subtitle,
                    MetoTheme.SubtitleFont,
                    MetoTheme.SubtitleSize,
                    false,
                    MetoTheme.Rgb(MetoTheme.GrayHex),
                    isAgenda ? PpAlignCenter : PpAlignLeft);
                ruleTop = height * 0.245;
            }

            AddRectangle(
                slide,
                left,
                ruleTop,
                contentWidth,
                2.5,
                MetoTheme.Rgb(MetoTheme.BrandBlueHex),
                false,
                0);

            var contentTop = ruleTop + height * 0.045;
            if (draft.Unit.Length > 0)
            {
                var unit = AddTextBox(
                    slide,
                    left + contentWidth * 0.5,
                    ruleTop + height * 0.012,
                    contentWidth * 0.5,
                    height * 0.045);
                SetText(
                    unit,
                    draft.Unit,
                    MetoTheme.LabelFont,
                    MetoTheme.ChartLabelSize,
                    false,
                    MetoTheme.Rgb(MetoTheme.GrayHex),
                    PpAlignRight);
                contentTop = ruleTop + height * 0.06;
            }

            return contentTop;
        }

        private static void RenderFootnote(
            dynamic slide,
            DraftSlide draft,
            double width,
            double height)
        {
            if (draft.Footnote.Length == 0)
            {
                return;
            }

            var text = draft.Footnote.StartsWith(
                "*",
                StringComparison.Ordinal)
                ? draft.Footnote
                : "*" + draft.Footnote;
            var box = AddTextBox(
                slide,
                width * 0.055,
                height * 0.905,
                width * 0.89,
                height * 0.05);
            SetText(
                box,
                text,
                MetoTheme.FootnoteFont,
                MetoTheme.FootnoteSize,
                false,
                MetoTheme.Rgb(MetoTheme.GrayHex),
                PpAlignLeft);
        }

        // The agenda page: a clean centered list of section titles.
        private static void RenderAgenda(
            dynamic slide,
            DraftSlide draft,
            double left,
            double top,
            double width,
            double height)
        {
            if (draft.Bullets.Count == 0)
            {
                return;
            }

            var lines = new List<string>();
            foreach (var bullet in draft.Bullets)
            {
                if (bullet.Text.Length > 0)
                {
                    lines.Add(bullet.Text);
                }
            }

            if (lines.Count == 0)
            {
                return;
            }

            var box = AddTextBox(
                slide,
                left + width * 0.18,
                top + height * 0.08,
                width * 0.64,
                height * 0.8);
            SetText(
                box,
                string.Join("\r", lines),
                MetoTheme.TitleFont,
                MetoTheme.SectionHeaderSize,
                true,
                MetoTheme.Rgb(MetoTheme.CharcoalHex),
                PpAlignLeft);
            try
            {
                dynamic range = box.TextFrame.TextRange;
                dynamic paragraphs = range.ParagraphFormat;
                paragraphs.SpaceAfter = 14;
                // 2 = ppBulletNumbered: numbered agenda entries in
                // brand blue.
                dynamic bullet = paragraphs.Bullet;
                bullet.Visible = MsoTrue;
                bullet.Type = 2;
                bullet.Font.Color.RGB =
                    MetoTheme.Rgb(MetoTheme.BrandBlueHex);
                bullet.Font.Name = MetoTheme.TitleFont;
            }
            catch
            {
            }
        }

        // The strategy grid: up to four slate cards, each headed by
        // a circled number in brand blue.
        private static void RenderCards(
            dynamic slide,
            DraftSlide draft,
            double left,
            double top,
            double width,
            double height,
            double slideWidth)
        {
            var cards = new List<DraftCard>();
            foreach (var card in draft.Cards)
            {
                if (cards.Count == MaxCards)
                {
                    break;
                }

                cards.Add(card);
            }

            if (cards.Count == 0)
            {
                return;
            }

            var gap = slideWidth * 0.018;
            var cardWidth =
                (width - gap * (cards.Count - 1)) / cards.Count;
            var cardHeight = height * 0.92;
            for (var index = 0; index < cards.Count; index++)
            {
                var cardLeft = left + index * (cardWidth + gap);
                AddRectangle(
                    slide,
                    cardLeft,
                    top,
                    cardWidth,
                    cardHeight,
                    MetoTheme.Rgb(MetoTheme.CardSlateHex),
                    true,
                    MetoTheme.Rgb(MetoTheme.LightGrayHex));

                var padding = cardWidth * 0.08;
                var innerLeft = cardLeft + padding;
                var innerWidth = cardWidth - padding * 2;
                var number = AddTextBox(
                    slide,
                    innerLeft,
                    top + cardHeight * 0.04,
                    innerWidth,
                    cardHeight * 0.12);
                SetText(
                    number,
                    index < MetoTheme.CircledNumbers.Length
                        ? MetoTheme.CircledNumbers[index]
                        : string.Empty,
                    MetoTheme.TitleFont,
                    MetoTheme.SectionHeaderSize,
                    true,
                    MetoTheme.Rgb(MetoTheme.BrandBlueHex),
                    PpAlignLeft);

                var heading = AddTextBox(
                    slide,
                    innerLeft,
                    top + cardHeight * 0.17,
                    innerWidth,
                    cardHeight * 0.16);
                SetText(
                    heading,
                    cards[index].Heading,
                    MetoTheme.TitleFont,
                    MetoTheme.CardHeaderSize,
                    true,
                    MetoTheme.Rgb(MetoTheme.TextBlackHex),
                    PpAlignLeft);

                var points = new List<string>();
                foreach (var point in cards[index].Points)
                {
                    if (points.Count == MaxCardPoints)
                    {
                        break;
                    }

                    if (point.Length > 0)
                    {
                        points.Add(point);
                    }
                }

                if (points.Count == 0)
                {
                    continue;
                }

                var body = AddTextBox(
                    slide,
                    innerLeft,
                    top + cardHeight * 0.35,
                    innerWidth,
                    cardHeight * 0.6);
                SetText(
                    body,
                    string.Join("\r", points),
                    MetoTheme.BodyFont,
                    MetoTheme.BodyBulletSize,
                    false,
                    MetoTheme.Rgb(MetoTheme.TextBlackHex),
                    PpAlignLeft);
                ApplyBulletStyle(body);
            }
        }

        // The subsidiary performance matrix. Cell highlighting is
        // decided here from the cell text, never by the model, and
        // is capped per status so the table stays low-noise.
        private static bool RenderTable(
            dynamic slide,
            DraftTable table,
            double left,
            double top,
            double width,
            double height)
        {
            if (table == null)
            {
                return false;
            }

            var rows = new List<IReadOnlyList<string>>();
            if (table.Headers.Count > 0)
            {
                rows.Add(table.Headers);
            }

            foreach (var row in table.Rows)
            {
                if (rows.Count == MaxTableRows)
                {
                    break;
                }

                if (row != null && row.Count > 0)
                {
                    rows.Add(row);
                }
            }

            if (rows.Count == 0)
            {
                return false;
            }

            var columns = 1;
            foreach (var row in rows)
            {
                if (row.Count > columns)
                {
                    columns = Math.Min(row.Count, MaxTableColumns);
                }
            }

            var hasHeader = table.Headers.Count > 0;
            try
            {
                dynamic shape = slide.Shapes.AddTable(
                    rows.Count,
                    columns,
                    (float)left,
                    (float)top,
                    (float)width,
                    (float)(height * 0.92));
                dynamic grid = shape.Table;
                try
                {
                    // "No Style, No Grid" - the theme paints every
                    // fill and border itself.
                    grid.ApplyStyle(
                        "{2D5ABB26-0587-4C30-8999-92F81FD0307C}",
                        true);
                    grid.FirstRow = MsoFalse;
                    grid.HorizBanding = MsoFalse;
                }
                catch
                {
                }

                var goodUsed = 0;
                var badUsed = 0;
                for (var row = 0; row < rows.Count; row++)
                {
                    var isHeader = hasHeader && row == 0;
                    for (var column = 0;
                         column < columns;
                         column++)
                    {
                        var text = column < rows[row].Count
                            ? rows[row][column]
                            : string.Empty;
                        var status = MetoTheme.StatusNone;
                        if (!isHeader && column > 0)
                        {
                            status = MetoTheme.CellStatus(text);
                            if (status == MetoTheme.StatusGood)
                            {
                                if (goodUsed >=
                                    MetoTheme
                                        .MaxHighlightsPerStatus)
                                {
                                    status = MetoTheme.StatusNone;
                                }
                                else
                                {
                                    goodUsed++;
                                }
                            }
                            else if (status == MetoTheme.StatusBad)
                            {
                                if (badUsed >=
                                    MetoTheme
                                        .MaxHighlightsPerStatus)
                                {
                                    status = MetoTheme.StatusNone;
                                }
                                else
                                {
                                    badUsed++;
                                }
                            }
                        }

                        RenderTableCell(
                            grid,
                            row + 1,
                            column + 1,
                            text,
                            isHeader,
                            column == 0,
                            status);
                    }
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        private static void RenderTableCell(
            dynamic grid,
            int row,
            int column,
            string text,
            bool isHeader,
            bool isLabelColumn,
            int status)
        {
            try
            {
                dynamic cell = grid.Cell(row, column);
                dynamic shape = cell.Shape;
                var fill = MetoTheme.Rgb(MetoTheme.WhiteHex);
                if (isHeader)
                {
                    fill = MetoTheme.Rgb(MetoTheme.CardSlateHex);
                }
                else if (status == MetoTheme.StatusGood)
                {
                    fill = MetoTheme.Rgb(MetoTheme.GoodGreenHex);
                }
                else if (status == MetoTheme.StatusBad)
                {
                    fill = MetoTheme.Rgb(MetoTheme.BadYellowHex);
                }

                try
                {
                    dynamic shapeFill = shape.Fill;
                    shapeFill.Solid();
                    shapeFill.ForeColor.RGB = fill;
                    shapeFill.Transparency = 0f;
                }
                catch
                {
                }

                for (var border = 1; border <= 4; border++)
                {
                    try
                    {
                        dynamic line = cell.Borders(border);
                        line.ForeColor.RGB =
                            MetoTheme.Rgb(MetoTheme.LightGrayHex);
                        line.Weight = 0.75f;
                    }
                    catch
                    {
                    }
                }

                dynamic frame = shape.TextFrame;
                try
                {
                    frame.MarginLeft = 5f;
                    frame.MarginRight = 5f;
                    frame.MarginTop = 2f;
                    frame.MarginBottom = 2f;
                }
                catch
                {
                }

                dynamic range = frame.TextRange;
                range.Text = TextBoundary.SingleLine(
                    SafeModelText.Format(
                        text,
                        MaxCellCharacters).PlainText,
                    MaxCellCharacters);
                var bold = isHeader || isLabelColumn;
                StyleRange(
                    range,
                    bold
                        ? MetoTheme.TitleFont
                        : MetoTheme.BodyFont,
                    isHeader
                        ? MetoTheme.TableHeaderSize
                        : MetoTheme.TableBodySize,
                    bold,
                    isHeader
                        ? MetoTheme.Rgb(MetoTheme.CharcoalHex)
                        : MetoTheme.Rgb(MetoTheme.TextBlackHex),
                    column == 1 ? PpAlignLeft : PpAlignCenter);
            }
            catch
            {
            }
        }

        private static void RenderBullets(
            dynamic slide,
            IReadOnlyList<DraftBullet> bullets,
            double left,
            double top,
            double width,
            double height)
        {
            var lines = new List<DraftBullet>();
            foreach (var bullet in bullets)
            {
                if (lines.Count == MaxBulletsPerSlide)
                {
                    break;
                }

                if (bullet.Text.Length > 0)
                {
                    lines.Add(bullet);
                }
            }

            if (lines.Count == 0)
            {
                return;
            }

            var texts = new List<string>();
            foreach (var line in lines)
            {
                texts.Add(line.Text);
            }

            var box = AddTextBox(
                slide,
                left,
                top,
                width,
                height * 0.94);
            SetText(
                box,
                string.Join("\r", texts),
                MetoTheme.BodyFont,
                MetoTheme.BodyBulletSize,
                false,
                MetoTheme.Rgb(MetoTheme.TextBlackHex),
                PpAlignLeft);
            ApplyBulletStyle(box);
            try
            {
                dynamic range = box.TextFrame.TextRange;
                for (var line = 0; line < lines.Count; line++)
                {
                    if (lines[line].Level > 1)
                    {
                        try
                        {
                            // Length 1: exactly this paragraph,
                            // never the rest of the text frame.
                            range.Paragraphs(line + 1, 1)
                                .IndentLevel = lines[line].Level;
                        }
                        catch
                        {
                        }
                    }
                }
            }
            catch
            {
            }
        }

        // Draws one native chart in the given rectangle and fills
        // its data grid from the bounded model-supplied numbers,
        // then restyles it to the corporate chart standard: fine
        // gridlines, small charcoal labels, brand series colors, and
        // never a 3-D effect. The data grid workbook is the chart's
        // own embedded store inside the unsaved draft presentation -
        // closing it only closes the editing grid; no user file is
        // touched or saved.
        private static bool AddChartToSlide(
            dynamic slide,
            DraftChart chart,
            double left,
            double top,
            double width,
            double height)
        {
            try
            {
                dynamic shape = slide.Shapes.AddChart2(
                    -1,
                    chart.TypeCode,
                    (float)left,
                    (float)top,
                    (float)width,
                    (float)(height * 0.94),
                    true);
                dynamic slideChart = shape.Chart;
                slideChart.ChartData.Activate();
                dynamic dataWorkbook =
                    slideChart.ChartData.Workbook;
                dynamic dataSheet =
                    dataWorkbook.Worksheets[1];
                dataSheet.Cells[1, 1].Value2 = " ";
                for (var series = 0;
                     series < chart.Series.Count;
                     series++)
                {
                    dataSheet.Cells[1, series + 2].Value2 =
                        TextBoundary.SingleLine(
                            chart.Series[series].Name,
                            80);
                }

                for (var category = 0;
                     category < chart.Categories.Count;
                     category++)
                {
                    dataSheet.Cells[category + 2, 1].Value2 =
                        TextBoundary.SingleLine(
                            chart.Categories[category],
                            80);
                    for (var series = 0;
                         series < chart.Series.Count;
                         series++)
                    {
                        var values =
                            chart.Series[series].Values;
                        dataSheet.Cells[
                            category + 2,
                            series + 2].Value2 =
                            category < values.Count
                                ? values[category]
                                : 0d;
                    }
                }

                try
                {
                    dataSheet.ListObjects["Table1"].Resize(
                        dataSheet.Range(
                            dataSheet.Cells[1, 1],
                            dataSheet.Cells[
                                chart.Categories.Count + 1,
                                chart.Series.Count + 1]));
                }
                catch
                {
                }

                try
                {
                    if (chart.Title.Length > 0)
                    {
                        slideChart.HasTitle = true;
                        slideChart.ChartTitle.Text =
                            TextBoundary.SingleLine(
                                chart.Title,
                                180);
                        dynamic titleFont =
                            slideChart.ChartTitle.Format
                                .TextFrame2.TextRange.Font;
                        titleFont.Size = MetoTheme.ChartTitleSize;
                        titleFont.Name = MetoTheme.TitleFont;
                        titleFont.Bold = MsoTrue;
                        titleFont.Fill.ForeColor.RGB =
                            MetoTheme.Rgb(MetoTheme.CharcoalHex);
                    }
                    else
                    {
                        slideChart.HasTitle = false;
                    }
                }
                catch
                {
                }

                StyleChart(slideChart, chart);

                try
                {
                    // Alerts off so closing the embedded grid can
                    // never raise a modal prompt and hang the
                    // draft.
                    dataWorkbook.Application.DisplayAlerts = false;
                    dataWorkbook.Close(true);
                }
                catch
                {
                    // Leaving the data grid window open is only
                    // cosmetic; the chart itself already holds the
                    // data.
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        // Corporate chart formatting. Every step is cosmetic and
        // individually guarded: a chart type that lacks an axis or a
        // legend must never fail the draft.
        private static void StyleChart(
            dynamic slideChart,
            DraftChart chart)
        {
            try
            {
                slideChart.ChartArea.Format.Line.Visible = MsoFalse;
            }
            catch
            {
            }

            try
            {
                slideChart.PlotArea.Format.Line.Visible = MsoFalse;
            }
            catch
            {
            }

            // 1 = xlCategory, 2 = xlValue.
            for (var axis = 1; axis <= 2; axis++)
            {
                try
                {
                    dynamic target = slideChart.Axes(axis);
                    dynamic font = target.Format.TextFrame2
                        .TextRange.Font;
                    font.Size = MetoTheme.ChartLabelSize;
                    font.Name = MetoTheme.LabelFont;
                    font.Fill.ForeColor.RGB =
                        MetoTheme.Rgb(MetoTheme.CharcoalHex);
                    target.Format.Line.ForeColor.RGB =
                        MetoTheme.Rgb(MetoTheme.LightGrayHex);
                }
                catch
                {
                }
            }

            try
            {
                dynamic gridlines =
                    slideChart.Axes(2).MajorGridlines;
                gridlines.Format.Line.ForeColor.RGB =
                    MetoTheme.Rgb(MetoTheme.GridlineHex);
                gridlines.Format.Line.Weight = 0.75f;
            }
            catch
            {
            }

            try
            {
                if (chart.Series.Count > 1)
                {
                    slideChart.HasLegend = true;
                    // -4107 = xlLegendPositionBottom.
                    slideChart.Legend.Position = -4107;
                    dynamic legendFont = slideChart.Legend.Format
                        .TextFrame2.TextRange.Font;
                    legendFont.Size = MetoTheme.ChartLabelSize;
                    legendFont.Name = MetoTheme.LabelFont;
                    legendFont.Fill.ForeColor.RGB =
                        MetoTheme.Rgb(MetoTheme.CharcoalHex);
                }
                else
                {
                    slideChart.HasLegend = false;
                }
            }
            catch
            {
            }

            var palette = MetoTheme.ChartSeriesColors();
            for (var series = 1;
                 series <= chart.Series.Count;
                 series++)
            {
                var color = palette[(series - 1) % palette.Length];
                try
                {
                    dynamic target =
                        slideChart.SeriesCollection(series);
                    try
                    {
                        target.Format.Fill.Solid();
                        target.Format.Fill.ForeColor.RGB = color;
                    }
                    catch
                    {
                    }

                    try
                    {
                        target.Format.Line.ForeColor.RGB = color;
                        target.Format.Line.Weight = 2f;
                    }
                    catch
                    {
                    }
                }
                catch
                {
                }
            }
        }

        // --- Shape helpers ----------------------------------------

        private static dynamic AddTextBox(
            dynamic slide,
            double left,
            double top,
            double width,
            double height)
        {
            dynamic box = slide.Shapes.AddTextbox(
                MsoTextOrientationHorizontal,
                (float)left,
                (float)top,
                (float)width,
                (float)height);
            try
            {
                dynamic frame = box.TextFrame;
                frame.WordWrap = MsoTrue;
                frame.AutoSize = 0;
                frame.MarginLeft = 0f;
                frame.MarginRight = 0f;
                frame.MarginTop = 0f;
                frame.MarginBottom = 0f;
            }
            catch
            {
            }

            try
            {
                // 2 = msoAutoSizeTextToFitShape: dense corporate
                // content shrinks to fit rather than overflowing.
                box.TextFrame2.AutoSize = 2;
                box.TextFrame2.WordWrap = MsoTrue;
            }
            catch
            {
            }

            return box;
        }

        private static void SetText(
            dynamic box,
            string text,
            string font,
            float size,
            bool bold,
            int color,
            int alignment)
        {
            try
            {
                dynamic range = box.TextFrame.TextRange;
                range.Text = text ?? string.Empty;
                StyleRange(range, font, size, bold, color, alignment);
            }
            catch
            {
            }
        }

        private static void StyleRange(
            dynamic range,
            string font,
            float size,
            bool bold,
            int color,
            int alignment)
        {
            try
            {
                dynamic target = range.Font;
                target.Name = font;
                target.Size = size;
                target.Bold = bold ? MsoTrue : MsoFalse;
                target.Color.RGB = color;
            }
            catch
            {
            }

            try
            {
                range.ParagraphFormat.Alignment = alignment;
            }
            catch
            {
            }
        }

        private static void ApplyBulletStyle(dynamic box)
        {
            try
            {
                dynamic paragraphs =
                    box.TextFrame.TextRange.ParagraphFormat;
                paragraphs.SpaceAfter = 6;
                dynamic bullet = paragraphs.Bullet;
                bullet.Visible = MsoTrue;
                // 1 = ppBulletUnnumbered, 8226 = bullet character.
                bullet.Type = 1;
                bullet.Character = 8226;
                bullet.RelativeSize = 0.8f;
                bullet.Font.Color.RGB =
                    MetoTheme.Rgb(MetoTheme.BrandBlueHex);
            }
            catch
            {
            }
        }

        private static dynamic AddRectangle(
            dynamic slide,
            double left,
            double top,
            double width,
            double height,
            int fill,
            bool hasBorder,
            int borderColor)
        {
            dynamic shape = slide.Shapes.AddShape(
                MsoShapeRectangle,
                (float)left,
                (float)top,
                (float)width,
                (float)height);
            try
            {
                dynamic shapeFill = shape.Fill;
                shapeFill.Solid();
                shapeFill.ForeColor.RGB = fill;
                shapeFill.Transparency = 0f;
            }
            catch
            {
            }

            try
            {
                if (hasBorder)
                {
                    shape.Line.Visible = MsoTrue;
                    shape.Line.ForeColor.RGB = borderColor;
                    shape.Line.Weight = 0.75f;
                }
                else
                {
                    shape.Line.Visible = MsoFalse;
                }
            }
            catch
            {
            }

            try
            {
                shape.Shadow.Visible = MsoFalse;
            }
            catch
            {
            }

            return shape;
        }

        private static void PaintBackground(
            dynamic slide,
            int color)
        {
            try
            {
                slide.FollowMasterBackground = MsoFalse;
                dynamic fill = slide.Background.Fill;
                fill.Solid();
                fill.ForeColor.RGB = color;
            }
            catch
            {
            }
        }

        // The draft marker rides on every drafted slide, in the same
        // corner, at label size. It is written by the writer and can
        // never be suppressed by the model.
        private static void AddDraftTag(
            dynamic slide,
            double width,
            double height,
            int color)
        {
            var box = AddTextBox(
                slide,
                width * 0.7,
                height * 0.035,
                width * 0.245,
                height * 0.04);
            SetText(
                box,
                DraftMarker,
                MetoTheme.LabelFont,
                MetoTheme.ChartLabelSize,
                false,
                color,
                PpAlignRight);
        }

        // --- Parsing ----------------------------------------------

        // Converts the model-supplied JSON slides value into bounded
        // draft slides, rejecting anything but arrays of objects.
        internal static IReadOnlyList<DraftSlide> ParseSlides(
            object value)
        {
            var outer = value as IEnumerable;
            if (outer == null || value is string)
            {
                throw new InvalidOperationException(
                    "slides must be an array of objects.");
            }

            var slides = new List<DraftSlide>();
            foreach (var slideValue in outer)
            {
                if (slides.Count == MaxDraftSlides)
                {
                    throw new InvalidOperationException("At most ten input slides per batch; continue in another call. No slides were silently omitted.");
                }

                var map = slideValue as
                    IDictionary<string, object>;
                if (map == null)
                {
                    throw new InvalidOperationException(
                        "Each slide must be an object with a title.");
                }

                ValidateSamsungInput(map);

                object titleValue;
                map.TryGetValue("title", out titleValue);
                // SafeModelText strips **bold** markers so slides
                // never show literal asterisks - the theme owns all
                // emphasis.
                var title = Clean(
                    Convert.ToString(titleValue),
                    MaxTitleCharacters);
                if (title.Length == 0)
                {
                    throw new InvalidOperationException(
                        "Each slide needs a non-empty title.");
                }

                var bullets = new List<DraftBullet>();
                object bulletsValue;
                if (map.TryGetValue("bullets", out bulletsValue))
                {
                    var list = bulletsValue as IEnumerable;
                    if (list != null && !(bulletsValue is string))
                    {
                        foreach (var bullet in list)
                        {
                            if (bullets.Count ==
                                MaxBulletsPerSlide)
                            {
                                break;
                            }

                            var raw = Convert.ToString(bullet) ??
                                string.Empty;
                            var leading = 0;
                            while (leading < raw.Length &&
                                   raw[leading] == ' ')
                            {
                                leading++;
                            }

                            bullets.Add(new DraftBullet(
                                Clean(
                                    raw,
                                    MaxBulletCharacters),
                                1 + leading / 2));
                        }
                    }
                }

                object layoutValue;
                object subtitleValue;
                object footnoteValue;
                object unitValue;
                map.TryGetValue("layout", out layoutValue);
                map.TryGetValue("subtitle", out subtitleValue);
                map.TryGetValue("footnote", out footnoteValue);
                map.TryGetValue("unit", out unitValue);

                slides.Add(new DraftSlide(
                    title,
                    Clean(
                        Convert.ToString(subtitleValue),
                        MaxSubtitleCharacters),
                    Clean(
                        Convert.ToString(footnoteValue),
                        MaxFootnoteCharacters),
                    Clean(
                        Convert.ToString(unitValue),
                        MaxUnitCharacters),
                    Convert.ToString(layoutValue),
                    bullets,
                    ParseCards(map),
                    ParseTable(map),
                    ParseChart(map))
                {
                    Takeaway = SamsungString(map, "takeaway", 400),
                    Caption = SamsungString(map, "caption", 180),
                    Sources = SamsungString(map, "sources", 2000),
                    Evidence = SamsungString(map, "evidence", 12000),
                    SecondaryTable = ParseTable(new Dictionary<string, object> { { "table", SamsungValue(map, "secondary_table") } }),
                    SecondaryChart = ParseChart(new Dictionary<string, object> { { "chart", SamsungValue(map, "secondary_chart") } }),
                    HighlightRows = SamsungIndices(map)
                });
            }

            return slides;
        }

        // Reads the optional strategy-grid cards. Anything malformed
        // simply yields no cards rather than failing the draft.
        private static IReadOnlyList<DraftCard> ParseCards(
            IDictionary<string, object> slideMap)
        {
            object cardsValue;
            if (!slideMap.TryGetValue("cards", out cardsValue))
            {
                return new DraftCard[0];
            }

            var list = cardsValue as IEnumerable;
            if (list == null || cardsValue is string)
            {
                return new DraftCard[0];
            }

            var cards = new List<DraftCard>();
            foreach (var entry in list)
            {
                if (cards.Count == MaxCards)
                {
                    break;
                }

                var map = entry as IDictionary<string, object>;
                if (map == null)
                {
                    continue;
                }

                object headingValue;
                object pointsValue;
                map.TryGetValue("heading", out headingValue);
                map.TryGetValue("points", out pointsValue);
                var points = new List<string>();
                var pointList = pointsValue as IEnumerable;
                if (pointList != null && !(pointsValue is string))
                {
                    foreach (var point in pointList)
                    {
                        if (points.Count == MaxCardPoints)
                        {
                            break;
                        }

                        var text = Clean(
                            Convert.ToString(point),
                            MaxBulletCharacters);
                        if (text.Length > 0)
                        {
                            points.Add(text);
                        }
                    }
                }

                var heading = Clean(
                    Convert.ToString(headingValue),
                    MaxTitleCharacters);
                if (heading.Length > 0 || points.Count > 0)
                {
                    cards.Add(new DraftCard(heading, points));
                }
            }

            return cards;
        }

        // Reads the optional performance matrix.
        private static DraftTable ParseTable(
            IDictionary<string, object> slideMap)
        {
            object tableValue;
            if (!slideMap.TryGetValue("table", out tableValue))
            {
                return null;
            }

            var map = tableValue as IDictionary<string, object>;
            if (map == null)
            {
                return null;
            }

            object headersValue;
            object rowsValue;
            map.TryGetValue("headers", out headersValue);
            map.TryGetValue("rows", out rowsValue);
            var headers = ParseCells(headersValue);
            var rows = new List<IReadOnlyList<string>>();
            var rowList = rowsValue as IEnumerable;
            if (rowList != null && !(rowsValue is string))
            {
                foreach (var row in rowList)
                {
                    if (rows.Count == MaxTableRows)
                    {
                        break;
                    }

                    var cells = ParseCells(row);
                    if (cells.Count > 0)
                    {
                        rows.Add(cells);
                    }
                }
            }

            if (headers.Count == 0 && rows.Count == 0)
            {
                return null;
            }

            return new DraftTable(headers, rows);
        }

        private static IReadOnlyList<string> ParseCells(
            object value)
        {
            var list = value as IEnumerable;
            if (list == null || value is string)
            {
                return new string[0];
            }

            var cells = new List<string>();
            foreach (var cell in list)
            {
                if (cells.Count == MaxTableColumns)
                {
                    break;
                }

                cells.Add(Clean(
                    Convert.ToString(cell),
                    MaxCellCharacters));
            }

            return cells;
        }

        // Reads the optional chart object of one slide; anything
        // malformed simply yields no chart rather than failing the
        // whole draft.
        private static DraftChart ParseChart(
            IDictionary<string, object> slideMap)
        {
            object chartValue;
            if (!slideMap.TryGetValue("chart", out chartValue))
            {
                return null;
            }

            var map = chartValue as IDictionary<string, object>;
            if (map == null)
            {
                return null;
            }

            object typeValue;
            object titleValue;
            object categoriesValue;
            object seriesValue;
            map.TryGetValue("type", out typeValue);
            map.TryGetValue("title", out titleValue);
            map.TryGetValue("categories", out categoriesValue);
            map.TryGetValue("series", out seriesValue);

            var categories = new List<string>();
            var categoryList = categoriesValue as IEnumerable;
            if (categoryList != null &&
                !(categoriesValue is string))
            {
                foreach (var category in categoryList)
                {
                    if (categories.Count == MaxChartCategories)
                    {
                        break;
                    }

                    categories.Add(Clean(
                        Convert.ToString(category),
                        80));
                }
            }

            var series = new List<DraftChartSeries>();
            var seriesList = seriesValue as IEnumerable;
            if (seriesList != null && !(seriesValue is string))
            {
                foreach (var entry in seriesList)
                {
                    if (series.Count == MaxChartSeries)
                    {
                        break;
                    }

                    var entryMap = entry as
                        IDictionary<string, object>;
                    if (entryMap == null)
                    {
                        continue;
                    }

                    object nameValue;
                    object valuesValue;
                    entryMap.TryGetValue("name", out nameValue);
                    entryMap.TryGetValue(
                        "values",
                        out valuesValue);
                    var values = new List<double>();
                    var valueList = valuesValue as IEnumerable;
                    if (valueList != null &&
                        !(valuesValue is string))
                    {
                        foreach (var value in valueList)
                        {
                            if (values.Count ==
                                MaxChartCategories)
                            {
                                break;
                            }

                            double parsed;
                            values.Add(double.TryParse(
                                Convert.ToString(
                                    value,
                                    System.Globalization
                                        .CultureInfo
                                        .InvariantCulture),
                                System.Globalization
                                    .NumberStyles.Any,
                                System.Globalization
                                    .CultureInfo
                                    .InvariantCulture,
                                out parsed)
                                ? parsed
                                : 0d);
                        }
                    }

                    series.Add(new DraftChartSeries(
                        Clean(Convert.ToString(nameValue), 80),
                        values));
                }
            }

            if (categories.Count == 0 || series.Count == 0)
            {
                return null;
            }

            return new DraftChart(
                DraftChartTypes.Resolve(
                    Convert.ToString(typeValue)),
                Clean(Convert.ToString(titleValue), 180),
                categories,
                series);
        }

        // One bounded, single-line, marker-free string. Model text
        // never carries its own formatting into a themed slide.
        private static string Clean(string value, int limit)
        {
            return TextBoundary.SingleLine(
                SafeModelText.Format(value, limit).PlainText,
                limit);
        }
    }
}
