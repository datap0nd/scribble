# Samsung Slide Presentation Design Specification for Scribble

This document establishes the official presentation design standard for Samsung corporate slides, compiled by analyzing 20 representative slides from high-level management decks (including President Strategy reviews, PM Division updates, and MX/B2B Regional conferences) stored in the Samsung presentation directory `\\MX-SHARE\Users\METOMX\Desktop\METO PPT AI`.

This specification provides precise, quantitative layout, typography, and styling metrics (expressed in percentages of slide width and height) to enable engineers to programmatically generate slides that conform perfectly to Samsung’s elite executive design standards.

---

## 1. Slide-by-Slide Visual Analysis (20 Representative Slides)

The following analysis is derived directly from empirical measurements of **President Top priorities.pptx** (20 slides), **PM Strategy 21.pptx** (19 slides), and **B2B Report_V1.pptx** (22 slides). Slide dimensions are standard widescreen 16:9 layout (**960.0 x 540.0 points** / 13.33" x 7.50" inches).

### Slide 1: Title Slide (Cover Layout)

- **Purpose:** Set presentation topic, date, and division with corporate branding.
- **Information Density:** Low. Center-out focus.
- **Reading Order:** Main Title -> Subtitle -> Division / Date.
- **Layout & Positions:**
  - **Main Title:** Left: 4.6%, Top: 28.2%, Width: 84.4%, Height: 34.1% (Multi-line text box).
  - **Date:** Left: 4.0%, Top: 84.7%, Width: 32.0%, Height: 7.2%.
  - **Division Box:** Left: 87.2%, Top: 3.8%, Width: 12.1%, Height: 10.5%.
  - **Branded Accent Bars:** Top solid bar: Left: 0.0%, Top: -11.5%, Width: 1.6%, Height: 2.8% (filled with Steel Blue `#4F81BD`). Bottom solid bar: Left: 0.0%, Top: 94.7%, Width: 100.0%, Height: 2.1% (Steel Blue `#4F81BD`).
- **Typography:**
  - **Main Title:** Samsung Sharp Sans Bold, 66.0pt, Bold, Solid Black (`#000000`).
  - **Date:** Samsung Sharp Sans Bold, 22.1pt, Regular, Black.
  - **Division Box:** Samsung Sharp Sans Bold, 19.6pt, White (`#FFFFFF`).
- **Colors & Highlights:** White background `#FFFFFF` with high-contrast bold black typography and distinct corporate Steel Blue (`#4F81BD` / `#5B9BD5`) accent bars at the edges.

### Slide 2: Section Divider / Transition Layout

- **Purpose:** Group content into chapters and establish a logical narrative flow.
- **Information Density:** Low. Right-heavy reading focus.
- **Reading Order:** Divider Header -> Divider Subtitle.
- **Layout & Positions:**
  - **Divider Heading Box:** Left: 24.1%, Top: 36.8%, Width: 75.9%, Height: 63.2%.
- **Typography:**
  - **Divider Heading:** Samsung Sharp Sans Bold (or Arial), 40.0pt, Bold, Black.
- **Colors & Highlights:** Standard clean White `#FFFFFF` background or Solid Dark Navy background for major division separators, with high-contrast text.

### Slide 3: Content Outline / Macro Overview

- **Purpose:** Present high-level context or outline structure.
- **Information Density:** Medium.
- **Reading Order:** Slide Title -> Bullet Outline (Left Column) -> Summary takeaway.
- **Layout & Positions:**
  - **Slide Title:** Left: 6.1%, Top: 5.2%, Width: 90.0%, Height: 9.6%.
  - **Content Outline Box:** Left: 3.8%, Top: 25.1%, Width: 92.3%, Height: 24.3%.
- **Typography:**
  - **Title:** Samsung Sharp Sans Bold, 24.0pt, Bold, Black.
  - **Body Bullets:** Arial, 18.0pt, Regular, Black.
- **Colors & Highlights:** Standard corporate colors, clean margins.

### Slide 4: Multi-Visual (Image Grid) Slide

- **Purpose:** Compare multiple external visual charts, mockups, or reports.
- **Information Density:** High. Z-pattern reading order.
- **Reading Order:** Title -> Top-Left Graphic -> Top-Right Graphic -> Bottom-Left Graphic.
- **Layout & Positions:**
  - **Slide Title:** Left: 6.1%, Top: 5.2%, Width: 90.0%, Height: 9.6%.
  - **Graphic 1 (Top Left):** Left: 2.6%, Top: 16.7%, Width: 46.7%, Height: 35.1%.
  - **Graphic 2 (Top Right):** Left: 54.8%, Top: 16.7%, Width: 39.0%, Height: 36.2%.
  - **Graphic 3 (Bottom Left):** Left: 3.8%, Top: 55.7%, Width: 47.5%, Height: 38.1%.
- **Typography:** Title: Arial, 24.0pt, Bold, Black.
- **Colors & Highlights:** Thin, clean borders or no borders on images.

### Slide 5: Single Visual Chart Slide with Annotation Overlay

- **Purpose:** Highlight specific data metrics from an offline/online preference survey.
- **Information Density:** Very High. Dual-visual with complex annotations.
- **Reading Order:** Slide Title -> Action Title -> Left Chart -> Highlight Box -> Right Chart -> Bottom Takeaway.
- **Layout & Positions:**
  - **Slide Title:** Left: 6.1%, Top: 5.2%, Width: 90.0%, Height: 9.6%.
  - **Action Title:** Left: 3.1%, Top: 21.2%, Width: 82.9%, Height: 4.5%.
  - **Left Chart (Primary):** Left: 11.8%, Top: 41.1%, Width: 41.9%, Height: 38.9%.
  - **Right Chart (Secondary):** Left: 55.8%, Top: 45.9%, Width: 29.2%, Height: 33.1%.
  - **Highlight Annotation Box 1:** Left: 28.1%, Top: 50.8%, Width: 2.8%, Height: 18.2%.
  - **Highlight Annotation Box 2:** Left: 33.0%, Top: 59.9%, Width: 2.3%, Height: 15.6%.
  - **Takeaway Banner Box (Dark Blue):** Left: 22.2%, Top: 89.5%, Width: 60.6%, Height: 4.5%.
  - **Footnote Label:** Left: 12.2%, Top: 93.5%, Width: 11.8%, Height: 2.9%.
- **Typography:**
  - **Slide Title:** Samsung Sharp Sans Bold, 39.4pt, Bold.
  - **Action Title:** Arial, 14.0pt, Regular, Black.
  - **Chart Title:** Arial Narrow, 14.0pt, Bold.
  - **Annotation Text:** Arial Narrow, 8.0pt, Regular.
  - **Takeaway Banner:** Arial Narrow, 14.0pt, Bold, White (`#FFFFFF`).
  - **Footnote:** Arial Narrow, 7.0pt, Regular, Gray (`#7F7F7F`).
- **Colors & Highlights:**
  - Accent highlight rectangles use a bright Red (`#FF0000`) border of 1.0pt weight to draw immediate attention.
  - Takeaway Banner uses Solid Steel Blue (`#4F81BD`) background with white text.

### Slide 6: Section Divider (Chapter 2)

- **Purpose:** Re-establish flow for the second core theme.
- **Information Density:** Low.
- **Reading Order:** Centered Divider Text.
- **Layout, Typography & Colors:** Identical to Slide 2 to maintain consistency.

### Slide 7: Transition / Content Slide

- **Purpose:** Introduce a new focus area (e.g., Rebound Operator Business).
- **Information Density:** Medium. Top-down reading.
- **Reading Order:** Title -> Summary Block -> Detail Points.
- **Layout & Positions:**
  - **Slide Title:** Left: 6.1%, Top: 5.2%, Width: 90.0%, Height: 9.6%.
  - **Main Outline Box:** Left: 3.7%, Top: 16.0%, Width: 92.3%, Height: 69.1%.
- **Typography:** Title: Arial, 24.0pt. Body: Arial, 18.0pt.

### Slide 8: Engagement Model (Conceptual Layout)

- **Purpose:** Illustrate a conceptual business model or partnership framework.
- **Information Density:** Medium.
- **Reading Order:** Slide Title -> Engagement Diagram (Left Column).
- **Layout & Positions:**
  - **Slide Title:** Left: 6.1%, Top: 5.2%, Width: 90.0%, Height: 9.6%.
  - **Visual Diagram (Picture):** Left: 3.8%, Top: 22.1%, Width: 30.3%, Height: 53.5%.
- **Typography:** Title: Arial, 24.0pt, Bold.

### Slide 9: Dual Column Visual Layout (Lifecycle Management)

- **Purpose:** Present two parallel lifecycle diagrams or process maps side by side.
- **Information Density:** High. Left-to-right parallel comparison.
- **Reading Order:** Slide Title -> Left Visual -> Right Visual.
- **Layout & Positions:**
  - **Slide Title:** Left: 6.1%, Top: 5.2%, Width: 90.0%, Height: 9.6%.
  - **Left Image/Model:** Left: 2.5%, Top: 21.1%, Width: 42.9%, Height: 34.6%.
  - **Right Image/Model:** Left: 51.1%, Top: 20.8%, Width: 44.3%, Height: 34.9%.
- **Typography:** Title: Arial, 22.0pt, Bold.

### Slide 10: High-Density Specification Comparison (Detailed Spec Sheet)

- **Purpose:** Side-by-side comparison of device specifications vs. competitors (with strong/neutral/weak highlights).
- **Information Density:** Extremely High. Dense data layout.
- **Reading Order:** Title -> Action Title -> Spec Comparison Table -> Legend -> Key Conclusion Banner.
- **Layout & Positions:**
  - **Slide Title:** Left: 6.1%, Top: 5.2%, Width: 90.0%, Height: 9.6%.
  - **Action Title:** Left: 3.1%, Top: 21.2%, Width: 82.9%, Height: 4.5%.
  - **Table Caption:** Left: 15.6%, Top: 27.5%, Width: 24.8%, Height: 4.5%.
  - **Spec Comparison Table:** Left: 16.5%, Top: 32.1%, Width: 60.7%, Height: 52.1% (20 rows x 7 columns).
  - **Legend Indicators:** Centered around Left: 73.3% and 78.2%, Top sequence starting at 32.8% to 81.9%, Width: 1.0%, Height: 1.8%.
  - **Legend Text Labels:** Left: 78.6%, Width: 3.7%, Height: 3.1%.
  - **Bottom Conclusion Banner:** Left: 16.5%, Top: 85.1%, Width: 60.6%, Height: 7.6%.
- **Typography:**
  - **Slide Title:** Samsung Sharp Sans Bold, 39.4pt, Bold.
  - **Action Title:** Arial, 14.0pt, Regular, Black.
  - **Table Caption:** Arial Narrow, 14.0pt, Bold.
  - **Table Header/Body Font:** Arial Narrow, 9.0pt, Bold headers, 8.0pt body.
  - **Bottom Banner Text:** Arial Narrow, 14.0pt, Bold, White (`#FFFFFF`).
- **Colors & Highlights:**
  - Table uses thin vertical and horizontal lines.
  - Legend indicators use Solid Green (`#00B050`) for Strong, Equal Sign (`=`) for Neutral, and Solid Crimson Red (`#C00000`) for Weak.
  - Bottom Conclusion Banner is filled with Solid Steel Blue (`#4F81BD`) to frame the white text.

### Slide 11: Single Visual Diagram (Large Visual Frame)

- **Purpose:** Emphasize a single, highly detailed channel landscape chart.
- **Information Density:** High. Left-to-right flow.
- **Reading Order:** Slide Title -> Central Landscape Diagram.
- **Layout & Positions:**
  - **Slide Title:** Left: 6.1%, Top: 5.2%, Width: 90.0%, Height: 9.6%.
  - **Large Visual Frame (Picture):** Left: 3.8%, Top: 26.8%, Width: 73.9%, Height: 70.8%.
- **Typography:** Title: Arial, 24.0pt, Bold.

### Slide 12: Visual Chart with Header/Footer Outlines

- **Purpose:** Standard landscape analysis slide.
- **Information Density:** High. Vertical hierarchy.
- **Reading Order:** Title -> Graphic -> Takeaway text in footnotes.
- **Layout & Positions:**
  - **Slide Title:** Left: 6.1%, Top: 5.2%, Width: 90.0%, Height: 9.6%.
  - **Visual Frame:** Left: 6.9%, Top: 24.3%, Width: 72.9%, Height: 58.1%.
  - **Footnote Banner:** Left: 3.8%, Top: 92.8%, Width: 92.3%, Height: 5.3%.
- **Typography:** Title: Arial, 20.0pt, Bold.

### Slide 13: Centered Graphic (Visual Focus Layout)

- **Purpose:** Center-focused diagram to focus audience attention on one core trend.
- **Information Density:** Medium.
- **Reading Order:** Title -> Centered Graphic.
- **Layout & Positions:**
  - **Slide Title:** Left: 6.1%, Top: 5.2%, Width: 90.0%, Height: 9.6%.
  - **Centered Graphic Box:** Left: 13.5%, Top: 31.9%, Width: 61.1%, Height: 46.8%.
- **Typography:** Title: Arial, 24.0pt, Bold.

### Slide 14: Left Column Visual with Right Comments

- **Purpose:** Connect a graphic directly with verbal explanations.
- **Information Density:** Medium. Left-to-right reading order.
- **Reading Order:** Title -> Left Visual -> Right commentary.
- **Layout & Positions:**
  - **Slide Title:** Left: 6.1%, Top: 5.2%, Width: 90.0%, Height: 9.6%.
  - **Visual Box:** Left: 3.8%, Top: 21.4%, Width: 53.3%, Height: 39.7%.
- **Typography:** Title: Arial, 20.0pt, Bold.

### Slide 15: Wide Landscape Presentation Layout (Retail Presence Map)

- **Purpose:** Maximize slide space for maps, spreadsheets, or horizontal process flows.
- **Information Density:** High. Horizontal flow.
- **Reading Order:** Title -> Wide Landscape Graphic.
- **Layout & Positions:**
  - **Slide Title:** Left: 6.1%, Top: 5.2%, Width: 90.0%, Height: 9.6%.
  - **Wide Visual Graphic:** Left: 3.8%, Top: 26.5%, Width: 96.2%, Height: 73.8% (Almost full width).
- **Typography:** Title: Arial, 24.0pt, Bold.

### Slide 16: Competitor Positioning Matrix (Wide Visual Slide)

- **Purpose:** Display horizontal comparisons and tables mapping SKUs vs. competitors.
- **Information Density:** High. Z-pattern reading order.
- **Reading Order:** Slide Title -> Subtitle -> Comparative Matrix in Center.
- **Layout & Positions:**
  - **Slide Title:** Left: 6.1%, Top: 5.2%, Width: 90.0%, Height: 9.6% (Title contains two lines: "Product positioning / Compare the Top 3 vol/val SKUs...").
  - **Matrix Area:** Left: 5.0%, Top: 18.8%, Width: 85.8%, Height: 69.8%.
- **Typography:** Title: Arial, 18.0pt, Bold.

### Slide 17: Multi-Phase Roadmap Slide (Roadmap Diagram)

- **Purpose:** Outline immediate fixes vs. long-term strategic plans.
- **Information Density:** High. Left-to-right timeline.
- **Reading Order:** Title -> Milestone 1 -> Milestone 2 -> Milestone 3.
- **Layout & Positions:**
  - **Slide Title:** Left: 6.1%, Top: 5.2%, Width: 90.0%, Height: 9.6%.
  - **Roadmap Diagram:** Left: 5.2%, Top: 25.0%, Width: 78.1%, Height: 62.9%.
- **Typography:** Title: Arial, 22.0pt, Bold.

### Slide 18: Structural Diagram (Margin Stack Optimization)

- **Purpose:** Illustrate vertical layers, pricing structures, or hierarchical components.
- **Information Density:** Medium. Bottom-up or top-down reading order.
- **Reading Order:** Title -> Bottom-up stack layers -> Annotations.
- **Layout & Positions:**
  - **Slide Title:** Left: 6.1%, Top: 5.2%, Width: 90.0%, Height: 9.6%.
  - **Stack Diagram:** Left: 17.6%, Top: 32.1%, Width: 59.4%, Height: 54.9%.
- **Typography:** Title: Arial, 22.0pt, Bold.

### Slide 19: High-Density Bullet Summary (Takeaway Outline)

- **Purpose:** List strategic takeaways or next steps in a clean list format.
- **Information Density:** Medium.
- **Reading Order:** Title -> Bullet List (Top-to-Bottom).
- **Layout & Positions:**
  - **Slide Title:** Left: 6.1%, Top: 5.2%, Width: 90.0%, Height: 9.6%.
  - **Bullet Text Frame:** Left: 3.7%, Top: 16.0%, Width: 92.3%, Height: 69.1%.
- **Typography:** Title: Arial, 18.0pt, Bold. Body Bullets: Arial, 18.0pt.

### Slide 20: Closing Slide (Outro Cover Layout)

- **Purpose:** Conclude the presentation professionally.
- **Information Density:** Low. Centered focus.
- **Reading Order:** Central "Thank you" statement.
- **Layout & Positions:**
  - **"Thank you" Box:** Left: 7.8%, Top: 18.6%, Width: 63.8%, Height: 68.7%.
- **Typography:**
  - **"Thank you":** Arial, 28.0pt, Bold, White (`#FFFFFF`).
- **Colors & Highlights:** Dark corporate Steel Blue or solid Black background with white text, representing a strong premium exit layout.

---

## 2. Consolidated Shared Design Rules

To ensure visual consistency and preserve the authentic "Samsung corporate identity," the following shared design rules must be strictly applied to every generated slide deck:

1.  **Strict 16:9 Widescreen Canvas [Slide 1, Slide 10]:** All presentations are designed with wide slide dimensions of exactly **960.0 x 540.0 points** (13.33" x 7.50" inches). This widescreen setting is absolute and must never be altered.
2.  **Typography Architecture & Specific Font Roles [Slide 1, Slide 2, Slide 5, Slide 10]:**
    - **Branded Titles & Transition Slides:** **Samsung Sharp Sans Bold** must be used for titles and transition slides. This geometric corporate font provides the signature Samsung brand appearance.
    - **Action Titles & Body Copy:** **Arial** is the default standard for descriptive headers and body copy.
    - **High-Density Tables & Dense Callouts:** **Arial Narrow** is mandatory for all high-density spec comparison sheets, complex annotations, and tight table structures to ensure maximum data density while maintaining crisp legibility.
    - **Bilingual (English/Korean) or Standard Reports:** **Malgun Gothic** (맑은 고딕) is used for standard weekly reviews, providing clean, highly readable characters for multi-language or standard reporting decks.
3.  **Color Palette Integrity (COLORREF and Standard Hex mappings) [Slide 1, Slide 5, Slide 10]:**
    - **Primary Corporate Blue (Steel Blue):** Hex `#4F81BD` (COLORREF `12419407`). Used for title covers, transition backgrounds, takeaway banner highlights, and primary headers.
    - **Accent Soft Blue:** Hex `#5B9BD5` (COLORREF `13998939`). Used for sub-sections, accent cards, and secondary highlights.
    - **Dark Corporate Blue (Border Outlines):** Hex `#41719C` (COLORREF `10252609`). Used for outlining active tables, highlight cards, and main containers.
    - **Background Container Gray (Cards):** Hex `#F2F2F2` (COLORREF `15921906`). Used for card backgrounds, content panes, or table header fills.
    - **Status Indicators (Strong/Positive):** Solid Green `#00B050` (COLORREF `5287936`).
    - **Status Indicators (Weak/Negative):** Solid Crimson `#C00000` (COLORREF `192`).
    - **Text & Grid Gray:** Medium Gray `#7F7F7F` (COLORREF `8355711`) for footnotes and subtitles, and `#A6A6A6` (COLORREF `10921638`) for divider lines.
4.  **Action Title Consistency [Slide 2, Slide 5, Slide 10]:** Slides must include a single-line summary (the "Action Title") directly below the main title at **Top: 15.0% to 16.5%**, utilizing **14.0pt Arial (Regular)** in solid black. This ensures that the slide’s core strategic message is readable in under 2 seconds.
5.  **Page Number Footer [Slide 5, Slide 10]:** Slide numbers must be placed consistently at the bottom right corner: **Left: 96.1%, Top: 96.8%**, formatted as `"- [Page Number] -"` (e.g. `"- 5 -"`) in **10.5pt Calibri (Regular)**.

---

## 3. Reusable Layout Recipes

The following layout recipes are mathematically defined from the slide coordinates. Position coordinates (Left, Top) and size metrics (Width, Height) are expressed as percentages of the slide width and height.

### Recipe A: Single Column Action List Grid (Structured 3-Column Layout)

- **Slide Context:** Ideal for showing projects, actions, or next steps in a clean structured list with description and frequency columns.
- **Slide Layout Percentages:**
  - **Main Title:** Left: 6.1%, Top: 5.2%, Width: 90.0%, Height: 9.6%
  - **Action Title:** Left: 3.1%, Top: 16.1%, Width: 82.9%, Height: 4.5%
  - **Column 1 (Item Name / Phase):** Left: 4.7%, Width: 16.9%
  - **Column 2 (Detailed Content / Description):** Left: 23.1%, Width: 54.8%
  - **Column 3 (Timing / Frequency):** Left: 79.3%, Width: 12.3%
  - **Vertical Row Spacing:** Top margins for rows start at **26.4%** and increment by **7.0%** for each subsequent row. Height of each item is exactly **6.4%**.
  - **Page Number:** Left: 96.1%, Top: 96.8%, Width: 3.9%, Height: 3.2%

### Recipe B: Two-Pane Asymmetrical Layout (Left Container Card & Right Stacked Tables)

- **Slide Context:** Used to present a descriptive operational plan (left) alongside numerical targets or segment breakdowns (right).
- **Slide Layout Percentages:**
  - **Main Title:** Left: 6.1%, Top: 5.2%, Width: 90.0%, Height: 9.6%
  - **Takeaway Header:** Left: 0.2%, Top: 8.5%, Width: 97.1%, Height: 13.5%
  - **Left Container Pane (Solid Gray Card):** Left: 1.8%, Top: 25.1%, Width: 62.8%, Height: 73.3%
    - **Left Card Title Banner:** Left: 1.8%, Top: 25.3%, Width: 62.8%, Height: 6.0% (Solid Steel Blue, White text)
    - **Inside Card Text Box 1:** Left: 1.8%, Top: 31.3%, Width: 63.1%, Height: 18.8%
    - **Inside Card Text Box 2:** Left: 1.8%, Top: 70.6%, Width: 63.1%, Height: 14.4%
    - **Inside Card Text Box 3:** Left: 1.8%, Top: 84.5%, Width: 63.1%, Height: 14.4%
  - **Right Top Table (Sales Target / Breakdown):** Left: 67.0%, Top: 28.7%, Width: 29.7%, Height: 30.5%
    - **Table Caption:** Left: 73.3%, Top: 21.9%, Width: 18.4%, Height: 6.7%
    - **Unit Indicator:** Left: 93.9%, Top: 25.5%, Width: 3.5%, Height: 3.1% (Calibri 8pt, Right-aligned)
  - **Right Bottom Table (Channel Portion / Split):** Left: 67.0%, Top: 66.8%, Width: 29.7%, Height: 30.5%
    - **Table Caption:** Left: 73.5%, Top: 59.8%, Width: 18.4%, Height: 6.7%
    - **Unit Indicator:** Left: 94.6%, Top: 63.6%, Width: 2.6%, Height: 3.1% (Calibri 8pt, Right-aligned)
  - **Page Number:** Left: 96.1%, Top: 96.8%, Width: 3.9%, Height: 3.2%

### Recipe C: High-Density Spec Sheet & Comparison Matrix

- **Slide Context:** Designed to show product or strategy matrices comparing key specifications or status across 5+ categories.
- **Slide Layout Percentages:**
  - **Main Title:** Left: 6.1%, Top: 5.2%, Width: 90.0%, Height: 9.6%
  - **Action Title:** Left: 3.1%, Top: 21.2%, Width: 82.9%, Height: 4.5%
  - **Table Caption:** Left: 15.6%, Top: 27.5%, Width: 24.8%, Height: 4.5%
  - **Dense Matrix Table:** Left: 16.5%, Top: 32.1%, Width: 60.7%, Height: 52.1%
  - **Legend Indicators:** Left: 73.3%, Width: 1.0%, Height: 1.8%
  - **Legend Label Columns:** Left: 78.6%, Width: 3.7%, Height: 3.1%
  - **Bottom Action Takeaway Banner:** Left: 16.5%, Top: 85.1%, Width: 60.6%, Height: 7.6% (Filled with Steel Blue `#4F81BD`, White text)
  - **Page Number:** Left: 96.1%, Top: 96.8%, Width: 3.9%, Height: 3.2%

### Recipe D: Visual Chart & Annotation Overlay

- **Slide Context:** Designed for data charts overlaid with accent annotations to point directly to key statistical trends.
- **Slide Layout Percentages:**
  - **Main Title:** Left: 6.1%, Top: 5.2%, Width: 90.0%, Height: 9.6%
  - **Action Title:** Left: 3.1%, Top: 21.2%, Width: 82.9%, Height: 4.5%
  - **Primary Left Chart Box:** Left: 11.8%, Top: 41.1%, Width: 41.9%, Height: 38.9%
  - **Secondary Right Chart Box:** Left: 55.8%, Top: 45.9%, Width: 29.2%, Height: 33.1%
  - **Accent Red Annotation Frame 1:** Left: 28.1%, Top: 50.8%, Width: 2.8%, Height: 18.2%
  - **Accent Red Annotation Frame 2:** Left: 33.0%, Top: 59.9%, Width: 2.3%, Height: 15.6%
  - **Accent Gray Annotation Box:** Left: 42.0%, Top: 56.1%, Width: 8.8%, Height: 13.5%
  - **Elbow Pointer Line:** Left: 50.8%, Top: 41.5%, Width: 5.3%, Height: 21.4% (Weight: 0.5pt, gray `#7F7F7F`)
  - **Bottom Accent Conclusion Banner:** Left: 22.2%, Top: 89.5%, Width: 60.6%, Height: 4.5% (Filled with Steel Blue `#4F81BD`, White text)
  - **Survey Sample Footnote:** Left: 12.2%, Top: 93.5%, Width: 11.8%, Height: 2.9%
  - **Page Number:** Left: 96.1%, Top: 96.8%, Width: 3.9%, Height: 3.2%

---

## 4. Table, Chart, and Annotation Styling

To programmatically recreate the highly polished tables, charts, and annotations from the slides, implement the following detailed styling parameters:

### Table Styling

- **Header Styling:**
  - **Background Fill:** Solid Very Light Gray `#F2F2F2` (COLORREF `15921906`) or Steel Blue `#4F81BD` (COLORREF `12419407`).
  - **Font:** **Arial Narrow (Bold)**, **9.0pt to 11.0pt**. Left-aligned or centered depending on data type. Text color: Black or White (if blue header).
- **Data Row Styling:**
  - **Font:** **Arial Narrow (Regular)**, **8.0pt to 10.0pt**.
  - **Borders:** Vertical and horizontal borders must be thin, solid gray lines: weight **0.5pt to 0.75pt**, color Light Gray `#A6A6A6` (COLORREF `10921638`) or Steel Blue Border `#41719C` (COLORREF `10252609`).
  - **Alternate Row Shading:** Optional, very light background `#F9F9F9`.
- **Units / Annotations:**
  - Place a small right-aligned text box above the top right corner of the table: **Height: 3.1%, Width: ~3.0%**, containing `"(Unit)"` or `"(%)"` in **Calibri, 8.0pt (Regular)**.

### Chart Styling

- **Layout:** Standard bar, column, or line charts positioned strictly within designated visual container areas.
- **Primary Color Accents:** Bar colors should alternate between Samsung Corporate Steel Blue (`#4F81BD`), Accent Blue (`#5B9BD5`), and Light Gray (`#F2F2F2`).
- **Y/X Axis Labels:** **Arial Narrow, 8.0pt to 9.0pt**, in Dark Gray (`#7F7F7F`) text.
- **Legends:** Place legends either directly below the chart or top right, using **Arial Narrow, 8.0pt**, with no border.

### Visual Annotation Overlays (High-Contrast Callouts)

Samsung slides utilize active annotations overlaid directly onto charts and tables to focus the executive's attention on the slide's ultimate conclusion:

1.  **Accent Highlight Boxes [Slide 5]:** To highlight specific columns in a chart, overlay a thin, hollow rectangle shape:
    - **Line Color:** Solid Crimson Red (`#C00000`) or standard Red (`#FF0000`) to highlight a target/competitor, or Dark Steel Blue (`#41719C`) for internal highlights.
    - **Line Weight:** Exactly **1.0pt**.
    - **Fill:** No fill (Hollow).
2.  **Connector Arrows & Elbows [Slide 5]:** To link highlight frames to comments, use elbow connectors or straight lines:
    - **Connector Type:** Elbow connector.
    - **Line Color:** Medium Gray `#7F7F7F` (COLORREF `8355711`).
    - **Line Weight:** Exactly **0.5pt**.
3.  **Takeaway Accent Banners [Slide 5, Slide 10]:** To clearly frame the primary conclusion:
    - **Position:** Bottom center of the content region (typically `Top = 85.0% to 90.0%`).
    - **Background Fill:** Solid Steel Blue `#4F81BD` (COLORREF `12419407`).
    - **Line Border:** No outline border.
    - **Text Styling:** **Arial Narrow (Bold), 14.0pt**, text color Solid White (`#FFFFFF`).
4.  **Legend Status Indicators [Slide 10]:** For high-density grid comparison tables:
    - **Strong Performance:** Solid Green circular shape (`#00B050`, COLORREF `5287936`).
    - **Neutral/Average:** Solid equal sign (`=`) in Black.
    - **Weak/Negative:** Solid Red circular shape (`#C00000` or `#FF0000`).
    - **Indicator Dimensions:** Tiny centered circles (**Width: 1.0%, Height: 1.8%** of slide dimensions) to ensure they fit precisely inside table grid cells.

---

## 5. Writing Conventions

Samsung executive presentations adhere to rigorous verbal writing standards. Avoid "fluff" or conversational language.

1.  **The Summary-First Hierarchy [Slide 5, Slide 10]:** Every slide must have its conclusion stated immediately in the **Action Title** (Top: 15% to 16.5%). The executive must be able to read this single line and understand the entire slide's business outcome without scanning the charts or tables.
2.  **Actionable and Quantitative [Slide 5, Slide 10]:** Bullet points and takeaway banners must be highly quantitative. Use exact percentages and units instead of broad terms:
    - _Incorrect:_ "Screen size preferences are large."
    - _Correct:_ "In MENA, Acceptable screen size is 6.1"(91%) and 6.5"(86%)."
3.  **Concise Bullet Point Grammar [Slide 10, Slide 13]:** Bullet points should start with an action verb or a quantitative subject, keeping sentences under 15 words.
4.  **Casing Standards:**
    - **Slide Titles:** Sentence Case or Title Case (e.g. `1) Product Survey`).
    - **Action Titles:** Sentence Case (e.g. `More competitive low-end model is required`).
    - **Table Headers:** Title Case or ALL CAPS for compact headings (e.g. `QTY`, `ASP`, `M/S`).
5.  **Evidence Connection [Slide 5, Slide 10]:** Every conclusion stated in the **Action Title** must map directly to a red highlight frame or a green/red status indicator in the tables below. If a slide says "More competitive low-end model is required," the table below must highlight the low-end row in red.

---

## 6. Rules for Content Overflow

When the data or text provided exceeds the designated slide canvas space, Scribble must apply the following mathematical downscaling rules instead of letting text overlap or overflow:

1.  **Column Grid Compressing (Horizontal Overflow):**
    - If a 3-column recipe overflows horizontally, reduce the vertical gap between column margins to **1.0%** (minimum limit) before shrinking text.
    - If it still overflows, downscale the horizontal Width of the columns by a factor of **0.9** and align the content to the left margin.
2.  **Table Row Scaling (Vertical Overflow):**
    - For tables with 15+ rows, automatically default the font to **Arial Narrow**.
    - If a table overflows the vertical bounds (Height > 55%), scale down the font size by **0.5pt** increments (down to a minimum threshold of **7.5pt**).
    - If it still overflows at 7.5pt, split the table into a **Two-Column Grid** side-by-side or create a second slide.
3.  **Bullet Text Autoshrinking:**
    - If body text in standard text frames overflows the vertical container, reduce the font size from **18.0pt** down to **16.0pt**, then to **14.0pt** (absolute minimum for body text).
    - Reduce line spacing (line leading) from standard **1.2** to **1.0** line height.
4.  **Takeaway Banner Scaling:**
    - Key takeaway banners at the bottom must remain single-line or maximum two-line. If text overflows, shrink font size from **14.0pt** to **12.0pt** or **11.0pt** before wrapping text to a third line.

---

## 7. Slide Quality Checklist

Use this checklist to programmatically validate or manually inspect generated slides:

- [ ] **Canvas Check:** Is the slide size set to exactly 16:9 aspect ratio (960.0 x 540.0 points)?
- [ ] **Font Check:** Are titles in `Samsung Sharp Sans Bold` or `Arial`? Is body text in `Arial` or `Arial Narrow`?
- [ ] **Action Title Check:** Is there an Action Title positioned at `Top: 15% to 16.5%`? Does it contain a clear, quantitative business takeaway?
- [ ] **Color Palette Check:** Are primary blue accents matching `#4F81BD`? Are gray container card backgrounds `#F2F2F2`? Are borders `#A6A6A6` or `#41719C`?
- [ ] **Grid Alignment Check:** Do column grids align perfectly with Recipe A/B coordinates? Is there a consistent column gap of at least 1.4%?
- [ ] **Annotation Overlay Check:** Are red callout frames hollow (no fill) and exactly 1.0pt weight? Are elbow connectors exactly 0.5pt?
- [ ] **Table Caption/Units Check:** Do tables have unit labels like `(Unit)` or `(%)` right-aligned above the top-right border?
- [ ] **Page Number Check:** Is the page number consistently placed at `Left: 96.1%, Top: 96.8%` and formatted as `"- [Page] -"`?
- [ ] **Visual Margin Check:** Is there at least a 3.8% Left and Right margin to prevent content from touching slide boundaries?

---

## 8. Uncertainties (Requires PowerPoint Verification)

The following items are identified from the analysis as requiring validation inside PowerPoint by an engineer to ensure perfect system-level compatibility:

1.  **Samsung Sharp Sans Bold Availability:** Since `Samsung Sharp Sans Bold` is a proprietary corporate branding font, verify if PowerPoint fallback mappings (e.g. falling back to `Arial Bold` or `Segoe UI Bold`) degrade slide aesthetics if the font is not installed on the generation machine.
2.  **Elbow Connector Pinning Coordinates:** When generating elbow connectors programmatically, verify if the attachment point coordinates (`ConnectionSite`) snap accurately to chart bar shapes without clipping or rendering out-of-bounds in headless COM environments.
3.  **Bilingual Line Height Spacing:** When using `Malgun Gothic` for slides containing both English and Korean characters, verify if the line leading (spacing between vertical lines) expands excessively, causing vertical text overflow. If so, apply a line spacing override of exactly **1.0**.
