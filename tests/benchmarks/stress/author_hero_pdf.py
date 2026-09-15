"""Create the deterministic 130-page synthetic report used by OL69."""
from __future__ import annotations

import json
from pathlib import Path
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak


def build(root: Path) -> Path:
    output = root / "inputs/pdf/ExecutiveRiskReport130.pdf"
    output.parent.mkdir(parents=True, exist_ok=True)
    facts = {
        1: "The approved operating baseline is AED 18.4 million.",
        43: "The verified service-level result is 96.7 percent against a 97.5 percent target.",
        87: "Risk R-317 is the only critical risk and its owner is Noura Ali.",
        130: "The final recommendation is to defer Wave 3 until 2026-11-15 and complete the failover drill first.",
    }
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="Section", parent=styles["Heading2"], textColor=colors.HexColor("#1428A0"), fontName="Helvetica-Bold", fontSize=16, leading=20, spaceAfter=8))
    styles.add(ParagraphStyle(name="BodyDense", parent=styles["BodyText"], fontName="Helvetica", fontSize=9.5, leading=13, textColor=colors.HexColor("#1F2937")))
    story = []
    for page in range(1, 131):
        story.append(Paragraph(f"ORION Continuity Review — Section {page:03d}", styles["Section"]))
        story.append(Paragraph("Synthetic management report | Reporting date: 2026-09-15 | Confidential test fixture", styles["BodyDense"]))
        story.append(Spacer(1, 7 * mm))
        if page in facts:
            story.append(Paragraph(f"Key evidence: <b>{facts[page]}</b>", styles["BodyDense"]))
            story.append(Spacer(1, 4 * mm))
        body = (f"Section {page:03d} reviews a fictional operating control, its evidence owner, and the distinction between observed performance and planned remediation. "
                "All values in this document are synthetic. Missing evidence is treated as unknown rather than zero. Proposed actions are not reported as completed outcomes.")
        story.append(Paragraph(body, styles["BodyDense"]))
        story.append(Spacer(1, 6 * mm))
        rows = [["Control", "Owner", "Status", "Evidence date"],
                [f"CTL-{page:03d}-A", "Amina Rahman", "Effective", "2026-09-10"],
                [f"CTL-{page:03d}-B", "Omar Saleh", "Review", "2026-09-11"],
                [f"CTL-{page:03d}-C", "Noura Ali", "Open", "2026-09-12"]]
        table = Table(rows, colWidths=[35*mm, 40*mm, 30*mm, 35*mm])
        table.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#1428A0")),("TEXTCOLOR",(0,0),(-1,0),colors.white),
                                   ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),("FONTNAME",(0,1),(-1,-1),"Helvetica"),
                                   ("FONTSIZE",(0,0),(-1,-1),8.5),("GRID",(0,0),(-1,-1),0.5,colors.HexColor("#CBD5E1")),
                                   ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white,colors.HexColor("#F1F5F9")]),("VALIGN",(0,0),(-1,-1),"MIDDLE")]))
        story.append(table)
        story.append(Spacer(1, 10 * mm))
        story.append(Paragraph(f"Page marker: ORION-PAGE-{page:03d}. Cross-reference only to the evidence stated on this page.", styles["BodyDense"]))
        if page < 130:
            story.append(PageBreak())
    doc = SimpleDocTemplate(str(output), pagesize=A4, rightMargin=20*mm, leftMargin=20*mm, topMargin=18*mm, bottomMargin=18*mm,
                            title="ORION Continuity Review", author="Scribble synthetic benchmark")
    doc.build(story)
    metadata = {"schema": 1, "path": "inputs/pdf/ExecutiveRiskReport130.pdf", "page_count": 130,
                "required_facts": list(facts.values()),
                # Check the facts, not one exact prose rendering of them.  A
                # correct executive summary may use symbols (96.7%) or split a
                # sentence into bullets while retaining every material value.
                "required_fragments": ["AED 18.4 million", "96.7", "97.5", "R-317", "Noura Ali",
                                       "defer Wave 3", "2026-11-15", "failover drill"],
                "terminal_marker": "ORION-PAGE-130"}
    evaluator = root / "evaluator-only/hero_pdf.json"
    evaluator.parent.mkdir(parents=True, exist_ok=True)
    evaluator.write_text(json.dumps(metadata, indent=2) + "\n", encoding="utf-8")
    return output


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    print(build(args.output.resolve()))
