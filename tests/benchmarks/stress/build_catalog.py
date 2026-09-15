"""Build 200 authored Office cases from deterministic, evaluator-only corpora.

This program never starts Office, contacts a model, or records a test pass.
Run the Office and mail corpus generators first; use --manifest only after all
native input files have been authored and verified.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import re
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1] / "generated" / "stress-corpus"
SUITE = "scribble-stress-v1"
FAMILIES = {"EX": 60, "OL": 70, "PP": 50, "XA": 20}


def read(path):
    return json.loads(Path(path).read_text(encoding="utf-8-sig"))


def write(root, relative, value):
    target = root / relative
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(json.dumps(value, ensure_ascii=False, separators=(",", ":"), allow_nan=False) + "\n", encoding="utf-8")


def safe(root, relative):
    if not isinstance(relative, str) or "\\" in relative or ":" in relative:
        raise ValueError(f"Invalid relative path: {relative!r}")
    target = (root / relative).resolve()
    if not target.is_relative_to(root.resolve()) or target == root.resolve():
        raise ValueError(f"Path escapes corpus: {relative!r}")
    return target


def numeric(cell, expected, *, formula=True, source_sheet="Ledger"):
    return {"kind": "numeric_cell", "sheet": "Scribble Draft*", "cell": cell,
            "expected": expected, "tolerance": 0.00001 if abs(expected) < 2 else 0.01,
            "formula_required": formula, "source_sheet": source_sheet}


def native_chart(labels, values, units=None, *, host):
    extension = {"Excel": "xlsx", "PowerPoint": "pptx"}[host]
    check = {"kind": "native_chart", "category_labels": labels, "series_values": values,
             "zero_baseline": True, "native_editable": True, "host": host, "artifact_extension": extension}
    if units:
        check["units"] = units
    return check


def chart_units(wb):
    return "units" if wb["domain"] == "inventory" else "hours" if wb["domain"] in ("workforce", "projects") else "EUR"


def source_checks(paths, artifacts):
    return ([{"kind": "native_artifact", "extension": ext} for ext in artifacts] +
            [{"kind": "source_unchanged", "path": path} for path in paths if path.endswith((".xlsx", ".pptx"))])


def add(cases, oracles, case_id, host, prompt, inputs, artifacts, checks, family,
        *, prerequisite=None, formula_probes=None, allow_source_edit=False, browser_allowed_hosts=None):
    case = {"id": case_id, "version": 1, "host": host, "prompt": prompt,
            "inputs": list(dict.fromkeys(inputs)), "artifacts": artifacts,
            "oracle_ref": f"evaluator-only/cases/{case_id}.json", "timeout_seconds": 900,
            "setup": "Use only the isolated synthetic corpus. Preserve source files and keep email drafts unsent.",
            "expected": "Evaluate sealed native output against the separate deterministic oracle; visual review is required.",
            "clarification_answers": {"audience": "the company operations leadership team",
                "period": "January through June 2026; current June versus May",
                "currency": "EUR where the source uses money; preserve all other source units",
                "format": "native editable Office output with source citations"},
            "task_family": family}
    if prerequisite:
        case["prerequisite_prompt"] = prerequisite
    if allow_source_edit:
        case["allow_source_edit"] = True
    if browser_allowed_hosts:
        case["browser_allowed_hosts"] = browser_allowed_hosts
    cases.append(case)
    oracles[case_id] = {"schema": 1, "suite_id": SUITE, "id": case_id,
        "checks": source_checks([] if allow_source_edit else inputs, artifacts) + checks,
        "review_required": ["native visual review", "source citation coverage", "no unsupported factual claims"]}
    if formula_probes:
        oracles[case_id]["formula_probes"] = formula_probes
        oracles[case_id]["review_required"].append("dependency perturbation on a throwaway output copy")


def workbook_rules(wb):
    return (f"Use {wb['company']} ({wb['id']}) {wb['domain']} data in {wb['primary_sheet']}. "
            "Count each RowID once, and disclose missing inputs rather than replacing them with zero. "
            f"{wb['aggregation']['period']}. Preserve all original worksheets. Create a new worksheet whose name begins Scribble Draft. "
            "Use live Excel formulas linked to the source observations for calculated outputs, not pasted answer constants. "
            "Keep full precision in formula results and use cell formatting for displayed rounding. ")


def excel_cases(cases, oracles, workbooks):
    for index, wb in enumerate(workbooks):
        facts = wb["facts"]
        common = workbook_rules(wb)
        ratio = wb.get("ratio_label")
        labels = [month["period"] for month in facts["monthly"]]
        series = [month["primary"] for month in facts["monthly"]]
        base = index * 3

        prompt = common + (f"Create a six-month trend audit. In A3:E3 use Period, {wb['primary_label']}, "
            f"{wb['secondary_label']}, Budget, Units as headers. In rows4–9 show January through June2026 in chronological order "
            "using YYYY-MM labels. The numeric columns are B primary, C secondary, D budget, E units. "
            f"Add a native editable column chart using A4:B9, with {chart_units(wb)} in the chart title and a zero-based value axis. "
            "Below the table explain June versus May and any incompleteness; do not sum stock snapshots across months.")
        checks = []
        for row, month in enumerate(facts["monthly"], 4):
            checks.extend([numeric(f"B{row}", month["primary"]), numeric(f"C{row}", month["secondary"]),
                           numeric(f"D{row}", month["budget"]), numeric(f"E{row}", month["units"])])
        checks.append(native_chart(labels, series, chart_units(wb), host="Excel"))
        if not facts["current"]["complete"]:
            checks.append({"kind": "required_text", "values": ["incomplete"]})
        add(cases, oracles, f"EX{base+1:02}", "Excel", prompt, [wb["path"]], ["xlsx"], checks,
            "monthly_formula_and_chart", formula_probes=wb["formula_probes"])

        groups = facts["by_group"]
        group_names = ", ".join(g["group"] for g in groups)
        prompt = common + (f"Audit June by group in this order: {group_names}. Put Group in A3, {wb['primary_label']} in B3, "
            f"{wb['secondary_label']} in C3, Budget in D3, Units in E3" + (f", and {ratio} in F3. " if ratio else ". ") +
            f"Start the group rows at row4. Show source-linked group calculations and a native editable column chart of group primary values with a zero-based value axis and {chart_units(wb)} in its title. "
            "Compare the two highest groups in a short note and explain any duplicate RowID or missing-input issue. " +
            ("Compute the rate from the correct aggregate numerator and denominator. If the rate cannot be determined, write incomplete instead of a number."
             if ratio else "Explain why these stock balances cannot be added across reporting dates."))
        checks = []
        for row, group in enumerate(groups, 4):
            checks.extend([numeric(f"B{row}", group["primary"]), numeric(f"C{row}", group["secondary"]),
                           numeric(f"D{row}", group["budget"]), numeric(f"E{row}", group["units"])])
            if ratio:
                checks.append(numeric(f"F{row}", group["ratio"]) if group.get("ratio") is not None else
                              {"kind": "cell_text", "sheet": "Scribble Draft*", "cell": f"F{row}", "expected": "incomplete"})
        checks.append(native_chart([g["group"] for g in groups], [g["primary"] for g in groups], chart_units(wb), host="Excel"))
        add(cases, oracles, f"EX{base+2:02}", "Excel", prompt, [wb["path"]], ["xlsx"], checks,
            "group_weighted_metric_and_data_quality")

        probe = wb["formula_probes"][0]
        prompt = common + (f"Build a June what-if comparison, treating {probe['sheet']}!{probe['cell']} as {probe['new_value']} "
            "only inside this draft scenario. Do not change the source cell. Use columns A Metric, B Baseline, C Scenario, D Change, with headers in row3. "
            f"Use rows4–7 for {wb['primary_label']}, {wb['secondary_label']}, Budget, Units respectively. "
            "The scenario must recalculate dependent metrics from that one overridden input, with Change equal to Scenario minus Baseline. " +
            (f"Put {ratio} in row8, using decimal fractions formatted as percentages in B8 and C8 and the formula C8-B8 in D8; write incomplete in a rate or rate-change cell that cannot be computed. " if ratio else "") +
            "Retain all other assumptions and explain the business implication and missing-data limits below the table.")
        checks = []
        for row, field in enumerate(["primary", "secondary", "budget", "units"], 4):
            before, after = facts["current"][field], probe["after_current"][field]
            checks.extend([numeric(f"B{row}", before), numeric(f"C{row}", after), numeric(f"D{row}", after-before)])
        if ratio:
            for address, value in [("B8", facts["current"].get("ratio")), ("C8", probe["after_current"].get("ratio"))]:
                checks.append(numeric(address, value) if value is not None else
                              {"kind": "cell_text", "sheet": "Scribble Draft*", "cell": address, "expected": "incomplete"})
            before_ratio, after_ratio = facts["current"].get("ratio"), probe["after_current"].get("ratio")
            checks.append(numeric("D8", after_ratio-before_ratio) if before_ratio is not None and after_ratio is not None else
                          {"kind": "cell_text", "sheet": "Scribble Draft*", "cell": "D8", "expected": "incomplete"})
        if not facts["current"]["complete"]:
            checks.append({"kind": "required_text", "values": ["incomplete"]})
        add(cases, oracles, f"EX{base+3:02}", "Excel", prompt, [wb["path"]], ["xlsx"], checks,
            "isolated_what_if_recalculation")


def presentation_check(deck, count, wb):
    current = wb["facts"]["current"]
    return {"kind": "presentation", "slide_count": count, "theme_ref": "evaluator-only/SamsungMD2.theme.json",
        "reference_pptx": deck["reference_path"],
        "geometry": {"width_points": 960, "height_points": 540, "minimum_body_font_pt": 14,
                     "no_overflow": True, "no_unintended_overlap": True},
        "required_facts": [current["primary"], current["secondary"]],
        "theme": deck["theme"]}


def powerpoint_cases(cases, oracles, presentations, workbook_by_id):
    for index, deck in enumerate(presentations):
        wb = workbook_by_id[deck["linked_workbook"]]
        monthly = wb["facts"]["monthly"]
        series = [[m["primary"] for m in monthly], [m["secondary"] for m in monthly]]
        series_description = "exactly two series, primary then secondary"
        if wb["domain"] == "inventory":
            series = [series[0]]
            series_description = "only the primary closing-units series; show the secondary inventory value in a separate table to avoid mixing units on one axis"
        prompt = (f"Create a repaired, editable draft of every slide in {deck['id']} for {deck['company']}. "
            f"Retain exactly {deck['slide_count']} output slides and the Samsung MD visual language. "
            f"Use the attached workbook as the authority for June facts. Preserve source slides; recreate the monthly chart as native editable clustered columns with {series_description}, a zero-based value axis, all six YYYY-MM categories and {chart_units(wb)} in its title. "
            "retain useful tables and speaker-note citations, fix stale chart categories, undersized text, overflow and unintended overlaps. "
            "Use the correct period and units, and disclose incomplete figures. Keep all text readable on the existing 16:9 canvas. "
            f"The current primary measure is {wb['primary_label']}; the secondary measure is {wb['secondary_label']}.")
        checks = [presentation_check(deck, deck["slide_count"], wb),
                  native_chart([m["period"] for m in monthly], series, chart_units(wb), host="PowerPoint")]
        if not wb["facts"]["current"]["complete"]:
            checks.append({"kind": "required_text", "values": ["incomplete"]})
        add(cases, oracles, f"PP{index+1:02}", "PowerPoint", prompt, [deck["path"], wb["path"]], ["pptx"], checks,
            "repair_preserve_and_reconcile")
    for index, wb in enumerate(workbook_by_id.values()):
        deck = presentations[index]
        prompt = (f"Using {wb['id']} for {wb['company']}, create exactly six new editable executive-review draft slides in the Samsung MD style "
            f"of {deck['id']}: one headline, one January–June {wb['primary_label']} trend chart, one June-by-group comparison, "
            "one current-versus-prior metric table, one data-quality risk assessment, and one recommended next-step slide. "
            f"The trend chart must be native editable columns with only the primary series, a zero-based value axis, YYYY-MM categories and {chart_units(wb)} in its title. Use native tables, explicit units and source citations in notes. "
            "Use 16:9, readable text and no unintended overlap. Preserve all source slides and worksheets. "
            "Derive facts from unique RowIDs and distinguish missing data from zero; proposed actions must be labelled as recommendations.")
        checks = [presentation_check(deck, 6, wb), native_chart(
            [m["period"] for m in wb["facts"]["monthly"]], [m["primary"] for m in wb["facts"]["monthly"]], chart_units(wb), host="PowerPoint")]
        add(cases, oracles, f"PP{index+31:02}", "PowerPoint", prompt, [deck["path"], wb["path"]], ["pptx"], checks,
            "new_six_slide_business_review")


def outlook_cases(cases, oracles, mail_catalog):
    tasks = mail_catalog["search_tasks"]
    if len(tasks) != 70:
        raise ValueError("The independent mail generator must provide exactly70 search tasks.")
    for task in tasks[:67]:
        case_id = task.get("case_id", task.get("id"))
        prompt = task["prompt"] + (" Search only the isolated Scribble synthetic mailbox. Continue native search pages until the scope is complete. "
            "In your final answer list every matching MAILnnnn identifier and write a separate line exactly in the form Total matches: N, including zero matches. "
            "Do not list nonmatching mail identifiers or send any message.")
        checks = [{"kind": "mail_search", "expected_ids": task["expected_ids"],
                   "expected_count": task.get("expected_count", len(task["expected_ids"])),
                   "scope": task["search_scope"]}]
        if task.get("required_text"):
            checks.append({"kind": "required_text", "values": task["required_text"]})
        known = []
        for facts in task.get("required_facts", {}).values():
            if not facts:
                continue
            for key, value in facts.items():
                if value is None or key in ("status", "mitigation", "mitigation_status", "data_status", "currency"):
                    continue
                if isinstance(value, (str, int, float)):
                    known.append(str(value))
                elif isinstance(value, list):
                    known.extend(str(item) for item in value)
        known.extend(task.get("expected_tail_markers", []))
        if known:
            checks.append({"kind": "required_text", "values": sorted(set(known))})
            prompt += " Report all dates as YYYY-MM-DD and retain the exact source owner names and stated numerical decision values."
        add(cases, oracles, case_id, "Outlook", prompt, [], [], checks,
            task.get("task_family", "native_mail_search_and_pagination"))
        cases[-1]["mailbox_mode"] = "native_synthetic_store"
        oracles[case_id]["mail_semantic_review"] = {"required_facts": task.get("required_facts", {}),
            "checks": task.get("checks", []), "requires_all_bodies": task.get("requires_all_bodies", False)}


def cross_app_cases(cases, oracles, workbooks, presentations):
    for index in range(10):
        wb, deck = workbooks[index], presentations[index]
        initial = workbook_rules(wb) + (f"Create a June-to-May audit table. Put Metric, May, June in A3:C3, "
            f"{wb['primary_label']} in A4 and {wb['secondary_label']} in A5. Use linked formulas for B4:C5. Explain any incomplete observations below.")
        prompt = (f"Use the draft worksheet just created from {wb['id']} to produce four new native editable Samsung MD PowerPoint slides for {wb['company']}: "
            f"headline, period comparison with a native chart, June group analysis and data-quality limits. The chart must use only primary values for May and June with YYYY-MM categories, a zero-based value axis and {chart_units(wb)} in the title. "
            "Read the derived workbook output through the normal cross-app handoff; retain units and cite the workbook in notes. Preserve original inputs.")
        checks = [numeric("B4", wb["facts"]["previous"]["primary"]), numeric("C4", wb["facts"]["current"]["primary"]),
                  numeric("B5", wb["facts"]["previous"]["secondary"]), numeric("C5", wb["facts"]["current"]["secondary"]),
                  presentation_check(deck, 4, wb), native_chart(["2026-05", "2026-06"],
                    [wb["facts"]["previous"]["primary"], wb["facts"]["current"]["primary"]], chart_units(wb), host="PowerPoint")]
        add(cases, oracles, f"XA{index+1:02}", "Excel", prompt, [wb["path"], deck["path"]], ["xlsx", "pptx"], checks,
            "excel_output_to_powerpoint", prerequisite=initial)
    for index in range(10, 20):
        wb, deck = workbooks[index], presentations[index]
        initial = (f"Create exactly three native editable Samsung MD draft slides for {wb['company']} from the attached workbook: "
            f"June {wb['primary_label']} and {wb['secondary_label']}, a comparison against May, and data-quality limitations. "
            "Preserve source slides and cite the workbook in the notes. Disclose incomplete values.")
        subject = f"{wb['id']} June operational review"
        prompt = (f"Create a new unsent Outlook email to review@example.test with subject {subject}. "
            "Use the new presentation through the normal cross-app handoff, and summarize its June primary and secondary values, "
            f"using one separate line for each metric in the exact form {wb['primary_label']}: value and {wb['secondary_label']}: value. "
            "Place explanation after those metric lines, including the comparison with May and any incomplete-data limitation. Cite the source workbook identifier and leave the native draft open for review. "
            "Do not invent actions, attach unrelated files or send the message.")
        checks = [presentation_check(deck, 3, wb), {"kind": "draft_mail", "to": ["review@example.test"], "cc": [], "bcc": [],
            "subject_contains": subject, "body_contains": [wb["id"]], "body_metrics": [
                {"label": wb["primary_label"], "expected": wb["facts"]["current"]["primary"]},
                {"label": wb["secondary_label"], "expected": wb["facts"]["current"]["secondary"]}], "unsent": True}]
        add(cases, oracles, f"XA{index+1:02}", "PowerPoint", prompt, [deck["path"], wb["path"]], ["pptx", "msg"], checks,
            "powerpoint_output_to_outlook", prerequisite=initial)


def hero_cases(cases, oracles, hero, report, presentations):
    cases[:] = [case for case in cases if case["id"] not in ("EX60", "XA20")]
    oracles.pop("EX60", None)
    oracles.pop("XA20", None)

    korean = hero["korean"]
    add(cases, oracles, "EX60", "Excel",
        "Translate every Korean text cell in every worksheet of the active workbook into English. Preserve worksheet names, row and column positions, numbers, dates, formatting, and table structure. Replace only the detected Korean text cells, leave formulas and merged cells unchanged, and complete the full-workbook translation without saving over the source file.",
        [korean["path"]], ["xlsx"], [{"kind": "workbook_exact_text", "sheets": korean["sheets"]}],
        "whole_workbook_korean_to_english", allow_source_edit=True)

    theme = {"kind": "presentation", "slide_count": 8, "theme_ref": "evaluator-only/SamsungMD2.theme.json",
        "reference_pptx": presentations[-1]["reference_path"],
        "required_facts": [420000, 365000, 92.4, 2.1, "R-317", "Noura Ali", "2026-11-15"],
        "minimum_shapes_per_slide": 3, "minimum_native_charts": 1, "minimum_table_cells": 8}
    add(cases, oracles, "OL68", "Outlook",
        "Find the five synthetic emails containing the exact packet marker ORION-FOLD-2026. Read all five complete message bodies and every attachment. Create exactly eight new editable PowerPoint slides in the supplied dense Samsung executive style: executive headline, portfolio mandate, channel performance, customer evidence, risk, source-backed analysis, recommendation, and appendix. Include a native editable chart comparing the two stated UAE revenue channels, at least one native table, concise source citations in speaker notes using the MAIL identifiers, explicit units, and clear separation of observed facts from recommendations. Do not send email and do not alter any source attachment.",
        [], ["pptx"], [theme, {"kind": "native_chart", "category_labels": ["UAE direct", "UAE partner"],
            "series_values": [420000, 365000], "zero_baseline": True, "native_editable": True,
            "host": "PowerPoint", "artifact_extension": "pptx", "units": "AED"}],
        "five_email_attachment_to_samsung_deck")
    cases[-1]["mailbox_mode"] = "native_synthetic_store"

    add(cases, oracles, "OL69", "Outlook",
        "Read the attached 130-page synthetic PDF in full and write a concise executive summary covering the approved baseline, service-level result versus target, the critical risk and owner, and the final recommendation with its date and prerequisite. State important caveats, distinguish evidence from proposed action, and include the final page marker to demonstrate complete-page coverage. Do not invent missing facts.",
        [report["path"]], [], [{"kind": "required_text", "values": report["required_facts"] + [report["terminal_marker"]]}],
        "long_pdf_full_document_summary")

    allowed = ["www.samsungtradein.ae", "samsungtradein.ae"]
    add(cases, oracles, "OL70", "Chrome",
        "Open https://www.samsungtradein.ae/ae-en/ and complete the public UAE trade-in estimate flow for a new Galaxy Z Fold8 using an Apple iPhone 16 Pro, 256 GB, in Flawless condition. Do not sign in, submit personal information, or buy anything. On the final result page, record verified browser evidence and summarize the exact trade-in amount, currency, configured products, storage, condition, market, source URL, observation caveat, and observation time. If the public flow blocks access, report the verified blocker rather than guessing.",
        [], [], [{"kind": "browser_evidence", "purchasedProduct": "Galaxy Z Fold8", "tradeInProduct": "Apple iPhone 16 Pro",
            "storage": "256 GB", "condition": "Flawless", "market": "United Arab Emirates", "currency": "AED",
            "allowed_hosts": allowed}], "live_uae_trade_in_verified_evidence", browser_allowed_hosts=allowed)

    word = hero["word"]
    add(cases, oracles, "XA20", "Excel",
        "Transfer all eight worksheets from the active workbook into one new Word document. For each worksheet, create a clearly titled native Word table in the original workbook order, preserving every header, row, column, value, symbol, date, percentage, and identifier exactly. Do not omit or reformat source values. After the eight tables, write a concise cross-sheet analysis that identifies the revenue pattern, budget variances, staffing gaps, inventory position, project and risk dependencies, and dated actions. Keep the source workbook unchanged and leave the Word draft open for review.",
        [word["path"]], ["docx"], [{"kind": "word_tables", "tables": word["tables"]},
            {"kind": "required_text", "values": ["revenue", "budget", "staffing", "inventory", "risk", "action"]}],
        "eight_sheet_workbook_to_exact_word_tables")


def validate(cases, oracles):
    counts = Counter(case["id"][:2] for case in cases)
    if dict(counts) != FAMILIES or len(cases) != 200:
        raise ValueError(f"Expected200 cases in {FAMILIES}, found {dict(counts)}.")
    ids = [case["id"] for case in cases]
    if len(set(ids)) != 200 or set(ids) != set(oracles):
        raise ValueError("Case IDs and oracle IDs must form a one-to-one mapping.")
    if len({case["prompt"] for case in cases}) != 200:
        raise ValueError("Case prompts must be distinct.")
    for case in cases:
        if not re.fullmatch(r"[A-Z]{2}[0-9]{2}", case["id"]):
            raise ValueError("Invalid case ID.")
        if any(not path.startswith("inputs/") or "evaluator-only" in path for path in case["inputs"]):
            raise ValueError("An evaluator file leaked into model-visible inputs.")
        if any(key in case for key in ("checks", "expected_ids", "formula_probes")):
            raise ValueError("An oracle leaked into the operator case.")
        if not oracles[case["id"]]["checks"]:
            raise ValueError("Every case must have deterministic checks.")
        for check in oracles[case["id"]]["checks"]:
            if check["kind"] == "native_chart" and (check.get("host"), check.get("artifact_extension")) not in (
                ("Excel", "xlsx"), ("PowerPoint", "pptx")
            ):
                raise ValueError("Every native chart must target its intended host and artifact type.")


def manifest(root):
    cases = read(root / "operator/cases.json")
    for case in cases:
        for relative in [*case["inputs"], case["oracle_ref"]]:
            if not safe(root, relative).is_file():
                raise ValueError(f"Cannot seal an incomplete corpus: {relative}")
        for check in read(root / case["oracle_ref"])["checks"]:
            if check.get("theme_ref") and not safe(root, check["theme_ref"]).is_file():
                raise ValueError(f"Missing reference: {check['theme_ref']}")
    if len(list((root / "inputs/outlook").glob("MAIL*.eml"))) != 500:
        raise ValueError("The corpus must contain exactly500 native mail sources.")
    if len(list((root / "inputs/excel").glob("WB*.xlsx"))) != 20 or len(list((root / "inputs/powerpoint").glob("PPT*.pptx"))) != 30:
        raise ValueError("The corpus must contain exactly20 workbooks and30 presentations.")
    files = []
    for path in sorted(root.rglob("*")):
        relative = path.relative_to(root).as_posix()
        if not path.is_file() or relative in ("manifest.json", "manifest.sha256") or not relative.startswith(("inputs/", "operator/", "evaluator-only/")):
            continue
        if relative.startswith(("evaluator-only/native-validation/", "evaluator-only/previews/")) or relative.endswith(("_validation.json", ".inspect.ndjson")):
            continue
        safe(root, relative)
        payload = path.read_bytes()
        files.append({"path": relative, "sha256": hashlib.sha256(payload).hexdigest(), "size": len(payload),
                      "role": "input" if relative.startswith("inputs/") else "oracle" if relative.startswith("evaluator-only/") else "operator"})
    write(root, "manifest.json", {"schema": 1, "suite_id": SUITE, "files": files})
    digest = hashlib.sha256((root / "manifest.json").read_bytes()).hexdigest()
    (root / "manifest.sha256").write_text(digest + "\n", encoding="ascii")
    return digest


def build(root, seal=False):
    office = read(root / "evaluator-only/office_catalog.json")
    mail = read(root / "evaluator-only/mail_catalog.json")
    hero = read(root / "evaluator-only/hero_inputs.json")
    report = read(root / "evaluator-only/hero_pdf.json")
    workbooks, presentations = office["workbooks"], office["presentations"]
    if len(workbooks) != 20 or len(presentations) != 30:
        raise ValueError("Office generator must provide20 workbooks and30 presentations.")
    cases, oracles = [], {}
    excel_cases(cases, oracles, workbooks)
    outlook_cases(cases, oracles, mail)
    powerpoint_cases(cases, oracles, presentations, {wb["id"]: wb for wb in workbooks})
    cross_app_cases(cases, oracles, workbooks, presentations)
    hero_cases(cases, oracles, hero, report, presentations)
    validate(cases, oracles)
    write(root, "operator/cases.json", cases)
    for case in cases:
        write(root, case["oracle_ref"], oracles[case["id"]])
    write(root, "operator/stress-plan.json", {"schema": 1, "suite_id": SUITE, "case_count": 200,
        "groups": FAMILIES, "case_ids": [c["id"] for c in cases], "execution_status": "not_run",
        "resume_policy": "Select remaining case IDs explicitly in a new immutable attempt; never overwrite or rerun completed results automatically.",
        "completion_policy": "All requested cases need terminal native evidence and a validated PDF. Infrastructure completion is separate from correctness."})
    return {"case_count": len(cases), "oracle_bytes": sum((root / c["oracle_ref"]).stat().st_size for c in cases),
            "manifest_sha256": manifest(root) if seal else None, "model_requests": 0, "execution_status": "not_run"}


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=ROOT)
    parser.add_argument("--manifest", action="store_true", help="Seal only a fully authored corpus; does not run tests.")
    arguments = parser.parse_args()
    print(json.dumps(build(arguments.root.resolve(), arguments.manifest), indent=2))
