#!/usr/bin/env python3
"""Build deterministic, fictional Office sources for the 200-case stress suite.

The Python layer computes source data and independent answer keys. The bundled
Artifact Tool authors the actual editable XLSX/PPTX files. Nothing under
evaluator-only is an input to the model. No Office automation or model call is
performed by this generator.

Examples:
  python generate_office.py --data-only
  python generate_office.py --runtime-node PATH --node-modules PATH --presentations-skill PATH

--only WB01,PPT01 limits authoring, not the catalog. A complete build must omit
that switch. --verify checks every catalogued package after a complete build.
"""
from __future__ import annotations

import argparse
import collections
import datetime as dt
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import xml.etree.ElementTree as ET
import zipfile


VERSION = "scribble-office-stress-v1"
DEFAULT_ROOT = Path(__file__).resolve().parents[1] / "generated" / "stress-corpus"
COMPANIES = ["Atlas Components", "Meridian Retail", "Cedar Logistics", "Orion Services"]
DOMAINS = ["sales", "inventory", "budget", "returns", "workforce", "projects"]
GROUPS = ["North", "South", "East", "West"]
OWNERS = ["Mira Cole", "Leon Park", "Nadia Shah", "Evan Reed"]
MONTHS = [f"2026-{m:02}" for m in range(1, 7)]
THEME = {
    "name": "Scribble Samsung MD 2.0", "source": "src/Scribble/Office/SamsungSlideDesign.cs",
    "slide_size_px": [1280, 720], "slide_size_emu": [12192000, 6858000],
    "native_points": [960, 540],
    "font": "Arial", "font_basis": "Samsung MD fallback font; not a claim of installed Samsung Sharp Sans",
    "colors": {"blue": "4F81BD", "soft_blue": "5B9BD5", "border": "41719C", "gray": "F2F2F2", "red": "C00000", "green": "00B050"},
    "title_region_percent": [6.1, 5.2, 90, 9.6], "content_region_percent": [3.8, 25, 92.4, 57],
    "title_font_pt": 24, "body_font_pt": 18, "minimum_body_font_pt": 14,
    "human_fidelity_review_required": True,
}


def write_json(path: Path, data) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, indent=2, ensure_ascii=False, allow_nan=False) + "\n", encoding="utf-8")


def date(value: str):
    return {"$date": value}


def digest(value) -> str:
    return hashlib.sha256(json.dumps(value, sort_keys=True, separators=(",", ":"), ensure_ascii=False).encode()).hexdigest()


def column(i: int) -> str:
    value = ""
    while i:
        i, remainder = divmod(i - 1, 26)
        value = chr(65 + remainder) + value
    return value


def metrics(domain: str, row: dict) -> dict:
    """Independent arithmetic over source fields, never over exported caches."""
    if domain == "sales":
        revenue = None if row["UnitPriceEUR"] is None else row["Units"] * row["UnitPriceEUR"]
        return dict(primary=revenue, secondary=row["Units"] * row["UnitCostEUR"], budget=row["BudgetEUR"], units=row["Units"])
    if domain == "inventory":
        closing = None if row["ReceivedUnits"] is None else row["OpeningUnits"] + row["ReceivedUnits"] - row["ShippedUnits"]
        return dict(primary=closing, secondary=None if closing is None else closing * row["UnitCostEUR"], budget=row["ReorderUnits"], units=row["ShippedUnits"])
    if domain == "budget":
        return dict(primary=row["ActualEUR"], secondary=row["BudgetEUR"], budget=row["BudgetEUR"], units=1)
    if domain == "returns":
        return dict(primary=row["RefundEUR"], secondary=row["NetSalesEUR"], budget=row["UnitsSold"], units=row["UnitsReturned"])
    if domain == "workforce":
        return dict(primary=row["HoursWorked"], secondary=row["HoursAvailable"], budget=row["CostEUR"], units=row["Headcount"])
    return dict(primary=row["ActualHours"], secondary=row["PlannedHours"], budget=None if row["ActualHours"] is None else row["ActualHours"] * row["HourlyRateEUR"], units=1)


def aggregate(domain: str, rows: list[dict]) -> dict:
    values = [metrics(domain, r) for r in rows]
    result = {key: sum(v[key] for v in values if v[key] is not None) for key in ("primary", "secondary", "budget", "units")}
    result["missing_primary_count"] = sum(v["primary"] is None for v in values)
    result["row_count"] = len(rows)
    p, s = result["primary"], result["secondary"]
    if domain == "sales":
        result.update(profit=p - s, ratio=(p - s) / p if p else None, variance=p - result["budget"])
    elif domain == "inventory":
        result.update(ratio=None, variance=p - result["budget"])
    elif domain == "budget":
        result.update(ratio=(p - s) / s if s else None, variance=p - s)
    elif domain == "returns":
        result.update(ratio=result["units"] / result["budget"] if result["budget"] else None, variance=None)
    elif domain == "workforce":
        result.update(ratio=p / s if s else None, variance=p - s)
    else:
        result.update(ratio=(p - s) / s if s else None, variance=p - s)
    result["complete"] = not result["missing_primary_count"]
    # Margins/utilization derived from incomplete facts are not authoritative.
    if not result["complete"]:
        result["known_subtotal"] = result["primary"]
        if domain != "returns":
            result["ratio"] = None
        if domain == "sales":
            result["profit"] = None
    return result


def formula_values(domain: str, row: dict) -> dict:
    m=metrics(domain,row)
    if domain=="sales":
        return {"RevenueEUR":m["primary"],"CostEUR":m["secondary"]}
    if domain=="inventory":
        return {"ClosingUnits":m["primary"],"InventoryValueEUR":m["secondary"]}
    if domain=="budget":
        return {"VarianceEUR":None if m["primary"] is None else m["primary"]-m["secondary"]}
    if domain=="returns":
        return {"ReturnRate":row["UnitsReturned"]/row["UnitsSold"] if row["UnitsSold"] else None}
    if domain=="workforce":
        return {"Utilization":row["HoursWorked"]/row["HoursAvailable"] if row["HoursWorked"] is not None and row["HoursAvailable"] else None}
    return {"ActualCostEUR":m["budget"]}


def workbook_spec(number: int) -> tuple[dict, dict]:
    ident = f"WB{number:02}"
    domain = DOMAINS[(number - 1) % len(DOMAINS)]
    company = COMPANIES[(number - 1) % len(COMPANIES)]
    common = ["RowID", "Period", "Date", "Group", "Item"]
    definitions = {
        "sales": (["Units", "UnitPriceEUR", "UnitCostEUR", "RevenueEUR", "CostEUR", "BudgetEUR", "Status"], "Revenue EUR", "Cost EUR", "Gross margin", "RevenueEUR", "CostEUR", "BudgetEUR", "Units"),
        "inventory": (["OpeningUnits", "ReceivedUnits", "ShippedUnits", "ClosingUnits", "UnitCostEUR", "InventoryValueEUR", "ReorderUnits", "Status"], "Closing units", "Inventory value EUR", None, "ClosingUnits", "InventoryValueEUR", "ReorderUnits", "ShippedUnits"),
        "budget": (["ActualEUR", "BudgetEUR", "VarianceEUR", "Owner", "Status"], "Actual EUR", "Budget EUR", "Budget variance %", "ActualEUR", "BudgetEUR", "BudgetEUR", None),
        "returns": (["UnitsSold", "UnitsReturned", "NetSalesEUR", "RefundEUR", "ReturnRate", "Reason"], "Refund EUR", "Net sales EUR", "Return rate", "RefundEUR", "NetSalesEUR", "UnitsSold", "UnitsReturned"),
        "workforce": (["Headcount", "HoursAvailable", "HoursWorked", "CostEUR", "Utilization", "Status"], "Hours worked", "Hours available", "Utilization", "HoursWorked", "HoursAvailable", "CostEUR", "Headcount"),
        "projects": (["PlannedHours", "ActualHours", "HourlyRateEUR", "ActualCostEUR", "Owner", "DueDate", "Status"], "Actual hours", "Planned hours", "Hours overrun %", "ActualHours", "PlannedHours", "ActualCostEUR", None),
    }
    fields, primary_label, secondary_label, ratio_label, primary_field, secondary_field, budget_field, units_field = definitions[domain]
    headers = common + fields
    rows = []
    formulas = []
    missing = []
    for month in range(1, 7):
        for group in range(4):
            for item in range(6):
                row_number = len(rows) + 2
                k = number * 101 + month * 37 + group * 19 + item * 11
                item_name = {"sales": "Product", "inventory": "SKU", "budget": "Cost center", "returns": "Product", "workforce": "Role", "projects": "Workstream"}[domain] + f" {item + 1:02}"
                r = dict(RowID=f"{ident}-{len(rows)+1:04}", Period=f"2026-{month:02}", Date=date(f"2026-{month:02}-{1+(group*6+item)%27:02}"), Group=GROUPS[group], Item=item_name)
                if domain == "sales":
                    units, price, cost = 18 + k % 43, 45 + k % 89, 24 + k % 29
                    r.update(Units=units, UnitPriceEUR=price, UnitCostEUR=cost, RevenueEUR=None, CostEUR=None, BudgetEUR=(units + 3) * (price + 2), Status="Final")
                    f = {"RevenueEUR": f'=IF(G{row_number}="","",F{row_number}*G{row_number})', "CostEUR": f"=F{row_number}*H{row_number}"}
                    missing_field = "UnitPriceEUR"
                elif domain == "inventory":
                    r.update(OpeningUnits=160+k%90, ReceivedUnits=40+k%55, ShippedUnits=55+k%95, ClosingUnits=None, UnitCostEUR=12+k%37, InventoryValueEUR=None, ReorderUnits=95+k%45, Status="Counted")
                    f = {"ClosingUnits": f'=IF(G{row_number}="","",F{row_number}+G{row_number}-H{row_number})', "InventoryValueEUR": f'=IF(I{row_number}="","",I{row_number}*J{row_number})'}
                    missing_field = "ReceivedUnits"
                elif domain == "budget":
                    r.update(ActualEUR=1100+(k*13)%3500, BudgetEUR=1400+(k*7)%2900, VarianceEUR=None, Owner=OWNERS[group], Status="Accrued" if item==5 else "Posted")
                    f = {"VarianceEUR": f'=IF(F{row_number}="","",F{row_number}-G{row_number})'}
                    missing_field = "ActualEUR"
                elif domain == "returns":
                    sold, returned, price = 80+k%150, 1+k%13, 20+k%90
                    r.update(UnitsSold=sold, UnitsReturned=returned, NetSalesEUR=sold*price, RefundEUR=returned*price, ReturnRate=None, Reason=["Damaged", "Wrong item", "Customer choice"][item%3])
                    f = {"ReturnRate": f'=IF(F{row_number}=0,"",G{row_number}/F{row_number})'}
                    missing_field = "RefundEUR"
                elif domain == "workforce":
                    available, worked = 140+k%29, 110+k%48
                    r.update(Headcount=1, HoursAvailable=available, HoursWorked=worked, CostEUR=worked*(20+k%31), Utilization=None, Status="Submitted")
                    f = {"Utilization": f'=IF(H{row_number}="","",IF(G{row_number}=0,"",H{row_number}/G{row_number}))'}
                    missing_field = "HoursWorked"
                else:
                    r.update(PlannedHours=30+k%65, ActualHours=28+k%85, HourlyRateEUR=35+k%50, ActualCostEUR=None, Owner=OWNERS[group], DueDate=date(f"2026-{month:02}-{5+(group*4+item)%23:02}"), Status=["Planned", "In progress", "Complete"][item%3])
                    f = {"ActualCostEUR": f'=IF(G{row_number}="","",G{row_number}*H{row_number})'}
                    missing_field = "ActualHours"
                # Five workbooks have one unknown June source measure. The blank
                # must survive, and totals are explicitly known subtotals.
                if number % 4 == 0 and month == 6 and group == 3 and item == 5:
                    r[missing_field] = None
                    missing.append({"sheet": "Ledger", "cell": f"{column(headers.index(missing_field)+1)}{row_number}", "row_id": r["RowID"], "field": missing_field, "period": r["Period"], "group": r["Group"]})
                rows.append(r)
                formulas.extend({"cell": f"{column(headers.index(key)+1)}{row_number}", "formula": value} for key, value in f.items())
    duplicates = []
    if number % 5 == 0:
        source_index = 130
        duplicate = dict(rows[source_index])
        rows.append(duplicate)
        original_row, duplicate_row = source_index + 2, len(rows) + 1
        duplicates.append({"row_id": duplicate["RowID"], "original_row": original_row, "duplicate_row": duplicate_row, "rule": "Identical RowID and fields; keep first once"})
        for formula in list(formulas):
            if formula["cell"].lstrip("ABCDEFGHIJKLMNOPQRSTUVWXYZ") == str(original_row):
                import re
                formulas.append({"cell": re.sub(r"\d+$", str(duplicate_row), formula["cell"]), "formula": re.sub(r"(?<=[A-Z])"+str(original_row)+r"\b", str(duplicate_row), formula["formula"])})
    unique = list({r["RowID"]: r for r in rows}.values())
    monthly = [dict(period=period, **aggregate(domain, [r for r in unique if r["Period"]==period])) for period in MONTHS]
    current, previous = monthly[-1], monthly[-2]
    by_group = [dict(group=group, **aggregate(domain, [r for r in unique if r["Period"]==MONTHS[-1] and r["Group"]==group])) for group in GROUPS]
    header_to_column = {name: column(index+1) for index,name in enumerate(headers)}
    probe_field = {"sales":"UnitPriceEUR", "inventory":"ReceivedUnits", "budget":"ActualEUR", "returns":"RefundEUR", "workforce":"HoursWorked", "projects":"ActualHours"}[domain]
    probe_index = next(i for i,r in enumerate(unique) if r["Period"]=="2026-06" and r[probe_field] is not None)
    changed = [dict(r) for r in unique]
    changed[probe_index][probe_field] += 10
    probe = {"description":"Change one source input in an isolated copied workbook, recalculate natively, verify dependent output and chart values change", "sheet":"Ledger", "cell":f"{header_to_column[probe_field]}{probe_index+2}", "delta":10, "original_value":unique[probe_index][probe_field], "new_value":changed[probe_index][probe_field], "after_current":aggregate(domain,[r for r in changed if r["Period"]=="2026-06"]), "native_execution_required":True}
    probe["dependent_cells"]=[{"sheet":"Ledger","cell":f"{header_to_column[field]}{probe_index+2}","expected":value} for field,value in formula_values(domain,changed[probe_index]).items()]
    source = dict(id=ident, path=f"inputs/excel/{ident}.xlsx", domain=domain, company=company, primary_sheet="Ledger", columns=headers, rows=[[r[h] for h in headers] for r in rows], formulas=formulas, primary_label=primary_label, secondary_label=secondary_label, ratio_label=ratio_label, primary_field=primary_field, secondary_field=secondary_field, budget_field=budget_field, units_field=units_field, header_to_column=header_to_column, historical_months=MONTHS[:5], historical_complete=all(x["complete"] for x in monthly[:5]))
    catalog = {k:source[k] for k in ("id","path","domain","company","primary_sheet","columns","primary_label","secondary_label","ratio_label","primary_field","secondary_field","budget_field","units_field")}
    period_aggregation = "Latest-month snapshot: do not sum closing inventory or inventory valuation across months" if domain=="inventory" else "Additive monthly activity; compute rates from aggregate numerators and denominators, never average row percentages"
    total=current if domain=="inventory" else aggregate(domain,unique)
    catalog.update(row_count=len(rows), unique_row_count=len(unique), source_sha256=digest(source), aggregation={"period":period_aggregation,"group":"Sum unique RowID observations within the requested month", "headcount":"Monthly workforce headcount is not a count of distinct employees across six months" if domain=="workforce" else None}, facts={"current_period":"2026-06", "previous_period":"2026-05", "monthly":monthly, "current":current, "previous":previous, "total":total, "by_group":by_group, "missing_cells":missing, "duplicate_rows":duplicates}, formula_probes=[probe], source_formulas=formulas, chart_specs=[{"sheet":"History", "type":"line", "categories":MONTHS[:5], "primary_series":primary_label, "secondary_series":secondary_label, "update_target_periods":MONTHS, "editable":True}], required_checks=["preserve_source", "native_formulas", "native_recalculation", "dependency_perturbation", "current_period_facts", "native_chart_bindings", "no_formula_errors", "missing_is_not_zero", "deduplicate_RowID"])
    catalog["native_formula_cells"]=[{"sheet":"Ledger","cell":f"{header_to_column[field]}{i+2}","expected":value,"tolerance":1e-8} for i,row in enumerate(rows) for field,value in formula_values(domain,row).items()]
    return source, catalog


def deck_spec(number: int, workbook: dict) -> tuple[dict, dict]:
    ident = f"PPT{number:02}"
    workbook_number = (number - 1) % 20 + 1
    slide_count = 6 + (number - 1) % 7
    # PP01 is the development repair smoke case. Its source must actually
    # contain repairable defects rather than match the clean reference.
    repair = number == 1 or number % 3 == 0
    defects = []
    if repair:
        defects = [
            {"slide":2,"kind":"out_of_bounds_chart","object":"monthly-chart","expected_repair":"Fit editable chart inside the content region; preserve all series and six categories"},
            {"slide":4,"kind":"text_overflow","object":"commentary","expected_repair":"Reflow complete text into readable native text; do not discard facts"},
            {"slide":3,"kind":"wrong_theme_accent","object":"table-header","actual":"E91E63","expected":"4F81BD"},
        ]
    source = {"id":ident,"path":f"inputs/powerpoint/{ident}.pptx", "domain":workbook["domain"],"company":workbook["company"],"linked_workbook":f"WB{workbook_number:02}","slide_count":slide_count,"repair_case":repair,"primary_label":workbook["primary_label"],"secondary_label":workbook["secondary_label"],"ratio_label":workbook["ratio_label"],"monthly":workbook["facts"]["monthly"],"current":workbook["facts"]["current"],"by_group":workbook["facts"]["by_group"],"reference_path":f"evaluator-only/powerpoint-references/{ident}-reference.pptx"}
    chart_slides = [2] + ([7] if slide_count >= 7 else []) + ([10] if slide_count >= 10 else [])
    table_slides = [3,5] + ([8] if slide_count >= 8 else []) + ([11] if slide_count >= 11 else [])
    oracle = {k:source[k] for k in ("id","path","domain","company","linked_workbook","slide_count","reference_path","primary_label","secondary_label","ratio_label")}
    series=[{"name":workbook["primary_label"],"values":[x["primary"] for x in source["monthly"]]}]
    if workbook["domain"]!="inventory":
        series.append({"name":workbook["secondary_label"],"values":[x["secondary"] for x in source["monthly"]]})
    oracle.update(source_sha256=digest(source), native_chart_slides=chart_slides, native_table_slides=table_slides, intentional_defects=defects, theme=THEME, theme_ref="evaluator-only/SamsungMD2.theme.json", facts=workbook["facts"], chart_specs=[{"slide":2,"name":"monthly-chart","categories":MONTHS,"series":series}], required_checks=["source_content_preserved","exact_slide_count","native_charts_and_tables","chart_series_and_categories","no_overflow_or_unintended_overlap","Samsung_theme_tokens","speaker_notes_source_coverage","independent_visual_review"], model_quality_proven=False)
    return source, oracle


def build_data(root: Path) -> dict:
    workbooks, workbook_catalog = [], []
    for number in range(1,21):
        source, oracle = workbook_spec(number)
        workbooks.append(source)
        workbook_catalog.append(oracle)
    decks, deck_catalog = [], []
    for number in range(1,31):
        source, oracle = deck_spec(number,workbook_catalog[(number-1)%20])
        decks.append(source)
        deck_catalog.append(oracle)
    # Build payload is PRIVATE; it includes expected source content and source
    # defect instructions, and must never be included in a model-visible kit.
    payload = {"schema_version":1,"version":VERSION,"workbooks":workbooks,"presentations":decks,"theme":THEME}
    write_json(root/".build"/"office_payload.json",payload)
    catalog = {"schema_version":1,"version":VERSION,"seed":"fixed integer arithmetic; no randomness or clock dependencies","evaluation_policy":{"model_visible_roots":["inputs"],"evaluator_only_roots":["evaluator-only",".build"],"oracles_are_not_model_inputs":True,"package_checks_are_not_native_or_model_acceptance":True,"native_recalculation_required":True,"independent_visual_review_required":True},"workbooks":workbook_catalog,"presentations":deck_catalog}
    write_json(root/"evaluator-only"/"office_catalog.json",catalog)
    write_json(root/"evaluator-only"/"SamsungMD2.theme.json",{"schema_version":1,"name":THEME["name"],"width":960,"height":540,"fonts":["Arial","Arial Narrow","Calibri","Samsung Sharp Sans Bold"],"font_roles":{"title":["Samsung Sharp Sans Bold","Arial"],"body":["Arial"],"table":["Arial Narrow","Arial"],"caption":["Arial Narrow","Arial"],"footer":["Arial Narrow","Arial"],"folio":["Calibri","Arial"]},"palette":list(THEME["colors"].values())+["FFFFFF","000000","7F7F7F","A6A6A6","202A35","596674","D7DDE3","D4D4D4"],"source":[THEME["source"],"src/Scribble/Office/MetoTheme.cs","src/Scribble/Office/PresentationDraftWriter.Samsung.cs"],"title_font_pt":24,"body_font_pt":18,"minimum_body_font_pt":14,"minimum_table_font_pt":7.5,"minimum_footer_font_pt":7,"minimum_folio_font_pt":8,"human_fidelity_review_required":True})
    # The safe index deliberately excludes computed facts and repair answers.
    safe = {"schema_version":1,"workbooks":[{k:w[k] for k in ("id","path","domain","company","primary_sheet","columns")} for w in workbooks],"presentations":[{k:d[k] for k in ("id","path","domain","company","linked_workbook","slide_count")} for d in decks]}
    write_json(root/"inputs"/"office_sources.json",safe)
    return catalog


def verify(root: Path, catalog: dict) -> dict:
    """Read exported ZIP/XML, not authoring objects; no Office/model involved."""
    ns = {"s":"http://schemas.openxmlformats.org/spreadsheetml/2006/main","p":"http://schemas.openxmlformats.org/presentationml/2006/main","c":"http://schemas.openxmlformats.org/drawingml/2006/chart","a":"http://schemas.openxmlformats.org/drawingml/2006/main"}
    checks = []
    for item in catalog["workbooks"] + catalog["presentations"]:
        file = root/item["path"]
        if not file.is_file():
            raise RuntimeError(f"Missing {file}")
        with zipfile.ZipFile(file) as package:
            if package.testzip():
                raise RuntimeError(f"Corrupt package {file}")
            names = package.namelist()
            if any("evaluator-only" in n for n in names):
                raise RuntimeError(f"Evaluator data embedded in {file}")
            if file.suffix == ".xlsx":
                sheets = [n for n in names if n.startswith("xl/worksheets/sheet") and n.endswith(".xml")]
                parsed = [ET.fromstring(package.read(n)) for n in sheets]
                formula_count = sum(len(x.findall(".//s:f",ns)) for x in parsed)
                errors = [c.attrib.get("r") for x in parsed for c in x.findall(".//s:c",ns) if c.attrib.get("t")=="e"]
                if formula_count < 1 or errors:
                    raise RuntimeError(f"Formula/cache failure in {file}: formulas={formula_count}, errors={errors}")
                charts = [n for n in names if (n.startswith("xl/charts/chart") or n.startswith("xl/drawings/charts/chart")) and n.endswith(".xml")]
                if len(charts)<1:
                    raise RuntimeError(f"No native workbook chart in {file}")
                checks.append({"id":item["id"],"sheets":len(sheets),"formulas":formula_count,"formula_cache_errors":errors,"charts":len(charts),"sha256":hashlib.sha256(file.read_bytes()).hexdigest()})
            else:
                slides = [n for n in names if n.startswith("ppt/slides/slide") and n.endswith(".xml")]
                charts = [n for n in names if (n.startswith("ppt/charts/chart") or n.startswith("ppt/slides/charts/chart")) and n.endswith(".xml")]
                tables = sum(len(ET.fromstring(package.read(n)).findall(".//a:tbl",ns)) for n in slides)
                if len(slides)!=item["slide_count"] or len(charts)!=len(item["native_chart_slides"]) or tables!=len(item["native_table_slides"]):
                    raise RuntimeError(f"Native structure count mismatch in {file}: slides={len(slides)}, charts={len(charts)}, tables={tables}")
                # Each editable native chart must carry its own workbook data.
                external_data=[ET.fromstring(package.read(n)).find(".//c:externalData",ns) is not None for n in charts]
                if not all(external_data):
                    raise RuntimeError(f"Missing embedded chart data in {file}")
                checks.append({"id":item["id"],"slides":len(slides),"native_charts":len(charts),"native_tables":tables,"intentional_defects":len(item["intentional_defects"]),"sha256":hashlib.sha256(file.read_bytes()).hexdigest()})
    result={"schema_version":1,"scope":"OOXML package/readback only; no native Office or model-quality proof","workbooks":20,"presentations":30,"checks":checks}
    write_json(root/"evaluator-only"/"office_package_verification.json",result)
    return result


def main() -> int:
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output",type=Path,default=DEFAULT_ROOT)
    parser.add_argument("--data-only",action="store_true")
    parser.add_argument("--verify",action="store_true",help="Verify all 50 exported sources; optionally use with --data-only")
    parser.add_argument("--only",default="")
    parser.add_argument("--runtime-node",default=os.getenv("RUNTIME_NODE"))
    parser.add_argument("--node-modules",default=os.getenv("RUNTIME_NODE_MODULES"))
    parser.add_argument("--presentations-skill",default=os.getenv("PRESENTATIONS_SKILL"))
    parser.add_argument("--skip-previews",action="store_true",help="Skip private PNG previews; not valid for final visual acceptance")
    args=parser.parse_args()
    root=args.output.resolve()
    root.mkdir(parents=True,exist_ok=True)
    catalog=build_data(root)
    if not args.data_only:
        if not all([args.runtime_node,args.node_modules,args.presentations_skill]):
            parser.error("Authoring requires --runtime-node, --node-modules and --presentations-skill from the bundled runtime")
        build=root/".build"
        modules=build/"node_modules"
        if not modules.exists():
            if os.name=="nt":
                # Fixed-path junction creation only; no shell concatenation.
                junction_script=build/"link-runtime.ps1"
                junction_script.write_text('param([string]$LinkPath,[string]$TargetPath)\nNew-Item -ItemType Junction -Path $LinkPath -Target $TargetPath -ErrorAction Stop | Out-Null\n',encoding="utf-8")
                subprocess.run(["powershell","-NoProfile","-ExecutionPolicy","Bypass","-File",str(junction_script),"-LinkPath",str(modules),"-TargetPath",str(Path(args.node_modules).resolve())],check=True)
            else:
                modules.symlink_to(Path(args.node_modules).resolve(),target_is_directory=True)
        shutil.copyfile(Path(__file__).with_name("author_office.mjs"),build/"author_office.mjs")
        shutil.copyfile(Path(__file__).with_name("author_hero_inputs.mjs"),build/"author_hero_inputs.mjs")
        env=os.environ.copy()
        env.update(PRESENTATIONS_SKILL=str(Path(args.presentations_skill).resolve()),RUNTIME_PYTHON=sys.executable,RUNTIME_NODE=str(Path(args.runtime_node).resolve()),RUNTIME_NODE_MODULES=str(Path(args.node_modules).resolve()),OFFICE_STRESS_ROOT=str(root),OFFICE_STRESS_ONLY=args.only,OFFICE_STRESS_PREVIEWS="0" if args.skip_previews else "1")
        subprocess.run([args.runtime_node,str(build/"author_office.mjs")],env=env,check=True,cwd=build)
        subprocess.run([args.runtime_node,str(build/"author_hero_inputs.mjs")],env=env,check=True,cwd=build)
        subprocess.run([sys.executable,str(Path(__file__).with_name("author_hero_pdf.py")),"--output",str(root)],check=True)
    if args.verify:
        result=verify(root,catalog)
        print(f"Verified {len(result['checks'])} Office sources; native and model acceptance remain separate.")
    print(root/"evaluator-only"/"office_catalog.json")
    return 0


if __name__=="__main__":
    raise SystemExit(main())
