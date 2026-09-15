"""Pure catalog contract tests; no Office processes, sessions, or model calls."""
import copy
import tempfile
import unittest
from pathlib import Path

import build_catalog
import generate_mail
import generate_office


class CatalogContractTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.directory = tempfile.TemporaryDirectory(prefix="scribble-catalog-contract-")
        cls.root = Path(cls.directory.name)
        cls.office = generate_office.build_data(cls.root)
        records, mail_oracles = generate_mail.build_records()
        cls.mail = {"search_tasks": generate_mail.build_tasks(records, mail_oracles)}
        build_catalog.write(cls.root, "evaluator-only/mail_catalog.json", cls.mail)
        build_catalog.write(cls.root, "evaluator-only/hero_inputs.json", {
            "korean": {"path": "inputs/excel/KoreanOperations.xlsx", "sheets": [{"name": "Operations", "cells": [{"row": 1, "column": 1, "text": "Category"}]}]},
            "word": {"path": "inputs/excel/EightSheetWorkbook.xlsx", "tables": [[['Header'], ['Value']]]},
        })
        build_catalog.write(cls.root, "evaluator-only/hero_pdf.json", {"path": "inputs/pdf/ExecutiveRiskReport130.pdf",
            "required_facts": ["late-page fact"], "terminal_marker": "ORION-PAGE-130"})
        cls.summary = build_catalog.build(cls.root)
        cls.cases = build_catalog.read(cls.root / "operator/cases.json")
        cls.oracles = {c["id"]: build_catalog.read(cls.root / c["oracle_ref"]) for c in cls.cases}

    @classmethod
    def tearDownClass(cls):
        cls.directory.cleanup()

    def test_complete_case_scope_and_no_answer_leak(self):
        self.assertEqual(self.summary["case_count"], 200)
        self.assertEqual(self.summary["model_requests"], 0)
        self.assertEqual(self.summary["execution_status"], "not_run")
        visible = {path for case in self.cases for path in case["inputs"]}
        self.assertEqual({p for p in visible if p.endswith(".xlsx")},
                         {f"inputs/excel/WB{i:02}.xlsx" for i in range(1, 21)} |
                         {"inputs/excel/KoreanOperations.xlsx", "inputs/excel/EightSheetWorkbook.xlsx"})
        self.assertEqual({p for p in visible if p.endswith(".pptx")}, {f"inputs/powerpoint/PPT{i:02}.pptx" for i in range(1, 31)})
        poisoned = copy.deepcopy(self.cases)
        poisoned[0]["inputs"].append(poisoned[0]["oracle_ref"])
        with self.assertRaisesRegex(ValueError, "leaked"):
            build_catalog.validate(poisoned, self.oracles)
        poisoned = copy.deepcopy(self.cases)
        poisoned[0]["expected_ids"] = ["MAIL0001"]
        with self.assertRaisesRegex(ValueError, "oracle leaked"):
            build_catalog.validate(poisoned, self.oracles)

    def test_corpus_boundary_search_and_independent_scenario_values(self):
        all_mail = self.oracles["OL01"]["checks"][0]
        self.assertEqual(all_mail["expected_count"], 500)
        self.assertEqual(set(all_mail["expected_ids"]), {f"MAIL{i:04}" for i in range(1, 501)})
        no_mail = self.oracles["OL10"]["checks"][0]
        self.assertEqual(no_mail["expected_ids"], [])
        self.assertEqual(no_mail["expected_count"], 0)
        cases_by_id = {case["id"]: case for case in self.cases}
        self.assertIn("Total matches: N", cases_by_id["OL01"]["prompt"])
        for index, workbook in enumerate(self.office["workbooks"]):
            monthly_rules = self.oracles[f"EX{index*3+1:02}"]["checks"]
            monthly_cells = {r["cell"]: r for r in monthly_rules if r["kind"] == "numeric_cell"}
            for row, facts in enumerate(workbook["facts"]["monthly"], 4):
                self.assertEqual(monthly_cells[f"D{row}"]["expected"], facts["budget"])
                self.assertEqual(monthly_cells[f"E{row}"]["expected"], facts["units"])
            group_rules = self.oracles[f"EX{index*3+2:02}"]["checks"]
            group_cells = {r["cell"]: r for r in group_rules if r["kind"] in ("numeric_cell", "cell_text")}
            for row, facts in enumerate(workbook["facts"]["by_group"], 4):
                self.assertEqual(group_cells[f"D{row}"]["expected"], facts["budget"])
                self.assertEqual(group_cells[f"E{row}"]["expected"], facts["units"])
                if workbook["ratio_label"] and facts["ratio"] is None:
                    self.assertEqual(group_cells[f"F{row}"], {"kind": "cell_text", "sheet": "Scribble Draft*", "cell": f"F{row}", "expected": "incomplete"})
            if index < 19:
                rules = self.oracles[f"EX{index*3+3:02}"]["checks"]
                scenario = next(r for r in rules if r["kind"] == "numeric_cell" and r["cell"] == "C4")
                self.assertEqual(scenario["expected"], workbook["formula_probes"][0]["after_current"]["primary"])
                self.assertTrue(scenario["formula_required"])
        cases_by_id = {case["id"]: case for case in self.cases}
        self.assertTrue(cases_by_id["EX60"]["allow_source_edit"])
        self.assertEqual(cases_by_id["OL70"]["host"], "Chrome")
        self.assertEqual(cases_by_id["OL70"]["browser_allowed_hosts"], ["www.samsungtradein.ae", "samsungtradein.ae"])
        self.assertEqual(cases_by_id["OL70"]["browser_start_url"], "https://www.samsungtradein.ae/ae-en/")
        self.assertTrue(any(rule["kind"] == "word_tables" for rule in self.oracles["XA20"]["checks"]))
        # Incomplete June stock is a known subtotal; it is not zero or an
        # annual sum, and the output must disclose that it is incomplete.
        wb08 = self.office["workbooks"][7]
        self.assertFalse(wb08["facts"]["current"]["complete"])
        rules = self.oracles["EX22"]["checks"]
        june = next(r for r in rules if r["kind"] == "numeric_cell" and r["cell"] == "B9")
        self.assertEqual(june["expected"], wb08["facts"]["current"]["primary"])
        self.assertTrue(any(r["kind"] == "required_text" and "incomplete" in r["values"] for r in rules))

    def test_incomplete_corpus_cannot_be_sealed_or_claim_execution(self):
        with self.assertRaisesRegex(ValueError, "incomplete corpus"):
            build_catalog.manifest(self.root)
        self.assertFalse((self.root / "manifest.json").exists())
        self.assertEqual(build_catalog.read(self.root / "operator/stress-plan.json")["execution_status"], "not_run")
        with self.assertRaises(ValueError):
            build_catalog.safe(self.root, "../outside.json")

    def test_chart_requirements_target_the_requested_output_host(self):
        charts = [(case_id, rule) for case_id, oracle in self.oracles.items()
                  for rule in oracle["checks"] if rule["kind"] == "native_chart"]
        self.assertEqual(len(charts), 101)
        for case_id, rule in charts:
            expected = ("Excel", "xlsx") if case_id.startswith("EX") else ("PowerPoint", "pptx")
            self.assertEqual((rule["host"], rule["artifact_extension"]), expected, case_id)
        for host, extension in [(None, "pptx"), ("PowerPoint", None), ("PowerPoint", "xlsx")]:
            poisoned = copy.deepcopy(self.oracles)
            rule = next(r for r in poisoned["XA01"]["checks"] if r["kind"] == "native_chart")
            rule.update(host=host, artifact_extension=extension)
            with self.assertRaisesRegex(ValueError, "intended host"):
                build_catalog.validate(self.cases, poisoned)


if __name__ == "__main__":
    unittest.main()
