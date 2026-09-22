"""Offline checks for sanitized local diagnostic usage summaries."""
import json
import tempfile
import unittest
from pathlib import Path
from zipfile import ZipFile

from summarize_usage import summarize, summarize_trace_export


class TraceExportSummaryTests(unittest.TestCase):
    def test_existing_suite_manifest_route_still_counts_zip_evidence(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            evidence = root / "case.zip"
            with ZipFile(evidence, "w") as archive:
                archive.writestr("timeline.jsonl", "\n".join((
                    json.dumps({"stage": "inference_request"}),
                    json.dumps({"stage": "inference_response", "detail": {
                        "model": "synthetic", "response": json.dumps({
                            "provider": "local", "usage": {
                                "prompt_tokens": 7, "completion_tokens": 2, "cost": 0.0,
                            },
                        }),
                    }}),
                )))
            suite = root / "suite.json"
            suite.write_text(json.dumps({"suite": {"id": "suite"}, "cases": [
                {"id": "XA01", "status": "finished", "evidence": str(evidence)},
            ]}), encoding="utf-8")
            result = summarize(suite)
            self.assertEqual(result["response_count"], 1)
            self.assertEqual(result["cases"][0]["calls"][0]["prompt_tokens"], 7)
            self.assertTrue(result["cost_complete"])

    def test_counts_requests_and_omits_prompt_and_response_text(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "run.json").write_text(json.dumps({
                "suite_id": "suite", "case_id": "PP01", "run_id": "run",
                "assembly_sha256": "hash", "selected_model": "synthetic",
                "prompt": "PRIVATE PROMPT", "status": "finished",
                "trace_complete": True,
            }), encoding="utf-8")
            events = [
                {"stage": "inference_request", "detail": {"body": "PRIVATE REQUEST"}},
                {"stage": "inference_response", "detail": {
                    "model": "synthetic", "http_status": 200,
                    "response": json.dumps({
                        "provider": "local", "choices": [{"message": {"content": "PRIVATE RESPONSE"}}],
                        "usage": {"prompt_tokens": 17, "completion_tokens": 3, "cost": 0.0},
                    }),
                }},
                {"stage": "inference_request", "detail": {"body": "PRIVATE REQUEST 2"}},
            ]
            (root / "timeline.jsonl").write_text(
                "\n".join(json.dumps(event) for event in events) + "\n", encoding="utf-8")
            result = summarize_trace_export(root)
            self.assertEqual((result["request_count"], result["response_count"]), (2, 1))
            self.assertEqual((result["prompt_tokens"], result["completion_tokens"]), (17, 3))
            self.assertEqual(result["requests_without_response"], 1)
            self.assertFalse(result["cost_complete"])
            self.assertEqual(result["provider_counts"], {"local": 1})
            serialized = json.dumps(result)
            for secret in ("PRIVATE PROMPT", "PRIVATE REQUEST", "PRIVATE RESPONSE"):
                self.assertNotIn(secret, serialized)


if __name__ == "__main__":
    unittest.main()
