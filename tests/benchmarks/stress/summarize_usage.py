"""Summarize recorded provider usage from real Scribble evidence; no API calls.

Missing provider usage is unknown, never zero. Source messages and credentials
are deliberately excluded from this report.
"""
from __future__ import annotations
import argparse
import json
from pathlib import Path
from zipfile import ZipFile


def summarize(suite_path: Path):
    manifest = json.loads(suite_path.read_text(encoding="utf-8-sig"))
    records = []
    for case in manifest["cases"]:
        calls = []
        requests = 0
        evidence = case.get("evidence")
        if evidence:
            path = Path(evidence).resolve()
            if not path.is_relative_to(suite_path.parent.resolve()):
                raise ValueError("Evidence escapes the suite directory")
            with ZipFile(path) as archive:
                for line in archive.read("timeline.jsonl").decode("utf-8-sig").splitlines():
                    event = json.loads(line)
                    if event.get("stage") == "inference_request":
                        requests += 1
                    if event.get("stage") != "inference_response":
                        continue
                    detail = event.get("detail", {})
                    try:
                        response = json.loads(detail.get("response", "{}"))
                    except (ValueError, TypeError):
                        response = {}
                    usage = response.get("usage") or {}
                    calls.append({"model": detail.get("model"), "status": detail.get("http_status"),
                                  "prompt_tokens": usage.get("prompt_tokens"),
                                  "completion_tokens": usage.get("completion_tokens"),
                                  "cost_usd": usage.get("cost"), "provider": response.get("provider")})
        records.append({"id": case["id"], "status": case["status"], "calls": calls,
                        "response_count": len(calls),
                        "request_count": requests,
                        "requests_without_response": max(0, requests - len(calls)),
                        "reported_cost_usd": sum(c["cost_usd"] for c in calls if c["cost_usd"] is not None),
                        "responses_without_reported_cost": sum(c["cost_usd"] is None for c in calls),
                        "execution_observed": bool(calls)})
    return {"schema": 1, "suite_id": manifest["suite"]["id"], "cases": records,
            "reported_cost_usd": sum(r["reported_cost_usd"] for r in records),
            "cost_complete": all(not r["responses_without_reported_cost"] and not r["requests_without_response"] for r in records),
            "response_count": sum(r["response_count"] for r in records),
            "note": "Provider-reported inference costs only; excludes credit purchase fees. Missing usage is unknown. No model calls were made by this summarizer."}


def summarize_trace_export(trace_dir: Path):
    """Summarize a local diagnostic export without copying prompt or response text."""
    run = json.loads((trace_dir / "run.json").read_text(encoding="utf-8-sig"))
    requests = 0
    calls = []
    with (trace_dir / "timeline.jsonl").open(encoding="utf-8-sig") as stream:
        for line in stream:
            event = json.loads(line)
            if event.get("stage") == "inference_request":
                requests += 1
            if event.get("stage") != "inference_response":
                continue
            detail = event.get("detail") or {}
            try:
                response = json.loads(detail.get("response") or "{}")
            except (ValueError, TypeError):
                response = {}
            usage = response.get("usage") or {}
            calls.append({"model": detail.get("model"), "status": detail.get("http_status"),
                          "prompt_tokens": usage.get("prompt_tokens"),
                          "completion_tokens": usage.get("completion_tokens"),
                          "cost_usd": usage.get("cost"), "provider": response.get("provider")})
    return {"schema": 1, "suite_id": run.get("suite_id"), "case_id": run.get("case_id"),
            "run_id": run.get("run_id"), "run_status": run.get("status"),
            "trace_complete": run.get("trace_complete"),
            "assembly_version": run.get("assembly_version"),
            "assembly_sha256": run.get("assembly_sha256"),
            "selected_model": run.get("selected_model"),
            "started_utc": run.get("started_utc"), "finished_utc": run.get("finished_utc"),
            "request_count": requests, "response_count": len(calls),
            "requests_without_response": max(0, requests - len(calls)),
            "prompt_tokens": sum(c["prompt_tokens"] for c in calls if c["prompt_tokens"] is not None),
            "completion_tokens": sum(c["completion_tokens"] for c in calls if c["completion_tokens"] is not None),
            "responses_without_token_usage": sum(c["prompt_tokens"] is None or c["completion_tokens"] is None for c in calls),
            "reported_cost_usd": sum(c["cost_usd"] for c in calls if c["cost_usd"] is not None),
            "responses_without_reported_cost": sum(c["cost_usd"] is None for c in calls),
            "cost_complete": requests == len(calls) and all(c["cost_usd"] is not None for c in calls),
            "provider_counts": {provider: sum(c["provider"] == provider for c in calls)
                                for provider in sorted({c["provider"] for c in calls if c["provider"]})},
            "calls": calls,
            "note": "Provider-reported inference usage only. Missing usage is unknown. No model calls were made."}


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("suite", type=Path, help="Suite manifest or local trace-export directory")
    args = parser.parse_args()
    result = summarize_trace_export(args.suite) if args.suite.is_dir() else summarize(args.suite)
    target = (args.suite if args.suite.is_dir() else args.suite.parent) / "model-usage.json"
    target.write_text(json.dumps(result, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"path": str(target), "response_count": result["response_count"],
                      "reported_cost_usd": result["reported_cost_usd"], "cost_complete": result["cost_complete"]}))
