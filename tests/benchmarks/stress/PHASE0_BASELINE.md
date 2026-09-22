# Reliability phase 0 baseline (in progress)

Prepared 23 September 2026 from `bffd25e` on `codex/stress-suite-200`.
The authoritative plan is `RELIABILITY_PLAN.md` at `9792f2de` on
`claude/plan-review-corrections-f5p3ce` (PR #21); its companion review explains
the corrected baseline. The owner authorized stacking on PR #20. This branch
contains no paid model run.

## Reconciliation

Five commits follow planning baseline `fad5cdd`: `ebaad7b`, `b873435`,
`037c7eb`, `5fbe0a9`, and `bffd25e`. The first two alter stress-budget retry
handling; the next two alter PowerPoint review/repair and their guardrails;
the last raises the authorized hard cap to $40. The PowerPoint files therefore
overlap phase 3 and must be tested against the new repair contract before
their heuristics are removed. `LegacySamsung/` is untouched.

The latest checkout checkpoint says candidate 2.0.398 passed CI and installed,
then PP01 stopped during `SnapshotExternalKit` before any case, model request,
or document write (`suite-20260922-194429-80f7e1bb`). It was not a pass or a
model failure. No further paid live run is authorized before phases 0–3 pass
offline. The $40 key had $14.47 remaining at 19:35 UTC on 22 September; this
is a dated observation, not a current balance.

## Archived traces and sanitized usage

Raw traces are archived only outside Git at
`Documents/Scribble Reliability Evidence/phase0-baseline-traces-20260922.zip`.
SHA-256: `0ffa5d14ede74b86095606929866fe3d801166df9df1193abe0e2f7dc906bb93`.
The archive contains 93 files from the three exports below. The adjacent
`baselines/*-model-usage.json` files contain model/provider/status/token/cost
metadata only. `summarize_usage.py` derives them from `run.json` and
`timeline.jsonl` without copying prompts or response text.

| Export / fixture suite | Assembly SHA-256 | Requests / responses | Prompt / completion tokens | Providers | Reported cost |
| --- | --- | ---: | ---: | ---: | ---: |
| XA01 2.0.385 / `suite-20260922-163550-3fa83055` | `1a043b9f8e9d2e02ce22c2c75d0c058208b08b36afee13ebbae4809870d3b062` | 26 / 26 | 356,054 / 59,712 | 7 | $0.2000213108 |
| PP01 2.0.388 / `suite-20260922-165042-738f6117` | `a6c6b88ce459a671d348270cf21ed8768ae875aee4e4c45bfbea77c6772661f4` | 87 / 87 | 3,387,583 / 57,211 | 13 | $0.5176867968 |
| PP01 2.0.391 / `suite-20260922-172330-f0091961` | `299ebfc778a641e042218d627042437095f70984ec63fd627f2b57bc5fff2a14` | 74 / 74 | 3,404,690 / 70,867 | 6 | $0.48670295 |

Every recorded response in these exports has token and cost usage. The 2.0.391
run's trace is marked incomplete, so the 74 responses do not establish a
completed PP01 task. Token totals include repeated and cached prompt input;
they are not unique-context sizes. Costs are provider-reported inference cost,
not a full account balance reconciliation.

## Pilot configuration and local feasibility

The pilot test model remains hosted OpenRouter `qwen/qwen3.8-27b`; the configured
Scribble endpoint was verified as `https://openrouter.ai/api/v1`. Current code
restricts only tool-bearing requests to a six-provider allow-list and permits
fallbacks. Tool-less reviewers can route to other providers. The pilot should
pin a provider or record each serving provider before comparing performance.

The local deployment candidate is Ollama `qwen3.8:27b` on
`127.0.0.1:11434`. A bounded 4,096-context vision probe with a synthetic red
square returned `red` (37 prompt tokens, 2 generated tokens, 27.35 seconds).
Free RAM fell from 13.43 GiB to 0.35 GiB. The model was unloaded and the exact
server process started for the probe was stopped; free RAM recovered to 14.5
GiB. Larger-context, strict-schema, and tool-call probes remain unverified.
Do not run them with the current user-owned Excel session open and this memory
headroom. A name-based vision-capability flag is not a substitute for this
measured result or for the missing production-path certification.

## Candidate model-call floor

For a single analysis reused by both outputs, a proposed zero-repair stage
design has one interpretation/analysis proposal, one workbook plan, one deck
plan, one narrative review, and one rendered montage review: **5 model
requests** for XA01. PP01 needs one workbook binding/analysis proposal, one
deck patch plan, one narrative review, and one montage review: **4 requests**.
If visual review must run once per slide, replace one montage call with four
or six per-slide calls: floors become **8** and **9**. These counts include the
first chat-loop planning request and assume deterministic source capture,
arithmetic, execution, and final receipt; they exclude repairs, syntax retries,
or an extra model-based final review. Before using the proposed 12/18-call
targets, measure the actual implementation's stage calls and adjust this floor.

## Offline mechanism status

- Existing guards exercise stale briefs, incorrect ranking objections, ISO
  date suffixes, and native folio identity, though they still test filters.
- New synthetic endpoint coverage records repeated evidence in successive
  request payloads, exposing context growth without paid inference.
- New fake Excel coverage injects Formula and FormulaLocal rejection and
  checks that the grid and marked sheet were already written. This documents
  current write-then-validate ordering; phase 5 must replace that behavior.
- A focused offline baseline test supplies `CHART_HIGHLIGHT_BOUNDS` and checks
  that the current repair directive requests a complete slide. The current
  `RepairOwnedGroupAsync` calls `RepairSlideContentAsync` for visual findings;
  phase 3 must replace this geometry route with renderer-owned repair.

The Python usage-summary test passes locally. C# compilation and guardrail
execution are pending because this workstation has no current MSBuild or .NET
SDK; the legacy Framework MSBuild cannot parse PackageReference. Phase 0 is
**not complete**. Do not begin phase 1 until runtime-capability evidence and
the CI gate are reported.
