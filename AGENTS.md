# Scribble — agent instructions

Follow `CLAUDE.md` for repository, branch and code conventions. The only
current plan is `DELIVERY.md`. Earlier plan, phase and checkpoint documents
are history only.

## Standing permission: test-owned Office automation

The owner grants this permission for this repository, and it replaces the
per-task approval that `~/.codex/AGENTS.md` otherwise requires for Office
automation.

You may launch, drive through COM and then close **test-owned** Excel,
PowerPoint and Word processes to do any of the following:
- run `tests/GuardrailTests` native harness modes;
- run `tests/NativeAcceptance/Test-DeliveryCandidate.ps1`;
- run `tests/benchmarks/stress/Run-StressSuite.ps1`;
- install development builds of Scribble on this workstation for those runs.

Limits:
- Record the PIDs and start times of Office processes that exist before a run.
  Never automate, close or stop them, or open their documents.
- Stop only processes the run itself started, identified by PID and start
  time. Never stop Office by process name.
- Use only disposable copies of corpus inputs.
- This permission does not cover Computer Use or clicking through the Windows
  desktop UI. Those still need an explicit request.

## Paid model runs

- Stay inside the spending authorization in `DELIVERY.md` §2 and its daily
  cap.
- Check the budget before every batch.
- Never print, copy or rotate API keys.

## How to work

- Work on the first open stage in `DELIVERY.md` §3.
- End every session with a new row in `tests/benchmarks/stress/SCOREBOARD.md`
  (a real or native run). If you can't, stop and write a blocker of 10 lines
  or fewer in `DELIVERY.md` §5.
- Progress is measured only by the scoreboard.
- Do not:
  - write new plan, roadmap, checkpoint or phase documents;
  - add gates or acceptance criteria;
  - create new worktrees or long-lived branches;
  - make commits that only archive evidence or add tracing;
  - add code that names corpus cases, companies, fixture labels, or fixed
    slide counts or indices;
  - add a prompt prohibition to rescue one run;
  - weaken a safety check to obtain a pass.
- If the same first-failing stage comes back after two fixes, stop patching.
  Write down the mechanism-level hypothesis, then change the approach or
  escalate.
