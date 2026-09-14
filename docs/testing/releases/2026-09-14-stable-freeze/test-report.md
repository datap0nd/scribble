# Test report: public stable freeze

## Revision and evidence cutoff

- Date: 2026-09-14; initial evidence cutoff: 18:50:28 UTC.
- Source base: `28cb2d5071badf34bd3de03cbb467ba02394f9a2` plus this PR's
  release-freeze changes. This report is committed with those tested changes;
  the PR records its final head SHA and subsequent CI/merge evidence.
- Stable source: `f971839c873cb5fa3347b407a7dc37e49d1894f8`.
- Original installer: [successful build 91](https://github.com/datap0nd/scribble/actions/runs/33855311737),
  artifact `ScribbleSetup`, artifact ID `9929959164`.
- Environment: Windows, PowerShell 7, GitHub CLI 2.96.0. Scribble's focused
  PowerShell release checks were used; this is the Scribble .NET repository.

## Completed evidence

| Case | Result | Actual evidence |
| --- | --- | --- |
| FREEZE-01 | PASS | Original installer downloaded with `gh run download 33855311737 --repo datap0nd/scribble --name ScribbleSetup --dir artifacts/stable-2.0.91`; embedded version 2.0.91.0, 2,789,039 bytes, expected SHA-256. |
| FREEZE-02 | PASS | Direct `Invoke-WebRequest` downloads from Latest, continuous and stable-2.0.91 all returned the original hash/version at 18:50:26-27 UTC. `Invoke-RestMethod` verified the continuous manifest. GitHub Latest API returned stable-2.0.91, not a prerelease. |
| FREEZE-03 | PASS | `./tests/ReleaseGate/Test-PublicReleaseFreeze.ps1`: two attempted promotions refused before GitHub access; workflow publication forbidden and development trigger present. Synthetic CLI fixture only. |
| FREEZE-04 syntax | PASS | PowerShell Parser::ParseFile on scripts/Publish-ScribbleCandidate.ps1 and tests/ReleaseGate/Test-PublicReleaseFreeze.ps1: zero syntax errors. `git diff --check`: no whitespace errors. Git emitted normal LF-to-CRLF normalization warnings. |

Installer SHA-256:
`4c655e006c79b3d1ddc1995d32530e4636e37ce7ff4696c1fef1f84fbfe239fa`.
The compatibility manifest was created during restoration from verified
artifact/source identity; the installer itself was not modified or rebuilt.
Local download evidence: `artifacts/stable-2.0.91/verified/download-results.json`.
Previous public 2.0.187 installer and manifest were retained locally before
replacement. Original stable source and current development were preserved as
separate remote branches without rewriting source history.

## Later evidence

At this report's cutoff the final diff/link review, PR CI, merge, development
branch fast-forward/default-branch switch and post-CI download verification
are still pending. Their actual outcomes and exact revisions will be recorded
in the PR before delivery. Do not interpret the artifact's original successful
build as CI for this release-policy change.

## Limitations

Public routing is verified by downloading bytes, not by installing them.
Existing newer updater helpers reject a lower version and do not automatically
downgrade already-upgraded installations. No application code or UI changes
are included. Development CI artifacts can have higher file versions while
the public installer remains at 2.0.91.
