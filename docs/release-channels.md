# Scribble release channels

The owner froze the public release at **2.0.91** on 2026-09-14. Keep one
repository so fixes and history remain shared, with separate source branches
and a deliberate public-release boundary.

| Purpose | Branch or release | Policy |
| --- | --- | --- |
| Ongoing development | `codex/development` | Feature PRs target this branch. CI artifacts are for development testing. |
| Original stable source | `codex/stable-2.0.91` | Frozen at `f971839c873cb5fa3347b407a7dc37e49d1894f8`. Do not advance it. |
| Integration history | `main` | Preserved, including the release-freeze change. Builds cannot publish releases. |
| Public Latest | `stable-2.0.91` release/tag | Original 2.0.91 installer, not a rebuild. |
| Newer updater compatibility | `continuous` release/tag | Same original installer plus a matching `candidate.json`. |

The original artifact is from [successful build 91](https://github.com/datap0nd/scribble/actions/runs/33855311737).
Its Windows file version is `2.0.91.0`, browser extension version is `1.5.2`,
and installer SHA-256 is
`4c655e006c79b3d1ddc1995d32530e4636e37ce7ff4696c1fef1f84fbfe239fa`.
The manifest was added during restoration for compatibility with newer updaters;
it was not an artifact of build 91. It identifies the unchanged original bytes.

## User downloads

- [Latest download](https://github.com/datap0nd/scribble/releases/latest/download/ScribbleSetup.exe): used by 2.0.91's Update button and existing installation links.
- [Continuous download](https://github.com/datap0nd/scribble/releases/download/continuous/ScribbleSetup.exe): used by newer updater helpers.
- [Versioned stable download](https://github.com/datap0nd/scribble/releases/download/stable-2.0.91/ScribbleSetup.exe): fixed recovery link.

All three must return the hash and version above. A person already running
2.0.91 can reinstall the same stable release but receives no development build.
Newer updater helpers reject downgrades, so an already-upgraded installation
does not silently roll back. An intentional rollback requires saving work,
closing Office and running the versioned stable installer. No compatibility
claim for rolling newer settings back to 2.0.91 is made here.

## Development and future releases

Start new feature branches from current `origin/codex/development`. Merge tested
PRs back there. The repository default branch should be `codex/development` so
new checkouts and PRs start in the correct place. Download development installers
from Actions artifacts when explicitly testing development. Artifact version
numbers may increase; that does not change the stable public release.

The build workflow has no release job and uses `contents: read`. The old
`Publish-ScribbleCandidate.ps1` entry point always throws before contacting
GitHub, including when passed successful runs or acceptance evidence.

Only a new explicit owner release request can lift the freeze. Prepare and
review a separate promotion change with the exact approved version, successful
build, original artifact hash, recovery copy and matching manifest. Do not
rebuild the old source and label that artifact 2.0.91: the workflow stamps its
new run number into the file. Do not mark any development release Latest or
replace the continuous assets during ordinary development.

See the [freeze test plan](testing/releases/2026-09-14-stable-freeze/test-plan.md)
and [test report](testing/releases/2026-09-14-stable-freeze/test-report.md).
