# Test plan: public stable freeze

## Scope and prerequisites

Verify release routing, preserved branches and refusal to promote development.
Use Windows PowerShell/PowerShell 7, Git and GitHub CLI in the Scribble checkout.
No application UI or user journey is changed. The original stable installer must
come from build `33855311737`, source
`f971839c873cb5fa3347b407a7dc37e49d1894f8`; do not rebuild it.

## Cases

1. **FREEZE-01: exact artifact.** Download `ScribbleSetup` using
   `gh run download 33855311737 --repo datap0nd/scribble --name ScribbleSetup --dir <empty-folder>`.
   Read `(Get-Item <installer>).VersionInfo` and
   `Get-FileHash <installer> -Algorithm SHA256`. Expect file version
   `2.0.91.0` and SHA-256
   `4c655e006c79b3d1ddc1995d32530e4636e37ce7ff4696c1fef1f84fbfe239fa`.
   Record source run, version, hash and timestamp.
2. **FREEZE-02: all public paths.** Download each installer URL in
   [release channels](../../../release-channels.md) with `Invoke-WebRequest
   -Uri <url> -OutFile <distinct-output.exe>`. Repeat the version/hash checks
   above for Latest, continuous and the versioned stable URL. Fetch
   `https://github.com/datap0nd/scribble/releases/download/continuous/candidate.json`
   with `Invoke-RestMethod`; require the same hash/version and stable commit.
   `gh api repos/datap0nd/scribble/releases/latest --jq .tag_name` must return
   `stable-2.0.91`. Record individual URL results and UTC timestamps.
3. **FREEZE-03: failure and bypass prevention.** Run
   `./tests/ReleaseGate/Test-PublicReleaseFreeze.ps1`. Its synthetic CLI rejects
   any network access. Both the original stable run and newer development run,
   even with an evidence argument, must throw the explicit freeze message
   before contacting GitHub. The build workflow must have read-only repository
   permissions, build `codex/development`, and contain no publication job.
4. **FREEZE-04: syntax and review.** Parse the publisher and freeze test with
   `[System.Management.Automation.Language.Parser]::ParseFile`; require zero
   parse errors. Run `git diff --check`. Review the actual diff once, including
   old and new updater download URLs and the newer helper's downgrade refusal.
   Follow all relative links added to documentation. No updater code changes
   or new installer build may masquerade as stable 2.0.91.
5. **FREEZE-05: final CI and branch separation.** Wait for all required checks
   from `.github/workflows/build.yml` on the final PR head. Record run URL and
   exact head SHA before a head-pinned merge to `main`. Fast-forward
   `codex/development` to include the tested freeze, set it as the default
   branch, and verify stable branch/tag still resolve to the original stable
   SHA. Confirm pre-split `main` remains an ancestor of development. After
   development CI completes, recheck public metadata and downloads to prove
   the build did not change the public release.

## Recovery and cleanup

If a public asset mismatches its manifest, stop promotion and restore the exact
verified 2.0.91 installer and manifest together, then repeat FREEZE-02. Preserve
the previous public installer/manifest before replacement. A failed check is
not permission to publish a newer candidate. Newer installed updaters refuse
downgrades; this task does not claim automatic rollback of existing installs.

Retain release artifacts as recovery copies and local hash evidence under
ignored `artifacts/stable-2.0.91/`; the previous release is under ignored
`artifacts/public-before-stable-freeze/`. Temporary downloads can be removed
after evidence is retained. Do not remove GitHub stable release assets or the
source branches during cleanup.
