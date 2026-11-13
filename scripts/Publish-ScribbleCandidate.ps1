param(
    [Parameter(Mandatory=$true)][long]$BuildRunId,
    [string]$EvidencePath
)
$ErrorActionPreference = 'Stop'

# Kept as a fail-closed entry point for old operator commands. Continuous
# updates are published only by the build workflow's owner-dispatched
# publish_continuous run on codex/development; Latest stays at stable 2.0.91.
throw 'This legacy publisher is retired. No release was changed. Publish a continuous update by running the "Build Windows installer" workflow on codex/development with publish_continuous ticked. See docs/release-channels.md.'
