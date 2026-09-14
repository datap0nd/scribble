param(
    [Parameter(Mandatory=$true)][long]$BuildRunId,
    [string]$EvidencePath
)
$ErrorActionPreference = 'Stop'

# Kept as a fail-closed entry point for old operator commands. Development CI
# uploads artifacts only; neither a passing build nor acceptance evidence is
# authorization to replace the owner's frozen public release.
throw 'Public updates are frozen at Scribble 2.0.91. No release was changed. See docs/release-channels.md; a new public version requires explicit owner approval and a reviewed promotion change.'
