# Scribble (scribble)

Windows Office COM add-in suite (.NET Framework 4.8, C# 7.3, classic
csproj): one assembly hosts four add-ins — Outlook (`AddIn.cs`), Excel
(`ExcelAddIn.cs`), PowerPoint (`PowerPointAddIn.cs`), and Word
(`WordAddIn.cs`) — sharing the chat
stack, Gemini/local-model client, MCP client, and guardrails. It cannot be
built or run on Linux — the Windows CI workflow
(`.github/workflows/build.yml`) is the compile/test gate.

## Git workflow

- **This is the only actively developed repository** — its predecessor
  `outlook-local-ai-chat` (MetoAI) is frozen; do not port changes there
  unless explicitly asked.
- Public Scribble is **frozen at 2.0.91**. Both GitHub Latest and the
  `continuous` compatibility download must serve the original stable installer.
  Never publish a development build or change Latest without a new explicit
  owner request. See [release channels](docs/release-channels.md).
- Develop on `codex/development`, using `codex/` feature branches and tested
  PRs targeting `codex/development`. `codex/stable-2.0.91` preserves the exact
  stable source; do not advance or rewrite it. `main` retains the integration
  history and release-freeze change; it is not a public update channel.
- CI on `main`, `codex/development`, and PRs builds and tests installers,
  then uploads development artifacts only. The workflow has read-only
  repository permissions and no public-release job. The legacy publication
  script deliberately refuses every promotion while the freeze is in effect.

## Code conventions

- C# 7.3 only (no target-typed `new`, ranges, or switch expressions).
- String concatenation over interpolation; match the existing wrapping style.
- New source files must be added to `Scribble.csproj` (classic
  csproj — no globbing).
- Security boundaries are load-bearing: the static scan asserts exact strings
  in several files (tool names, draft authorization, working-set caps, the
  Scribble Draft sheet name, the [Scribble draft] slide marker, and the MCP
  namespace). Check `scripts/Test-Guardrails.ps1` before renaming or
  rewording anything it references.
- Hard capability rules across every host: the model can never send email,
  never save/delete/print/close documents, and every write surface is a
  clearly marked draft gated by a one-shot, prompt-authorized permission.
- Guardrail tests use only public APIs (no InternalsVisibleTo) and a
  hand-rolled runner in `tests/GuardrailTests/Program.cs` — register new
  tests in `Main`.
