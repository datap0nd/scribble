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
- This is a personal dev repository. **Always commit directly to `main` and
  push immediately after each change set.** The user pulls `main` on a work
  machine to test. Do not create side branches or pull requests unless
  explicitly asked.
- Every push to `main` triggers CI: MSBuild, guardrail tests
  (`tests/GuardrailTests`), the static capability scan
  (`scripts/Test-Guardrails.ps1`), and republishing the installer to the
  `continuous` release that the README download link points at.

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
