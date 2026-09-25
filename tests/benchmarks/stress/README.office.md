# Synthetic Office corpus

`generate_office.py` computes deterministic fictional records and independent
oracles. `author_office.mjs` uses the bundled Artifact Tool to write 20 editable
workbooks and 30 editable presentation sources. Each workbook has 144 operating
records across January–June 2026, four groups and six items; four workbooks add
one exact duplicate, and five have one unknown June amount. Domains cover sales,
inventory, budgets, returns, workforce and projects. Every source contains native
formulas and a native historical chart.

The 30 decks contain 6–12 slides, native tables and charts with embedded workbook
data. Ten sources deliberately contain three documented defects: an off-canvas
chart, overflowing commentary and a wrong table-header accent. Clean references
are stored separately for review. The theme uses the repository's Samsung MD 2.0
palette and geometry; Arial is its documented fallback font. The reference is a
controlled fixture, not a claim of human-approved Samsung fidelity.

## Files and boundaries

The ignored output directory defaults to `tests/benchmarks/generated/stress-corpus`.

| Location | Purpose | Model input |
| --- | --- | --- |
| `inputs/excel/WB01.xlsx` … `WB20.xlsx` | Operational source records | Yes |
| `inputs/powerpoint/PPT01.pptx` … `PPT30.pptx` | Operational source decks | Yes |
| `inputs/office_sources.json` | Safe source IDs, paths, domains and column names | Yes |
| `evaluator-only/office_catalog.json` | Independent facts, missing/duplicate contracts, formula probes | **No** |
| `evaluator-only/SamsungMD2.theme.json` | Theme grading contract in native PowerPoint points | **No** |
| `evaluator-only/powerpoint-references/` | Clean presentation references | **No** |
| `evaluator-only/previews/` | Authoring PNGs and private readback | **No** |
| `evaluator-only/native-validation/` | Native Office measurements and renders | **No** |
| `.build/` | Private authoring payload, candidates and validation receipts | **No** |

Deck `PPTnn` links workbook `WB((nn-1) % 20 + 1)`. All rows and amounts are invented;
no mailbox, company files, web sources or real employee information are used.

## Generate

Read the installed spreadsheet and presentation skills and record their artifact
operation markers before first authoring. Obtain the runtime paths from the
Codex workspace-dependency tool. Supply those paths explicitly:

```powershell
& $RuntimePython tests/benchmarks/stress/generate_office.py `
  --runtime-node $RuntimeNode --node-modules $RuntimeNodeModules `
  --presentations-skill $PresentationsSkill --verify
```

`--data-only` writes metadata without authoring Office files. `--only WB01,PPT01`
regenerates a bounded subset. Full generation omits `--only`; full verification
requires all 50 source packages. Source records, facts, formulas, chart values
and layouts are reproducible. OOXML package IDs and private receipt timestamps
may vary between builds; deterministic `source_sha256` values fingerprint the
semantic input payload, while verification records each actual file hash.

## Native verification

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File `
  tests/benchmarks/stress/Validate-OfficeCorpus.ps1
```

The script creates an isolated Excel instance and private workbook copies. It
recalculates every source formula, checks independent expected values, retains
unknowns, verifies chart values and categories, and changes one source input to
verify dependent formulas. It opens presentations read-only, renders every slide,
checks native chart data, table counts, theme fonts/colors, canvas bounds and text
fit, and verifies the ten repair sources reproduce only their declared defects.
It closes only objects it opened, does not activate chart data, update links or
refresh external data, and verifies source file hashes remain unchanged.

These checks prove source mechanics and measured structure. They do **not** prove
Qwen answers, independent visual quality, or a completed 200-case model run.
Inventory uses a latest-month snapshot; summing closing inventory across months
is invalid. Rates use aggregate numerators/denominators, not averages of row rates.
For incomplete facts, the oracle distinguishes a known subtotal from a full total.
