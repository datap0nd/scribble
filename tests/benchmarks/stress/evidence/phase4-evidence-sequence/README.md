# Phase 4 evidence sequence native render

- Code commit: `c8cbea6` (PR #40), stacked on PR #39.
- Native harness: `GuardrailTests.exe --native-phase4-reference` from Windows CI run `36102967482`'s `Phase2NativeHarness` artifact.
- Fixture SHA-256: `98cd27392938ddaa1ab714007808eb16b7a9dd9922c5e1b2cc19062900c82484`.
- Harness assembly SHA-256: `dbd7ec9216a9b91ccbdaea403cfcbe360cf547f73e522a8d3610af7b3120be65`.
- Native result: 18 reference slides rendered with `structural_passed=true`; `visual_approved=false`, reviewer unidentified. PowerPoint exported editable slides to PDF.
- `before-normal.png` is the corresponding render from PR #39's code commit `193c4b8`.
- `evidence-normal.png` and `evidence-long.png` are pages 4–5 of the second native deck at 120 DPI. Both use the new sequence. `evidence-max.png` is page 6 and retains the four-card recipe.
- The bounded-length follow-up at `c8cbea6` produced byte-identical PNGs to the first native run, showing the normal and long fixtures still take this layout.

| Image | SHA-256 |
| --- | --- |
| `before-normal.png` | `05edc447d94fe5652bb63c196818acf8b6406640d5a88adb40ad8db9326966fb` |
| `evidence-normal.png` | `e5a6ab7dd3a0889379ad28703acddcb79eb97f3bbfc478e2d9eecd7a2d659cf7` |
| `evidence-long.png` | `1492c32b6a89a9bb650f08ff9b09018863676a3ffa8115b78715904d583507ec` |
| `evidence-max.png` | `c16e65d880ee66a93a9317f1d7f5680978413e96ab785f78276323a0f156e401` |

This is component evidence, not Phase 4 visual approval or a paid-model acceptance run.
