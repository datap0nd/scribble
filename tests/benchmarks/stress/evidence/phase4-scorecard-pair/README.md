# Phase 4 paired scorecard native render

- Code commit: `193c4b8` (PR #39).
- Native harness: `GuardrailTests.exe --native-phase4-reference` from Windows CI run `36100834634`'s `Phase2NativeHarness` artifact.
- Fixture SHA-256: `98cd27392938ddaa1ab714007808eb16b7a9dd9922c5e1b2cc19062900c82484`.
- Harness assembly SHA-256: `34931d05780e7fe96587c9859f74c6ea4d61007ac9d5c08e41bbf7a14d4dd2f4`.
- Native result: 18 reference slides rendered, `structural_passed=true`; `visual_approved=false`, with no identified reviewer yet. The deck is editable and its PDF was exported by installed PowerPoint.
- `before-scorecard-normal.png` is the native render from PR #38, assembly SHA-256 `5640b10c020fa0b9253363b6e1d60abceec7da763daf6daff417061d7c4ce784`.
- `scorecard-normal.png`, `scorecard-long.png`, and `scorecard-max.png` are pages 1–3 of the first PowerPoint deck rendered from PR #39 at 120 DPI. All three take a source-preserving composition appropriate to their number of measures.

| Image | SHA-256 |
| --- | --- |
| `before-scorecard-normal.png` | `d8c01262b587ca148ef5dcdeb3fae13566ecf5e6421dd67d71388ec432ddcd2e` |
| `scorecard-normal.png` | `4539f52115f3081f4c688840e993546475b60b67b3088a7481b3b2d11a666fb7` |
| `scorecard-long.png` | `30d36e12e4ee543ebe8abb49baaacc12a24d3ff0fb830b1546d24e572e1b0c2` |
| `scorecard-max.png` | `608dd2e3d93a0aba20b73e47c084a871324db5d54ae277560b71e4610c78bd38` |

This is component evidence, not Phase 4 visual approval or a paid-model acceptance run.
