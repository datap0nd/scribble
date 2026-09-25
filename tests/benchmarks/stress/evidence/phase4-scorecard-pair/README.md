# Phase 4 paired scorecard native render

- Code commit: `649a8e2` (PR #39).
- Native harness: `GuardrailTests.exe --native-phase4-reference` from Windows CI run `36100320270`'s `Phase2NativeHarness` artifact.
- Fixture SHA-256: `98cd27392938ddaa1ab714007808eb16b7a9dd9922c5e1b2cc19062900c82484`.
- Harness assembly SHA-256: `b16a3a0c2a1fc4a536897daeaba4f3e3597f30559e025b85a77af23eee47242b`.
- Native result: 18 reference slides rendered, `structural_passed=true`; `visual_approved=false`, with no identified reviewer yet. The deck is editable and its PDF was exported by installed PowerPoint.
- `before-scorecard-normal.png` is the native render from PR #38, assembly SHA-256 `5640b10c020fa0b9253363b6e1d60abceec7da763daf6daff417061d7c4ce784`.
- `scorecard-normal.png`, `scorecard-long.png`, and `scorecard-max.png` are pages 1–3 of the first PowerPoint deck rendered from PR #39 at 120 DPI. Only the two-measure normal fixture takes the new composition; the denser fixtures are shown to verify they remain readable.

| Image | SHA-256 |
| --- | --- |
| `before-scorecard-normal.png` | `d8c01262b587ca148ef5dcdeb3fae13566ecf5e6421dd67d71388ec432ddcd2e` |
| `scorecard-normal.png` | `4539f52115f3081f4c688840e993546475b60b67b3088a7481b3b2d11a666fb7` |
| `scorecard-long.png` | `44ccba078045b3c4bbacee72323993039e297cef858ba5f1dbc356d8345eb1e0` |
| `scorecard-max.png` | `a316b02c338ed047ec586b82cb5caffb5a2fc44eed43b60dd0d6121f26c90c9a` |

This is component evidence, not Phase 4 visual approval or a paid-model acceptance run.
