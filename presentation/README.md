# scribble

A six-slide deck introducing **AI365** to a non-technical audience: what it
is, why it is private, why it costs nothing, and why its answers can be
trusted.

The output is `../scribble.pptx`.

## Rebuilding

```bash
python3 make_pixel_logo.py   # renders assets/pal_*.png (needs Pillow)
npm install pptxgenjs
node build_scribble.js       # writes ../scribble.pptx
```

The slides are generated, not hand-drawn, so wording changes are an edit to
`build_scribble.js` and a rebuild rather than a pass through PowerPoint.

## The logo

`assets/pal_idle.png` and `assets/pal_wave.png` are the AI365 "pixel pal" -
the little robot that types away in the chat sidebar while the model thinks.
`make_pixel_logo.py` carries its sprite rows and palette verbatim from the
`palFrames` and `palColors` literals in
`src/OutlookLocalAIChat/UI/ChatPaneWeb.html` and renders them at 48x, so the
deck's mark and the product's mark are the same drawing.

Nothing keeps the two in sync automatically. If the sprite in the chat pane
ever changes, copy the new rows across and re-run the script.
