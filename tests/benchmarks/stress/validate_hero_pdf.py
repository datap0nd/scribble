"""Validate and contact-sheet every rendered page of the 130-page fixture."""
from __future__ import annotations

import argparse
from pathlib import Path
from PIL import Image, ImageChops, ImageDraw
from pypdf import PdfReader


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--pdf", type=Path, required=True)
    parser.add_argument("--renders", type=Path, required=True)
    args = parser.parse_args()
    pages = sorted(args.renders.glob("page-*.png"))
    reader = PdfReader(str(args.pdf))
    if len(reader.pages) != 130 or len(pages) != 130:
        raise RuntimeError(f"Expected 130 PDF and rendered pages; found {len(reader.pages)} and {len(pages)}")
    text = "\n".join(page.extract_text() or "" for page in reader.pages)
    for required in ("18.4 million", "96.7 percent", "R-317", "2026-11-15", "ORION-PAGE-130"):
        if required not in text:
            raise RuntimeError(f"Missing extracted fact: {required}")
    sheets = args.renders / "contact-sheets"
    sheets.mkdir(exist_ok=True)
    for batch in range(13):
        selected = pages[batch * 10:(batch + 1) * 10]
        canvas = Image.new("RGB", (1500, 1240), "#d9dde5")
        draw = ImageDraw.Draw(canvas)
        for offset, page_path in enumerate(selected):
            with Image.open(page_path).convert("RGB") as page:
                if page.size[0] < 500 or page.size[1] < 700:
                    raise RuntimeError(f"Unexpected render dimensions: {page_path} {page.size}")
                difference = ImageChops.difference(page, Image.new("RGB", page.size, "white"))
                bounds = difference.getbbox()
                if bounds is None or bounds[0] < 8 or bounds[1] < 8 or bounds[2] > page.width - 8 or bounds[3] > page.height - 8:
                    raise RuntimeError(f"Blank or edge-clipped rendered page: {page_path} {bounds}")
                page.thumbnail((280, 540))
                x = 15 + (offset % 5) * 297
                y = 50 + (offset // 5) * 590
                canvas.paste(page, (x, y))
                draw.text((x, 18 + (offset // 5) * 590), f"Page {batch * 10 + offset + 1}", fill="#111827")
        canvas.save(sheets / f"pages-{batch*10+1:03d}-{batch*10+len(selected):03d}.png")
    print(f"Validated {len(pages)} rendered pages and wrote 13 contact sheets.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
