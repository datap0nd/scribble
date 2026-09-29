"""Create a labeled contact sheet from consecutive rendered PowerPoint pages."""

import math
import re
import sys
from pathlib import Path

from PIL import Image, ImageDraw, ImageOps


def main(folder):
    numbered = []
    for path in folder.glob("page-*.png"):
        match = re.fullmatch(r"page-(\d+)\.png", path.name)
        if match:
            numbered.append((int(match.group(1)), path))
    numbered.sort()
    if not numbered or [number for number, _ in numbered] != list(
        range(1, len(numbered) + 1)
    ):
        raise SystemExit("Expected consecutive page-1.png through page-N.png")
    pages = [path for _, path in numbered]
    width, height, gap, label = 480, 270, 18, 28
    columns = 2 if len(pages) <= 4 else 3
    rows = math.ceil(len(pages) / columns)
    sheet = Image.new("RGB", (columns * width + (columns + 1) * gap,
                              rows * (height + label) + (rows + 1) * gap),
                      "#e8edf2")
    draw = ImageDraw.Draw(sheet)
    for index, path in enumerate(pages):
        with Image.open(path) as image:
            tile = ImageOps.contain(image.convert("RGB"), (width, height),
                                    Image.Resampling.LANCZOS)
        x = gap + (index % columns) * (width + gap)
        y = gap + (index // columns) * (height + label + gap)
        sheet.paste(tile, (x + (width - tile.width) // 2,
                           y + (height - tile.height) // 2))
        draw.rectangle((x, y, x + width - 1, y + height - 1),
                       outline="#536579", width=1)
        draw.text((x + 5, y + height + 6), f"Slide {index + 1}",
                  fill="#16354f")
    sheet.save(folder / "contact.png", optimize=True)


if __name__ == "__main__":
    main(Path(sys.argv[1]).resolve())
