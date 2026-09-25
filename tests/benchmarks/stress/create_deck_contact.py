"""Create a labeled contact sheet from six rendered PowerPoint pages."""

import sys
from pathlib import Path

from PIL import Image, ImageDraw


def main(folder):
    pages = [folder / f"page-{index}.png" for index in range(1, 7)]
    if any(not page.is_file() for page in pages):
        raise SystemExit("Expected page-1.png through page-6.png")
    width, height, gap, label = 480, 270, 18, 28
    sheet = Image.new("RGB", (3 * width + 4 * gap,
                              2 * (height + label) + 3 * gap), "#e8edf2")
    draw = ImageDraw.Draw(sheet)
    for index, path in enumerate(pages):
        with Image.open(path) as image:
            tile = image.convert("RGB").resize((width, height),
                                                 Image.Resampling.LANCZOS)
        x = gap + (index % 3) * (width + gap)
        y = gap + (index // 3) * (height + label + gap)
        sheet.paste(tile, (x, y))
        draw.rectangle((x, y, x + width - 1, y + height - 1),
                       outline="#536579", width=1)
        draw.text((x + 5, y + height + 6), f"Slide {index + 1}",
                  fill="#16354f")
    sheet.save(folder / "contact.png", optimize=True)


if __name__ == "__main__":
    main(Path(sys.argv[1]).resolve())
