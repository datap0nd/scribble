"""Prepare the existing Phase 4 reference renders for human page review."""

import hashlib
import json
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parent / "evidence" / "phase4-native-candidate"
REFERENCES = ROOT / "reference"
PACKET = ROOT / "visual-review"
FILES = [REFERENCES / f"phase4-reference-{deck}-{page}.png"
         for deck in range(1, 4) for page in range(1, 7)]


def main():
    if any(not path.is_file() for path in FILES):
        raise SystemExit("All 18 full-size reference PNGs must exist")
    PACKET.mkdir(exist_ok=True)
    entries = []
    width, height, gap, label_height = 448, 252, 20, 32
    sheet = Image.new("RGB", (3 * width + 4 * gap,
                              6 * (height + label_height) + 7 * gap),
                      "#e8edf2")
    draw = ImageDraw.Draw(sheet)
    for number, path in enumerate(FILES):
        deck, page = divmod(number, 6)
        digest = hashlib.sha256(path.read_bytes()).hexdigest()
        with Image.open(path) as original:
            if original.size != (1334, 750):
                raise SystemExit(f"Unexpected render size: {path}")
            tile = original.convert("RGB").resize((width, height),
                                                    Image.Resampling.LANCZOS)
        column, row = number % 3, number // 3
        x = gap + column * (width + gap)
        y = gap + row * (height + label_height + gap)
        sheet.paste(tile, (x, y))
        draw.rectangle((x, y, x + width - 1, y + height - 1),
                       outline="#536579", width=1)
        draw.text((x + 4, y + height + 6),
                  f"{number + 1:02d}  Deck {deck + 1}, page {page + 1}",
                  fill="#16354f")
        entries.append({"page_id": f"D{deck + 1}P{page + 1}",
                        "image": "../reference/" + path.name,
                        "sha256": digest,
                        "verdict": "pending"})
    sheet.save(PACKET / "contact.png", optimize=True)
    (PACKET / "verdicts.json").write_text(
        json.dumps({"schema": 1, "reviewer": None,
                    "status": "awaiting_user_review", "pages": entries},
                   indent=2) + "\n", encoding="utf-8")
    print(PACKET / "contact.png")


if __name__ == "__main__":
    main()
