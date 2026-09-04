from __future__ import annotations

import base64
import math
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parent
ASSET_DIR = ROOT / "assets"
SHOT_DIR = ASSET_DIR / "setup-guide"
RAW_DIR = SHOT_DIR / "raw"
OUTPUT = ROOT / "setup-guide.html"
FONT_DIR = Path(r"C:\Windows\Fonts")

BLUE = "#376fe8"
ORANGE = "#ff8a3d"
WHITE = "#ffffff"


def font(filename: str, size: int) -> ImageFont.FreeTypeFont:
    return ImageFont.truetype(str(FONT_DIR / filename), size)


def arrow(
    draw: ImageDraw.ImageDraw,
    points: list[tuple[int, int]],
    color: str = ORANGE,
    width: int = 6,
) -> None:
    draw.line(points, fill=color, width=width, joint="curve")
    start = points[-2]
    end = points[-1]
    angle = math.atan2(end[1] - start[1], end[0] - start[0])
    size = 18
    spread = 0.55
    left = (
        end[0] - size * math.cos(angle - spread),
        end[1] - size * math.sin(angle - spread),
    )
    right = (
        end[0] - size * math.cos(angle + spread),
        end[1] - size * math.sin(angle + spread),
    )
    draw.polygon([end, left, right], fill=color)


def badge(
    draw: ImageDraw.ImageDraw,
    center: tuple[int, int],
    number: str,
    radius: int = 23,
) -> None:
    x, y = center
    draw.ellipse(
        (x - radius, y - radius, x + radius, y + radius),
        fill=BLUE,
        outline=WHITE,
        width=4,
    )
    number_font = font("segoeuib.ttf", 23)
    box = draw.textbbox((0, 0), number, font=number_font)
    width = box[2] - box[0]
    height = box[3] - box[1]
    draw.text(
        (x - width / 2, y - height / 2 - 3),
        number,
        font=number_font,
        fill=WHITE,
    )


def save_installer() -> Path:
    source = Image.open(RAW_DIR / "scribble-installer.jpg").convert("RGB")
    output = SHOT_DIR / "01-choose-apps.png"
    source.save(output, optimize=True)
    return output


def save_chrome() -> Path:
    image = Image.open(RAW_DIR / "chrome-extensions.jpg").convert("RGB")
    draw = ImageDraw.Draw(image)

    arrow(draw, [(1150, 183), (1200, 150), (1250, 115)])
    badge(draw, (1150, 183), "1")

    arrow(draw, [(205, 225), (150, 200), (90, 171)])
    badge(draw, (205, 225), "2")

    output = SHOT_DIR / "02-enable-chrome.png"
    image.save(output, optimize=True)
    return output


def save_folder_picker() -> Path:
    image = Image.open(RAW_DIR / "chrome-folder-picker.jpg").convert("RGB")
    draw = ImageDraw.Draw(image)

    arrow(draw, [(82, 410), (125, 410), (170, 410)], width=5)
    badge(draw, (82, 410), "3", radius=21)
    draw.rounded_rectangle(
        (398, 425, 516, 464),
        radius=6,
        outline=ORANGE,
        width=4,
    )

    output = SHOT_DIR / "03-select-extension-folder.png"
    image.save(output, optimize=True)
    return output


def save_settings() -> Path:
    source = Image.open(RAW_DIR / "scribble-settings.jpg").convert("RGB")
    output = SHOT_DIR / "04-connect-model.png"
    source.save(output, optimize=True)
    return output


def data_uri(path: Path) -> str:
    encoded = base64.b64encode(path.read_bytes()).decode("ascii")
    media_type = "image/jpeg" if path.suffix.lower() in {".jpg", ".jpeg"} else "image/png"
    return f"data:{media_type};base64,{encoded}"


def build_html(shots: list[Path]) -> str:
    installer, chrome, picker, settings = [data_uri(path) for path in shots]
    return fr"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <meta name="color-scheme" content="light">
  <title>Scribble setup guide</title>
  <style>
    * {{ box-sizing: border-box; }}
    html, body {{ margin: 0; padding: 0; background: #edf2f8; color: #17233b; }}
    body {{ font-family: "Segoe UI", Arial, sans-serif; padding: 28px 14px 48px; }}
    .guide {{ width: 100%; max-width: 820px; margin: 0 auto; background: #fff; border: 1px solid #d9e1ec; border-radius: 20px; overflow: hidden; box-shadow: 0 16px 44px rgba(29,48,81,.12); }}
    .accent {{ height: 6px; background: #376fe8; }}
    header {{ padding: 44px 48px 36px; display: grid; grid-template-columns: 1fr 112px; gap: 28px; align-items: center; }}
    .eyebrow, .step-number {{ margin: 0 0 10px; font: 700 12px/18px "Courier New", monospace; letter-spacing: .8px; color: #16845b; }}
    h1 {{ margin: 0 0 14px; font-size: 42px; line-height: 1.08; letter-spacing: -1.2px; }}
    header p {{ margin: 0; font-size: 18px; line-height: 1.55; color: #536177; }}
    .pal {{ width: 112px; height: 106px; object-fit: contain; image-rendering: pixelated; padding: 10px; border: 1px solid #c9dbff; border-radius: 16px; background: #eff5ff; }}
    .before {{ margin: 0 48px 12px; padding: 18px 20px; border: 1px solid #c8e7d8; border-radius: 12px; background: #f1faf6; color: #405269; line-height: 1.55; }}
    .before strong {{ color: #16845b; }}
    main {{ padding: 10px 48px 42px; }}
    .step {{ padding: 34px 0 38px; border-bottom: 1px solid #e1e7f0; }}
    .step:last-child {{ border-bottom: 0; padding-bottom: 12px; }}
    h2 {{ margin: 0 0 9px; font-size: 25px; line-height: 1.25; }}
    .copy {{ margin: 0 0 20px; color: #536177; font-size: 16px; line-height: 1.6; }}
    .shot {{ width: 100%; height: auto; display: block; border: 1px solid #d7e0ec; border-radius: 14px; background: #edf2f8; }}
    .optional {{ display: inline-block; margin-left: 8px; padding: 4px 8px; border-radius: 999px; background: #fff3dc; color: #9a5b16; font: 700 10px/14px "Courier New", monospace; vertical-align: 3px; }}
    .mini-steps {{ display: grid; grid-template-columns: repeat(3, 1fr); gap: 10px; margin: 0 0 18px; }}
    .mini-step {{ min-height: 74px; padding: 13px; border: 1px solid #d7e0ec; border-radius: 11px; color: #405269; line-height: 1.35; background: #f8faff; }}
    .mini-step b {{ display: inline-grid; place-items: center; width: 25px; height: 25px; margin-right: 7px; border-radius: 50%; background: #376fe8; color: #fff; }}
    .path-card {{ margin: 22px 0 18px; padding: 17px 18px; border-radius: 12px; background: #17233b; color: #fff; }}
    .path-card span {{ display: block; margin-bottom: 8px; color: #8ee0b9; font: 700 11px/16px "Courier New", monospace; letter-spacing: .7px; }}
    .path-card code {{ display: block; overflow-wrap: anywhere; font: 700 15px/22px "Courier New", monospace; color: #fff; }}
    .done {{ margin-top: 30px; padding: 24px; text-align: center; background: #17233b; border-radius: 14px; color: #fff; }}
    .done h2 {{ margin-bottom: 7px; }}
    .done p {{ margin: 0; color: #b9c8df; line-height: 1.5; }}
    footer {{ padding: 22px 30px; text-align: center; border-top: 1px solid #e1e7f0; background: #f8faff; font: 700 11px/18px "Courier New", monospace; }}
    footer a {{ color: #285fcf; }}
    @media (max-width: 620px) {{
      header {{ padding: 34px 24px 28px; grid-template-columns: 1fr; }}
      header .pal {{ grid-row: 1; }}
      h1 {{ font-size: 34px; }}
      .before {{ margin: 0 24px 8px; }}
      main {{ padding: 8px 24px 32px; }}
      .mini-steps {{ grid-template-columns: 1fr; }}
    }}
  </style>
</head>
<body>
  <div class="guide">
    <div class="accent"></div>
    <header>
      <div>
        <p class="eyebrow">SCRIBBLE / QUICK START</p>
        <h1>From download to first prompt in about 5 minutes.</h1>
        <p>Four clear stages. Real screens. No guesswork.</p>
      </div>
      <img class="pal" src="{data_uri(ASSET_DIR / 'pixel-pal.png')}" alt="Pixel Pal">
    </header>

    <div class="before"><strong>Before you start:</strong> close Outlook, Excel, PowerPoint, and Word. Have your AI endpoint URL and API key ready.</div>

    <main>
      <section class="step">
        <p class="step-number">01 / INSTALL</p>
        <h2>Download Scribble and choose your apps</h2>
        <p class="copy">Run <a href="https://github.com/datap0nd/scribble/releases/latest/download/ScribbleSetup.exe">ScribbleSetup.exe</a>. Choose <strong>Next</strong> on the destination screen. Keep all five components checked, or clear anything you do not need, then continue through Setup.</p>
        <img class="shot" src="{installer}" alt="Actual Scribble installer component-selection screen">
      </section>

      <section class="step">
        <p class="step-number">02 / CHROME <span class="optional">ONLY IF SELECTED</span></p>
        <h2>Approve the Chrome extension</h2>
        <div class="mini-steps">
          <div class="mini-step"><b>1</b>Turn on <strong>Developer mode</strong></div>
          <div class="mini-step"><b>2</b>Choose <strong>Load unpacked</strong></div>
          <div class="mini-step"><b>3</b>Paste the path and choose <strong>Select Folder</strong></div>
        </div>
        <img class="shot" src="{chrome}" alt="Actual Chrome Extensions page with numbered callouts for Developer mode and Load unpacked">
        <div class="path-card">
          <span>3 / COPY AND PASTE THIS PATH INTO THE FOLDER FIELD</span>
          <code>%LOCALAPPDATA%\Programs\Scribble\BrowserExtension</code>
        </div>
        <img class="shot" src="{picker}" alt="Actual Chrome folder picker with the Scribble extension path entered">
      </section>

      <section class="step">
        <p class="step-number">03 / CONNECT</p>
        <h2>Connect your AI model</h2>
        <p class="copy">Open Scribble from any supported app. In <strong>Settings</strong>, enter your endpoint and API key. Choose <strong>Connect &amp; load models</strong>, select a model, run <strong>Test selected model</strong>, then save.</p>
        <img class="shot" src="{settings}" alt="Actual Scribble Settings connection screen">
      </section>

      <section class="step">
        <p class="step-number">04 / START</p>
        <h2>Open Scribble and ask</h2>
        <p class="copy">In Outlook, open <strong>Scribble</strong> from its ribbon tab. In Excel, PowerPoint, and Word, use the <strong>Scribble</strong> button on Home. In Chrome, open Scribble from the extension side panel.</p>
        <div class="done">
          <h2>You are ready.</h2>
          <p>Add the context you want Scribble to use, then describe the outcome you need.</p>
        </div>
      </section>
    </main>

    <footer><a href="mailto:r.cunha@samsung.com?subject=Scribble%20setup%20help">NEED HELP? EMAIL THE DEVELOPER</a></footer>
  </div>
</body>
</html>
"""


def main() -> None:
    SHOT_DIR.mkdir(parents=True, exist_ok=True)
    required = [
        RAW_DIR / "scribble-installer.jpg",
        RAW_DIR / "chrome-extensions.jpg",
        RAW_DIR / "chrome-folder-picker.jpg",
        RAW_DIR / "scribble-settings.jpg",
    ]
    missing = [str(path) for path in required if not path.exists()]
    if missing:
        raise FileNotFoundError("Missing actual screenshot source: " + ", ".join(missing))

    shots = [
        save_installer(),
        save_chrome(),
        save_folder_picker(),
        save_settings(),
    ]
    OUTPUT.write_text(build_html(shots), encoding="utf-8")
    print(OUTPUT)
    for shot in shots:
        print(shot)


if __name__ == "__main__":
    main()
