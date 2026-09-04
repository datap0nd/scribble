from __future__ import annotations

import base64
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parent
ASSET_DIR = ROOT / "assets"
SHOT_DIR = ASSET_DIR / "setup-guide"
OUTPUT = ROOT / "setup-guide.html"
FONT_DIR = Path(r"C:\Windows\Fonts")

W, H = 1200, 675
NAVY = "#17233b"
TEXT = "#536177"
MUTED = "#7b8798"
BLUE = "#376fe8"
BLUE_SOFT = "#eff5ff"
GREEN = "#16845b"
GREEN_SOFT = "#effaf5"
LINE = "#dce4ef"
PAGE = "#edf2f8"


def font(filename: str, size: int) -> ImageFont.FreeTypeFont:
    return ImageFont.truetype(str(FONT_DIR / filename), size)


def regular(size: int) -> ImageFont.FreeTypeFont:
    return font("segoeui.ttf", size)


def semibold(size: int) -> ImageFont.FreeTypeFont:
    return font("segoeuib.ttf", size)


def mono(size: int) -> ImageFont.FreeTypeFont:
    return font("courbd.ttf", size)


def canvas() -> tuple[Image.Image, ImageDraw.ImageDraw]:
    image = Image.new("RGB", (W, H), PAGE)
    return image, ImageDraw.Draw(image)


def rounded(draw: ImageDraw.ImageDraw, box, radius: int, fill, outline=None, width=1) -> None:
    draw.rounded_rectangle(box, radius=radius, fill=fill, outline=outline, width=width)


def line_text(draw: ImageDraw.ImageDraw, xy, value: str, size: int, fill=TEXT, bold=False) -> None:
    draw.text(xy, value, font=semibold(size) if bold else regular(size), fill=fill)


def center_text(draw: ImageDraw.ImageDraw, box, value: str, text_font, fill) -> None:
    bounds = draw.textbbox((0, 0), value, font=text_font)
    tw, th = bounds[2] - bounds[0], bounds[3] - bounds[1]
    x0, y0, x1, y1 = box
    draw.text((x0 + (x1 - x0 - tw) / 2, y0 + (y1 - y0 - th) / 2 - 1), value, font=text_font, fill=fill)


def window(draw: ImageDraw.ImageDraw, title: str) -> None:
    rounded(draw, (54, 42, 1146, 632), 22, "#ffffff", "#ced8e6", 2)
    draw.rounded_rectangle((54, 42, 1146, 96), radius=22, fill="#f8faff")
    draw.rectangle((54, 74, 1146, 96), fill="#f8faff")
    draw.line((54, 96, 1146, 96), fill=LINE, width=2)
    rounded(draw, (76, 61, 94, 79), 9, BLUE)
    line_text(draw, (108, 57), title, 22, NAVY, True)
    for index, color in enumerate(("#f2c94c", "#62c88a", "#ef7373")):
        rounded(draw, (1060 + index * 22, 62, 1073 + index * 22, 75), 7, color)


def app_icon(name: str, size: int = 46) -> Image.Image:
    return Image.open(ASSET_DIR / f"{name}.png").convert("RGBA").resize((size, size), Image.Resampling.LANCZOS)


def screenshot_installer() -> Path:
    image, draw = canvas()
    window(draw, "Scribble Setup")
    line_text(draw, (94, 128), "Choose where Scribble works", 34, NAVY, True)
    line_text(draw, (94, 176), "All apps are selected by default. Clear anything you do not need.", 19, TEXT)

    apps = [
        ("outlook", "Outlook", "Mailbox chat and email drafts"),
        ("excel", "Excel", "Workbook analysis, tables, and charts"),
        ("powerpoint", "PowerPoint", "Executive-ready draft slides"),
        ("word", "Word", "Structured draft documents"),
        ("chrome", "Chrome", "Web research in dedicated work tabs"),
    ]
    for index, (icon_name, label, description) in enumerate(apps):
        column = index % 2
        row = index // 2
        x = 94 + column * 500
        y = 228 + row * 104
        rounded(draw, (x, y, x + 468, y + 82), 12, "#f8faff", LINE)
        rounded(draw, (x + 18, y + 26, x + 46, y + 54), 6, BLUE)
        center_text(draw, (x + 18, y + 26, x + 46, y + 54), "✓", semibold(19), "#ffffff")
        icon = app_icon(icon_name, 45)
        image.paste(icon, (x + 60, y + 18), icon)
        line_text(draw, (x + 120, y + 15), label, 20, NAVY, True)
        line_text(draw, (x + 120, y + 45), description, 15, TEXT)

    rounded(draw, (910, 551, 1108, 606), 9, BLUE)
    center_text(draw, (910, 551, 1108, 606), "Install", semibold(19), "#ffffff")
    path = SHOT_DIR / "01-choose-apps.png"
    image.save(path, optimize=True)
    return path


def screenshot_chrome() -> Path:
    image, draw = canvas()
    window(draw, "Google Chrome")
    rounded(draw, (170, 56, 880, 84), 14, "#edf1f6")
    line_text(draw, (194, 59), "chrome://extensions", 15, "#667487")
    line_text(draw, (92, 126), "Extensions", 34, NAVY, True)
    line_text(draw, (888, 134), "Developer mode", 17, NAVY, True)
    rounded(draw, (1040, 129, 1098, 157), 14, BLUE)
    rounded(draw, (1071, 133, 1094, 153), 11, "#ffffff")

    for index, label in enumerate(("Load unpacked", "Pack extension", "Update")):
        x = 92 + index * 178
        rounded(draw, (x, 190, x + 160, 235), 8, "#ffffff", "#bbc9dc", 2)
        center_text(draw, (x, 190, x + 160, 235), label, semibold(15), "#35527f")

    rounded(draw, (92, 275, 1108, 494), 16, "#ffffff", LINE, 2)
    pal = Image.open(ASSET_DIR / "pixel-pal.png").convert("RGBA").resize((96, 90), Image.Resampling.NEAREST)
    image.paste(pal, (125, 314), pal)
    line_text(draw, (250, 305), "Scribble", 25, NAVY, True)
    line_text(draw, (250, 347), "Your AI companion across Chrome and Office", 17, TEXT)
    rounded(draw, (250, 393, 370, 430), 18, GREEN_SOFT, "#bee3d2")
    center_text(draw, (250, 393, 370, 430), "Enabled", semibold(15), GREEN)
    line_text(draw, (92, 530), "Choose Load unpacked, then select the folder opened by Scribble Setup.", 19, NAVY, True)
    line_text(draw, (92, 570), r"Folder: %LOCALAPPDATA%\Programs\Scribble\BrowserExtension", 16, TEXT)

    path = SHOT_DIR / "02-enable-chrome.png"
    image.save(path, optimize=True)
    return path


def screenshot_settings() -> Path:
    image, draw = canvas()
    window(draw, "Scribble Settings")

    tabs = ("Connection", "Topics", "Skills", "Writing soul", "Support")
    x = 86
    for index, label in enumerate(tabs):
        width = 145 if label == "Writing soul" else 124
        if index == 0:
            rounded(draw, (x, 111, x + width, 153), 8, BLUE_SOFT, "#c5d8ff")
            center_text(draw, (x, 111, x + width, 153), label, semibold(15), BLUE)
        else:
            center_text(draw, (x, 111, x + width, 153), label, regular(15), MUTED)
        x += width + 8

    fields = [
        ("Endpoint URL", "https://your-endpoint.example/v1"),
        ("API key", "••••••••••••••••••••••••"),
        ("Model", "Select a compatible model"),
    ]
    y = 180
    for label, value in fields:
        line_text(draw, (92, y), label, 15, NAVY, True)
        rounded(draw, (92, y + 26, 736, y + 72), 7, "#ffffff", "#bfcbdc", 2)
        line_text(draw, (110, y + 37), value, 16, "#667487")
        y += 98

    rounded(draw, (782, 206, 1108, 261), 9, BLUE)
    center_text(draw, (782, 206, 1108, 261), "Connect & load models", semibold(17), "#ffffff")
    rounded(draw, (782, 304, 1108, 359), 9, "#ffffff", "#9db1cc", 2)
    center_text(draw, (782, 304, 1108, 359), "Test selected model", semibold(17), "#35527f")
    rounded(draw, (782, 402, 1108, 477), 12, GREEN_SOFT, "#bee3d2")
    center_text(draw, (782, 410, 1108, 445), "Connection ready", semibold(18), GREEN)
    center_text(draw, (782, 445, 1108, 474), "Authentication and tool calling passed", regular(13), TEXT)
    rounded(draw, (962, 548, 1108, 604), 9, BLUE)
    center_text(draw, (962, 548, 1108, 604), "Save", semibold(18), "#ffffff")

    path = SHOT_DIR / "03-connect-model.png"
    image.save(path, optimize=True)
    return path


def data_uri(path: Path) -> str:
    encoded = base64.b64encode(path.read_bytes()).decode("ascii")
    return f"data:image/png;base64,{encoded}"


def build_html(shots: list[Path]) -> str:
    screenshot_one, screenshot_two, screenshot_three = [data_uri(path) for path in shots]
    return f"""<!doctype html>
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
        <p>Three short steps. One installer. No admin account required.</p>
      </div>
      <img class="pal" src="{data_uri(ASSET_DIR / 'pixel-pal.png')}" alt="Pixel Pal">
    </header>

    <div class="before"><strong>Before you start:</strong> close Outlook, Excel, PowerPoint, and Word. Have your AI endpoint URL and API key ready.</div>

    <main>
      <section class="step">
        <p class="step-number">01 / INSTALL</p>
        <h2>Download Scribble and choose your apps</h2>
        <p class="copy">Run <a href="https://github.com/datap0nd/scribble/releases/latest/download/ScribbleSetup.exe">ScribbleSetup.exe</a>. Keep all five apps selected, or clear anything you do not need, then choose <strong>Install</strong>.</p>
        <img class="shot" src="{screenshot_one}" alt="Scribble installer app selection screen">
      </section>

      <section class="step">
        <p class="step-number">02 / CHROME <span class="optional">ONLY IF SELECTED</span></p>
        <h2>Approve the Chrome extension</h2>
        <p class="copy">Leave <strong>Finish setting up Scribble in Google Chrome</strong> selected. On the Extensions page, turn on <strong>Developer mode</strong>, choose <strong>Load unpacked</strong>, and select the folder opened by Setup.</p>
        <img class="shot" src="{screenshot_two}" alt="Chrome Extensions page with Scribble enabled">
      </section>

      <section class="step">
        <p class="step-number">03 / CONNECT</p>
        <h2>Connect your AI model</h2>
        <p class="copy">Open Scribble from any supported app. In <strong>Settings</strong>, enter your endpoint and API key. Choose <strong>Connect &amp; load models</strong>, select a model, run <strong>Test selected model</strong>, then save.</p>
        <img class="shot" src="{screenshot_three}" alt="Scribble Settings connection screen">
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
    shots = [screenshot_installer(), screenshot_chrome(), screenshot_settings()]
    OUTPUT.write_text(build_html(shots), encoding="utf-8")
    print(OUTPUT)
    for shot in shots:
        print(shot)


if __name__ == "__main__":
    main()
