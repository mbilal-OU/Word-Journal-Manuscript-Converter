from __future__ import annotations

import argparse
from pathlib import Path

from PIL import Image

SOURCE_ICON = Path("src/word_journal_manuscript_converter/assets/app-icon.png")
OUTPUT_DIR = Path("build/release-icons")
WINDOWS_ICON = OUTPUT_DIR / "app-icon.ico"
MACOS_ICON = OUTPUT_DIR / "app-icon.icns"
WINDOWS_SIZES = [(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]


def _load_source() -> Image.Image:
    if not SOURCE_ICON.exists():
        raise SystemExit(f"Missing source icon: {SOURCE_ICON}")
    image = Image.open(SOURCE_ICON).convert("RGBA")
    if image.width != image.height:
        raise SystemExit(f"Source icon must be square, got {image.width}x{image.height}")
    if min(image.size) < 256:
        raise SystemExit(f"Source icon is too small for release packaging: {image.size}")
    return image


def build_windows_icon(image: Image.Image) -> Path:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    image.save(WINDOWS_ICON, format="ICO", sizes=WINDOWS_SIZES)
    with Image.open(WINDOWS_ICON) as check:
        if check.format != "ICO":
            raise SystemExit("Generated Windows icon is not a valid ICO file")
    return WINDOWS_ICON


def build_macos_icon(image: Image.Image) -> Path:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    image.save(MACOS_ICON, format="ICNS")
    with Image.open(MACOS_ICON) as check:
        if check.format != "ICNS":
            raise SystemExit("Generated macOS icon is not a valid ICNS file")
    return MACOS_ICON


def main() -> int:
    parser = argparse.ArgumentParser(description="Build validated platform-native release icons from the canonical PNG source.")
    parser.add_argument("--platform", choices=("windows", "macos"), required=True)
    args = parser.parse_args()

    image = _load_source()
    output = build_windows_icon(image) if args.platform == "windows" else build_macos_icon(image)
    print(f"Created validated release icon: {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
