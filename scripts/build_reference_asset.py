"""Build the static label artwork from the supplied reference PDF.

The variable sample fields in the lower section are removed.  The brand header,
feature icons, dividers, and Make in India artwork remain as the print backdrop.
"""

from __future__ import annotations

import argparse
import shutil
import subprocess
import tempfile
from pathlib import Path

from PIL import Image, ImageDraw


def find_pdftoppm(explicit: str | None) -> str:
    if explicit:
        return explicit
    discovered = shutil.which("pdftoppm")
    if discovered:
        return discovered
    bundled = Path.home() / ".cache/codex-runtimes/codex-primary-runtime/dependencies/native/poppler/Library/bin/pdftoppm.exe"
    if bundled.exists():
        return str(bundled)
    raise FileNotFoundError("pdftoppm was not found; install Poppler or pass --pdftoppm.")


def build(source_pdf: Path, output_png: Path, pdftoppm: str) -> None:
    with tempfile.TemporaryDirectory() as temp_dir:
        prefix = Path(temp_dir) / "reference"
        subprocess.run(
            [pdftoppm, "-png", "-r", "600", "-f", "1", "-singlefile", str(source_pdf), str(prefix)],
            check=True,
        )
        original = Image.open(prefix.with_suffix(".png")).convert("RGB")

    width, height = original.size
    result = original.copy()
    draw = ImageDraw.Draw(result)
    dynamic_top = round(height * 0.605)
    draw.rectangle((0, dynamic_top, width, height), fill="black")

    # Preserve only the supplied Make in India artwork from the lower section.
    lion_box = (
        round(width * 0.045),
        round(height * 0.910),
        round(width * 0.550),
        round(height * 0.985),
    )
    lion = original.crop(lion_box)
    result.paste(lion, lion_box[:2])

    output_png.parent.mkdir(parents=True, exist_ok=True)
    result.save(output_png, format="PNG", optimize=True)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("source_pdf", type=Path)
    parser.add_argument("output_png", type=Path)
    parser.add_argument("--pdftoppm")
    args = parser.parse_args()
    build(args.source_pdf, args.output_png, find_pdftoppm(args.pdftoppm))


if __name__ == "__main__":
    main()
