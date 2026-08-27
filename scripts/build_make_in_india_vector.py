from __future__ import annotations

import argparse
import math
from pathlib import Path

import cv2
import numpy as np
from reportlab.graphics import renderPM
from svglib.svglib import svg2rlg


def build(source: Path, output: Path, dpi: int = 144) -> None:
    drawing = svg2rlg(str(source))
    if drawing is None:
        raise ValueError(f"Could not parse {source}")

    scale = dpi / 72
    image = renderPM.drawToPIL(drawing, dpi=dpi).convert("RGB")
    x1, y1, x2, y2 = drawing.getBounds()
    crop_box = (
        math.floor(x1 * scale),
        math.floor((drawing.height - y2) * scale),
        math.ceil(x2 * scale),
        math.ceil((drawing.height - y1) * scale),
    )
    # The Canva export carries a one-pixel white frame around its black artboard.
    # Remove only that frame before tracing so it cannot appear on the label.
    crop_inset = max(1, round(scale))
    cropped = np.array(image.crop(crop_box))[crop_inset:-crop_inset, crop_inset:-crop_inset]
    gray = cv2.cvtColor(cropped, cv2.COLOR_RGB2GRAY)
    _, mask = cv2.threshold(gray, 127, 255, cv2.THRESH_BINARY)
    contours, _ = cv2.findContours(mask, cv2.RETR_LIST, cv2.CHAIN_APPROX_NONE)

    commands: list[str] = []
    kept = 0
    for contour in contours:
        if len(contour) < 3 or abs(cv2.contourArea(contour)) < 2.0:
            continue
        simplified = cv2.approxPolyDP(contour, 0.3, True)
        if len(simplified) < 3:
            continue
        points = simplified[:, 0, :]
        commands.append(f"M{points[0][0]},{points[0][1]}")
        commands.extend(f"L{x},{y}" for x, y in points[1:])
        commands.append("Z")
        kept += 1

    height, width = mask.shape
    svg = (
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}" '
        f'viewBox="0 0 {width} {height}" preserveAspectRatio="xMidYMid meet">'
        f'<path fill="#fff" fill-rule="evenodd" d="{"".join(commands)}"/>'
        '</svg>'
    )
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(svg, encoding="utf-8")
    print(f"wrote {output} with {kept} contours ({width} x {height})")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Convert the supplied Make in India SVG into path-only artwork.")
    parser.add_argument("source", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--dpi", type=int, default=144)
    args = parser.parse_args()
    build(args.source, args.output, args.dpi)
