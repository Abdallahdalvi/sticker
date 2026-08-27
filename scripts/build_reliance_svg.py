from __future__ import annotations

import argparse
from pathlib import Path

import cv2


def build(source: Path, output: Path) -> None:
    image = cv2.imread(str(source), cv2.IMREAD_GRAYSCALE)
    if image is None:
        raise ValueError(f"Could not read {source}")
    height, width = image.shape
    _, mask = cv2.threshold(image, 127, 255, cv2.THRESH_BINARY)
    contours, _ = cv2.findContours(mask, cv2.RETR_LIST, cv2.CHAIN_APPROX_NONE)

    commands: list[str] = []
    kept = 0
    for contour in contours:
        if len(contour) < 3 or abs(cv2.contourArea(contour)) < 1.0:
            continue
        simplified = cv2.approxPolyDP(contour, 0.25, True)
        if len(simplified) < 3:
            continue
        points = simplified[:, 0, :]
        commands.append(f"M{points[0][0]},{points[0][1]}")
        commands.extend(f"L{x},{y}" for x, y in points[1:])
        commands.append("Z")
        kept += 1

    svg = (
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}" '
        f'viewBox="0 0 {width} {height}" preserveAspectRatio="xMidYMid meet">'
        f'<path fill="#fff" fill-rule="evenodd" d="{"".join(commands)}"/>'
        '</svg>'
    )
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(svg, encoding="utf-8")
    print(f"wrote {output} with {kept} contours ({len(svg):,} characters)")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Convert the supplied Reliance lockup into exact vector contours.")
    parser.add_argument("source", type=Path)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()
    build(args.source, args.output)
