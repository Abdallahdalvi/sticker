from __future__ import annotations

import argparse
import json
from pathlib import Path

import cv2
import numpy as np
from PIL import Image


LABEL_WIDTH_MM = 24.08
LABEL_HEIGHT_MM = 74.08


def _contours(mask: np.ndarray, width: int, height: int) -> list[list[list[float]]]:
    found, _ = cv2.findContours(mask, cv2.RETR_LIST, cv2.CHAIN_APPROX_TC89_KCOS)
    result: list[list[list[float]]] = []
    for contour in found:
        if len(contour) < 3 or abs(cv2.contourArea(contour)) < 0.6:
            continue
        epsilon = max(0.18, cv2.arcLength(contour, True) * 0.00035)
        simplified = cv2.approxPolyDP(contour, epsilon, True)
        if len(simplified) < 3:
            continue
        points = [
            [
                round(float(point[0][0]) * LABEL_WIDTH_MM / width, 4),
                round(float(point[0][1]) * LABEL_HEIGHT_MM / height, 4),
            ]
            for point in simplified
        ]
        result.append(points)
    return result


def build(source: Path, output: Path) -> None:
    image = np.array(Image.open(source).convert("RGB"))
    height, width, _ = image.shape
    red = image[:, :, 0].astype(np.int16)
    green = image[:, :, 1].astype(np.int16)
    blue = image[:, :, 2].astype(np.int16)

    white = ((red > 135) & (green > 135) & (blue > 135)).astype(np.uint8) * 255
    green_mask = ((green > 80) & (green > red * 1.28) & (green > blue * 1.20)).astype(np.uint8) * 255

    keep_top = np.zeros((height, width), dtype=np.uint8)
    keep_top[:1047, :] = 255
    keep_lion = np.zeros((height, width), dtype=np.uint8)
    keep_lion[1570:1742, 20:325] = 255
    white = cv2.bitwise_and(white, cv2.bitwise_or(keep_top, keep_lion))
    green_mask = cv2.bitwise_and(green_mask, keep_top)

    # All fixed artwork now comes from the three dedicated SVG assets.
    white[:, :] = 0
    green_mask[:, :] = 0

    payload = {
        "label_width_mm": LABEL_WIDTH_MM,
        "label_height_mm": LABEL_HEIGHT_MM,
        "white": _contours(white, width, height),
        "green": _contours(green_mask, width, height),
    }
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(payload, separators=(",", ":")), encoding="utf-8")
    print(f"wrote {output} with {len(payload['white'])} white and {len(payload['green'])} green contours")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Trace the fixed sticker artwork into vector contours.")
    parser.add_argument("source", type=Path)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()
    build(args.source, args.output)
