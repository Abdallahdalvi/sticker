"""Production bulk sticker generator.

Run with::

    python server.py

The app is intentionally built around one locked Reliance 4G dongle label.
Every accepted spreadsheet row becomes one physical-size PDF page.
"""

from __future__ import annotations

import base64
import csv
import html
import io
import json
import os
import re
from pathlib import Path
from typing import Any

import qrcode
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import JSONResponse, Response, StreamingResponse
from fastapi.staticfiles import StaticFiles
from openpyxl import load_workbook
from pydantic import BaseModel, Field
from reportlab.graphics import renderPDF
from reportlab.graphics.barcode.code128 import Code128Auto
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas as rl_canvas
from svglib.svglib import svg2rlg
import uvicorn


ROOT_DIR = Path(__file__).resolve().parent
STATIC_DIR = ROOT_DIR / "static"
VECTOR_ART_ASSET = STATIC_DIR / "reference_vector_paths.json"
TECH_SVG_ASSET = STATIC_DIR / "tech_section.svg"
RELIANCE_SVG_ASSET = STATIC_DIR / "reliance_logo.svg"
MAKE_IN_INDIA_SVG_ASSET = STATIC_DIR / "make_in_india_vector.svg"

LABEL_WIDTH_MM = 24.08
LABEL_HEIGHT_MM = 74.08
A4_COLUMNS = 7
A4_ROWS = 3
A4_LABELS_PER_PAGE = A4_COLUMNS * A4_ROWS
A4_GUTTER_MM = 4.0
UPPER_SEPARATOR_Y_MM = 22.05
LOWER_SEPARATOR_Y_MM = 44.72
SEPARATOR_X_START_MM = 1.75
SEPARATOR_X_END_MM = 22.05
SEPARATOR_WIDTH_MM = 0.08
SEPARATOR_COVER_HEIGHT_MM = 0.80
TECH_SECTION_SIZE_MM = LOWER_SEPARATOR_Y_MM - UPPER_SEPARATOR_Y_MM
TECH_SECTION_X_MM = (LABEL_WIDTH_MM - TECH_SECTION_SIZE_MM) / 2
RELIANCE_X_MM = 2.5392
RELIANCE_TOP_MM = 4.3178
RELIANCE_WIDTH_MM = 19.2132
RELIANCE_HEIGHT_MM = 14.4350
MAKE_IN_INDIA_X_MM = 1.6928
MAKE_IN_INDIA_TOP_MM = 68.4660
MAKE_IN_INDIA_WIDTH_MM = 9.4796
MAKE_IN_INDIA_HEIGHT_MM = 4.3280
MAX_UPLOAD_BYTES = 10 * 1024 * 1024
MAX_ROWS = 20_000
REQUIRED_COLUMNS = ("Device ID", "IMEI", "CCID")


def _load_vector_art() -> dict[str, list[list[list[float]]]]:
    if not VECTOR_ART_ASSET.exists():
        raise RuntimeError(f"Vector artwork is missing: {VECTOR_ART_ASSET}")
    payload = json.loads(VECTOR_ART_ASSET.read_text(encoding="utf-8"))
    return {"white": payload.get("white", []), "green": payload.get("green", [])}


VECTOR_ART = _load_vector_art()


def _load_svg_asset(asset: Path, label: str) -> tuple[str, Any]:
    if not asset.exists():
        raise RuntimeError(f"{label} SVG is missing: {asset}")
    source = asset.read_text(encoding="utf-8")
    match = re.search(r"<svg\b[^>]*>(.*)</svg>\s*$", source, flags=re.DOTALL)
    if not match:
        raise RuntimeError(f"{label} SVG is invalid: {asset}")
    drawing = svg2rlg(str(asset))
    if drawing is None:
        raise RuntimeError(f"{label} SVG could not be parsed: {asset}")
    return match.group(1), drawing


TECH_SVG_INNER, TECH_DRAWING = _load_svg_asset(TECH_SVG_ASSET, "Technical section")
RELIANCE_SVG_INNER, RELIANCE_DRAWING = _load_svg_asset(RELIANCE_SVG_ASSET, "Reliance logo")
MAKE_IN_INDIA_SVG_INNER, MAKE_IN_INDIA_DRAWING = _load_svg_asset(
    MAKE_IN_INDIA_SVG_ASSET, "Make in India"
)

HEADER_ALIASES = {
    "device id": "Device ID",
    "deviceid": "Device ID",
    "device_id": "Device ID",
    "imei": "IMEI",
    "ccid": "CCID",
    "iccid": "CCID",
    "sim number": "CCID",
    "sim no": "CCID",
    "sim_number": "CCID",
    "qr data": "QR Data",
    "qr": "QR Data",
    "qr code": "QR Data",
    "qr payload": "QR Data",
    "model": "Model",
}

# All Y values are measured from the top edge of the physical label.
LAYOUT = {
    "info_label_x": 1.81,
    "info_colon_x": 10.12,
    "model_value_x": 12.07,
    "device_value_x": 11.98,
    "model_top": 47.69,
    "device_top": 50.61,
    "barcode_x": 2.10,
    "barcode_w": 19.85,
    "barcode_h": 5.00,
    "barcode_text_x": 2.10,
    "ccid_barcode_top": 52.49,
    "ccid_text_top": 58.89,
    "imei_barcode_top": 60.28,
    "imei_text_top": 66.72,
    "qr_x": 16.40,
    "qr_top": 67.72,
    "qr_size": 5.80,
}


app = FastAPI(title="Bulk Sticker Generator", docs_url=None, redoc_url=None)


@app.middleware("http")
async def prevent_stale_app_shell(request, call_next):
    response = await call_next(request)
    if request.url.path in ("/", "/index.html"):
        response.headers["Cache-Control"] = "no-store, max-age=0"
        response.headers["Pragma"] = "no-cache"
    return response


def _register_fonts() -> tuple[str, str]:
    candidates = [
        ("C:/Windows/Fonts/arial.ttf", "C:/Windows/Fonts/arialbd.ttf"),
        ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
         "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
    ]
    for normal_path, bold_path in candidates:
        if Path(normal_path).exists() and Path(bold_path).exists():
            try:
                pdfmetrics.registerFont(TTFont("StickerArial", normal_path))
                pdfmetrics.registerFont(TTFont("StickerArial-Bold", bold_path))
                return "StickerArial", "StickerArial-Bold"
            except Exception:
                continue
    return "Helvetica", "Helvetica-Bold"


FONT_NORMAL, FONT_BOLD = _register_fonts()


def _clean_header(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip())


def _canonical_header(value: Any) -> str:
    cleaned = _clean_header(value)
    return HEADER_ALIASES.get(cleaned.casefold(), cleaned)


def _clean_cell_value(value: Any) -> str:
    """Normalize spreadsheet text without performing numeric conversion."""
    return "" if value is None else str(value).replace("\u200b", "").replace("\ufeff", "").strip()


def _canonicalize_rows(headers: list[Any], raw_rows: list[list[Any]]) -> list[dict[str, str]]:
    canonical_headers = [_canonical_header(h) for h in headers]
    if any(not h for h in canonical_headers):
        raise ValueError("Every spreadsheet column must have a header.")
    duplicates = sorted({h for h in canonical_headers if canonical_headers.count(h) > 1})
    if duplicates:
        raise ValueError(f"Duplicate columns after normalization: {', '.join(duplicates)}")

    rows: list[dict[str, str]] = []
    for raw in raw_rows:
        values = list(raw) + [""] * max(0, len(canonical_headers) - len(raw))
        row = {
            header: _clean_cell_value(values[idx])
            for idx, header in enumerate(canonical_headers)
        }
        if any(row.values()):
            rows.append(row)
    return rows


def _read_csv(data: bytes) -> list[dict[str, str]]:
    try:
        text = data.decode("utf-8-sig")
    except UnicodeDecodeError as exc:
        raise ValueError("CSV must be UTF-8 encoded.") from exc
    reader = csv.reader(io.StringIO(text))
    try:
        headers = next(reader)
    except StopIteration as exc:
        raise ValueError("The CSV file is empty.") from exc
    return _canonicalize_rows(headers, [list(row) for row in reader])


def _read_xlsx(data: bytes) -> list[dict[str, str]]:
    try:
        workbook = load_workbook(io.BytesIO(data), read_only=True, data_only=False)
    except Exception as exc:
        raise ValueError(f"Could not open XLSX file: {exc}") from exc
    sheet = workbook.active
    iterator = sheet.iter_rows()
    try:
        header_cells = next(iterator)
    except StopIteration as exc:
        raise ValueError("The XLSX file is empty.") from exc

    headers = [_canonical_header(cell.value) for cell in header_cells]
    if any(not header for header in headers):
        raise ValueError("Every spreadsheet column must have a header.")
    duplicates = sorted({h for h in headers if headers.count(h) > 1})
    if duplicates:
        raise ValueError(f"Duplicate columns after normalization: {', '.join(duplicates)}")

    rows: list[dict[str, str]] = []
    numeric_identifier_cells: list[str] = []
    formula_identifier_cells: list[str] = []
    for cells in iterator:
        row: dict[str, str] = {}
        for idx, header in enumerate(headers):
            cell = cells[idx] if idx < len(cells) else None
            value = None if cell is None else cell.value
            if header in REQUIRED_COLUMNS and cell is not None:
                if cell.data_type == "n" and value not in (None, ""):
                    numeric_identifier_cells.append(cell.coordinate)
                if cell.data_type == "f":
                    formula_identifier_cells.append(cell.coordinate)
            row[header] = _clean_cell_value(value)
        if any(row.values()):
            rows.append(row)

    if numeric_identifier_cells:
        sample = ", ".join(numeric_identifier_cells[:8])
        raise ValueError(
            "Identifier cells must be stored as Text in Excel; numeric cells can lose digits "
            f"or leading zeros. Fix cells such as {sample}."
        )
    if formula_identifier_cells:
        sample = ", ".join(formula_identifier_cells[:8])
        raise ValueError(f"Identifier cells cannot contain formulas. Fix cells such as {sample}.")
    return rows


def read_device_file(data: bytes, filename: str) -> list[dict[str, str]]:
    if len(data) > MAX_UPLOAD_BYTES:
        raise ValueError("File is larger than the 10 MB upload limit.")
    suffix = Path(filename).suffix.lower()
    if suffix == ".csv":
        rows = _read_csv(data)
    elif suffix == ".xlsx":
        rows = _read_xlsx(data)
    elif suffix == ".xls":
        raise ValueError("Legacy .xls files are not accepted. Save the workbook as .xlsx or CSV.")
    else:
        raise ValueError("Upload an .xlsx or .csv file.")
    if len(rows) > MAX_ROWS:
        raise ValueError(f"Workbook contains more than {MAX_ROWS:,} data rows.")
    return rows


def validate_rows(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    errors: list[dict[str, Any]] = []
    if not rows:
        return [{"row": 0, "field": "File", "message": "No data rows were found."}]

    seen: dict[str, dict[str, int]] = {field: {} for field in ("Device ID", "IMEI", "CCID")}
    for row_number, raw_row in enumerate(rows, start=2):
        row = {str(key): "" if value is None else str(value).strip() for key, value in raw_row.items()}
        for field in REQUIRED_COLUMNS:
            if not row.get(field):
                errors.append({"row": row_number, "field": field, "message": "Required value is blank."})

        device_id = row.get("Device ID", "")
        imei = row.get("IMEI", "")
        ccid = row.get("CCID", "")
        if device_id and len(device_id) > 32:
            errors.append({"row": row_number, "field": "Device ID", "message": "Maximum length is 32 characters."})
        if imei and not re.fullmatch(r"\d{15}", imei):
            errors.append({"row": row_number, "field": "IMEI", "message": "IMEI must contain exactly 15 digits."})
        if ccid and not re.fullmatch(r"\d{18,22}", ccid):
            errors.append({"row": row_number, "field": "CCID", "message": "CCID must contain 18 to 22 digits."})
        for field, value in (("Device ID", device_id), ("IMEI", imei), ("CCID", ccid)):
            if not value:
                continue
            previous = seen[field].get(value)
            if previous is not None:
                errors.append({"row": row_number, "field": field, "message": f"Duplicate of spreadsheet row {previous}."})
            else:
                seen[field][value] = row_number
    return errors


def _barcode_spec(value: str, width_pt: float, printer_dpi: int) -> tuple[Code128Auto, float, float]:
    probe = Code128Auto(value, barWidth=1, barHeight=10, quiet=0, humanReadable=False)
    modules = float(probe.width)
    quiet_modules = 10.0
    module_width = width_pt / (modules + quiet_modules * 2)
    minimum = 2 * 72.0 / printer_dpi
    if module_width + 1e-9 < minimum:
        raise ValueError(
            f"Barcode value {value!r} is too dense for {printer_dpi} DPI at this label width. "
            "Use a 600 DPI printer or a wider label."
        )
    barcode = Code128Auto(value, barWidth=module_width, barHeight=10, quiet=0, humanReadable=False)
    return barcode, module_width, quiet_modules * module_width


def _barcode_pattern(value: str) -> tuple[str, int]:
    barcode = Code128Auto(value, barWidth=1, barHeight=10, quiet=0, humanReadable=False)
    barcode.validate()
    barcode.encode()
    pattern = barcode.decompose()
    total_modules = sum(ord(char.upper()) - ord("A") + 1 for char in pattern)
    return pattern, total_modules


def _qr_matrix(value: str) -> list[list[bool]]:
    qr = qrcode.QRCode(version=None, error_correction=qrcode.constants.ERROR_CORRECT_M, box_size=1, border=4)
    qr.add_data(value)
    qr.make(fit=True)
    return qr.get_matrix()


def _qr_payload(row: dict[str, str]) -> str:
    return f"Device ID: {row['Device ID']}\nCCID: {row['CCID']}\nIMEI: {row['IMEI']}"


def _draw_fitted_text(
    canvas: rl_canvas.Canvas,
    text: str,
    x_pt: float,
    baseline_y_pt: float,
    max_width_pt: float,
    size_pt: float,
    minimum_pt: float = 3.2,
) -> None:
    fitted = size_pt
    while fitted > minimum_pt and pdfmetrics.stringWidth(text, FONT_BOLD, fitted) > max_width_pt:
        fitted -= 0.1
    canvas.setFillColorRGB(1, 1, 1)
    canvas.setFont(FONT_BOLD, fitted)
    canvas.drawString(x_pt, baseline_y_pt, text)


def _draw_barcode(canvas: rl_canvas.Canvas, value: str, top_mm: float, printer_dpi: int) -> None:
    x_pt = LAYOUT["barcode_x"] * mm
    width_pt = LAYOUT["barcode_w"] * mm
    height_pt = LAYOUT["barcode_h"] * mm
    y_pt = (LABEL_HEIGHT_MM - top_mm - LAYOUT["barcode_h"]) * mm
    barcode, _, quiet_pt = _barcode_spec(value, width_pt, printer_dpi)
    barcode.barHeight = height_pt
    canvas.setFillColorRGB(1, 1, 1)
    canvas.rect(x_pt, y_pt, width_pt, height_pt, fill=1, stroke=0)
    canvas.setFillColorRGB(0, 0, 0)
    barcode.drawOn(canvas, x_pt + quiet_pt, y_pt)


def _draw_traced_contours(
    canvas: rl_canvas.Canvas,
    contours: list[list[list[float]]],
    color: tuple[float, float, float],
) -> None:
    path = canvas.beginPath()
    for contour in contours:
        if len(contour) < 3:
            continue
        first_x, first_y = contour[0]
        path.moveTo(first_x * mm, (LABEL_HEIGHT_MM - first_y) * mm)
        for x_value, y_value in contour[1:]:
            path.lineTo(x_value * mm, (LABEL_HEIGHT_MM - y_value) * mm)
        path.close()
    canvas.setFillColorRGB(*color)
    canvas.drawPath(path, stroke=0, fill=1, fillMode=0)


def _draw_svg_in_box(
    canvas: rl_canvas.Canvas,
    drawing: Any,
    x_mm: float,
    top_mm: float,
    width_mm: float,
    height_mm: float,
) -> None:
    """Place one supplied SVG in a fixed physical box without distortion."""
    scale = min((width_mm * mm) / drawing.width, (height_mm * mm) / drawing.height)
    rendered_width = drawing.width * scale
    rendered_height = drawing.height * scale
    left = x_mm * mm + (width_mm * mm - rendered_width) / 2
    bottom = (LABEL_HEIGHT_MM - top_mm - height_mm) * mm + (height_mm * mm - rendered_height) / 2
    canvas.saveState()
    canvas.translate(left, bottom)
    canvas.scale(scale, scale)
    renderPDF.draw(drawing, canvas, 0, 0)
    canvas.restoreState()


def _draw_vector_separators(canvas: rl_canvas.Canvas) -> None:
    for separator_y_mm in (UPPER_SEPARATOR_Y_MM, LOWER_SEPARATOR_Y_MM):
        cover_y = (LABEL_HEIGHT_MM - separator_y_mm - SEPARATOR_COVER_HEIGHT_MM / 2) * mm
        canvas.setFillColorRGB(0, 0, 0)
        canvas.rect(0, cover_y, LABEL_WIDTH_MM * mm, SEPARATOR_COVER_HEIGHT_MM * mm, fill=1, stroke=0)
        line_y = (LABEL_HEIGHT_MM - separator_y_mm) * mm
        canvas.setStrokeColorRGB(0.61, 0.62, 0.59)
        canvas.setLineWidth(SEPARATOR_WIDTH_MM * mm)
        canvas.setLineCap(1)
        canvas.line(SEPARATOR_X_START_MM * mm, line_y, SEPARATOR_X_END_MM * mm, line_y)


def _draw_vector_art(canvas: rl_canvas.Canvas) -> None:
    _draw_traced_contours(canvas, VECTOR_ART["white"], (1, 1, 1))
    _draw_traced_contours(canvas, VECTOR_ART["green"], (55 / 255, 173 / 255, 78 / 255))
    _draw_svg_in_box(
        canvas, RELIANCE_DRAWING, RELIANCE_X_MM, RELIANCE_TOP_MM, RELIANCE_WIDTH_MM, RELIANCE_HEIGHT_MM
    )
    _draw_svg_in_box(
        canvas,
        TECH_DRAWING,
        TECH_SECTION_X_MM,
        UPPER_SEPARATOR_Y_MM,
        TECH_SECTION_SIZE_MM,
        TECH_SECTION_SIZE_MM,
    )
    _draw_svg_in_box(
        canvas,
        MAKE_IN_INDIA_DRAWING,
        MAKE_IN_INDIA_X_MM,
        MAKE_IN_INDIA_TOP_MM,
        MAKE_IN_INDIA_WIDTH_MM,
        MAKE_IN_INDIA_HEIGHT_MM,
    )
    _draw_vector_separators(canvas)


def _draw_sticker(canvas: rl_canvas.Canvas, row: dict[str, str], printer_dpi: int) -> None:
    width_pt = LABEL_WIDTH_MM * mm
    height_pt = LABEL_HEIGHT_MM * mm
    canvas.setFillColorRGB(0, 0, 0)
    canvas.rect(0, 0, width_pt, height_pt, fill=1, stroke=0)
    _draw_vector_art(canvas)

    model = row.get("Model", "") or "4G Dongle"
    model_y = (LABEL_HEIGHT_MM - LAYOUT["model_top"]) * mm
    device_y = (LABEL_HEIGHT_MM - LAYOUT["device_top"]) * mm
    label_x = LAYOUT["info_label_x"] * mm
    colon_x = LAYOUT["info_colon_x"] * mm
    model_value_x = LAYOUT["model_value_x"] * mm
    device_value_x = LAYOUT["device_value_x"] * mm
    value_right = (LABEL_WIDTH_MM - 1.0) * mm
    _draw_fitted_text(canvas, "Model", label_x, model_y, colon_x - label_x, 4.78)
    _draw_fitted_text(canvas, ":", colon_x, model_y, 1.2 * mm, 4.78)
    _draw_fitted_text(canvas, model, model_value_x, model_y, value_right - model_value_x, 4.78)
    _draw_fitted_text(canvas, "Device ID", label_x, device_y, colon_x - label_x, 4.78)
    _draw_fitted_text(canvas, ":", colon_x, device_y, 1.2 * mm, 4.78)
    _draw_fitted_text(canvas, row["Device ID"], device_value_x, device_y, value_right - device_value_x, 4.78)

    _draw_barcode(canvas, row["CCID"], LAYOUT["ccid_barcode_top"], printer_dpi)
    _draw_barcode(canvas, row["IMEI"], LAYOUT["imei_barcode_top"], printer_dpi)
    ccid_y = (LABEL_HEIGHT_MM - LAYOUT["ccid_text_top"]) * mm
    imei_y = (LABEL_HEIGHT_MM - LAYOUT["imei_text_top"]) * mm
    barcode_text_x = LAYOUT["barcode_text_x"] * mm
    barcode_text_width = (LABEL_WIDTH_MM - LAYOUT["barcode_text_x"] - 1.0) * mm
    _draw_fitted_text(canvas, f"CCID {row['CCID']}", barcode_text_x, ccid_y, barcode_text_width, 4.14)
    _draw_fitted_text(canvas, f"IMEI {row['IMEI']}", barcode_text_x, imei_y, barcode_text_width, 4.14)

    matrix = _qr_matrix(_qr_payload(row))
    qr_x = LAYOUT["qr_x"] * mm
    qr_size = LAYOUT["qr_size"] * mm
    qr_y = (LABEL_HEIGHT_MM - LAYOUT["qr_top"] - LAYOUT["qr_size"]) * mm
    cell = qr_size / len(matrix)
    canvas.setFillColorRGB(1, 1, 1)
    canvas.rect(qr_x, qr_y, qr_size, qr_size, fill=1, stroke=0)
    canvas.setFillColorRGB(0, 0, 0)
    for row_idx, cells in enumerate(matrix):
        for col_idx, dark in enumerate(cells):
            if dark:
                canvas.rect(qr_x + col_idx * cell, qr_y + (len(matrix) - row_idx - 1) * cell, cell, cell, fill=1, stroke=0)


def render_to_pdf(
    rows: list[dict[str, str]],
    printer_dpi: int = 600,
    page_format: str = "label",
) -> io.BytesIO:
    if printer_dpi not in (300, 600):
        raise ValueError("Printer DPI must be 300 or 600.")
    if page_format not in ("label", "a4"):
        raise ValueError("Page format must be 'label' or 'a4'.")
    errors = validate_rows(rows)
    if errors:
        raise ValueError("Input rows failed validation.")
    for row in rows:
        _barcode_spec(row["CCID"], LAYOUT["barcode_w"] * mm, printer_dpi)
        _barcode_spec(row["IMEI"], LAYOUT["barcode_w"] * mm, printer_dpi)

    buffer = io.BytesIO()
    pagesize = (LABEL_WIDTH_MM * mm, LABEL_HEIGHT_MM * mm) if page_format == "label" else A4
    canvas = rl_canvas.Canvas(buffer, pagesize=pagesize, pageCompression=1)
    canvas.setTitle("Bulk device stickers" if page_format == "label" else "A4 bulk device sticker sheets")

    if page_format == "label":
        for index, row in enumerate(rows):
            if index:
                canvas.showPage()
            _draw_sticker(canvas, row, printer_dpi)
    else:
        a4_width_mm = A4[0] / mm
        a4_height_mm = A4[1] / mm
        grid_width_mm = A4_COLUMNS * LABEL_WIDTH_MM + (A4_COLUMNS - 1) * A4_GUTTER_MM
        grid_height_mm = A4_ROWS * LABEL_HEIGHT_MM + (A4_ROWS - 1) * A4_GUTTER_MM
        left_mm = (a4_width_mm - grid_width_mm) / 2
        top_mm = (a4_height_mm - grid_height_mm) / 2
        for index, row in enumerate(rows):
            slot = index % A4_LABELS_PER_PAGE
            if index and slot == 0:
                canvas.showPage()
            column = slot % A4_COLUMNS
            grid_row = slot // A4_COLUMNS
            x_mm = left_mm + column * (LABEL_WIDTH_MM + A4_GUTTER_MM)
            y_mm = a4_height_mm - top_mm - LABEL_HEIGHT_MM - grid_row * (LABEL_HEIGHT_MM + A4_GUTTER_MM)
            canvas.saveState()
            canvas.translate(x_mm * mm, y_mm * mm)
            _draw_sticker(canvas, row, printer_dpi)
            canvas.restoreState()
    canvas.save()
    buffer.seek(0)
    return buffer


def _svg_contours(contours: list[list[list[float]]], fill: str) -> str:
    pieces: list[str] = []
    for contour in contours:
        if len(contour) < 3:
            continue
        commands = [f"M{contour[0][0]:.4f},{contour[0][1]:.4f}"]
        commands.extend(f"L{x_value:.4f},{y_value:.4f}" for x_value, y_value in contour[1:])
        commands.append("Z")
        pieces.append("".join(commands))
    return f'<path d="{"".join(pieces)}" fill="{fill}" fill-rule="evenodd"/>'


def _svg_vector_art() -> str:
    separators = "".join(
        f'<rect x="0" y="{separator_y - SEPARATOR_COVER_HEIGHT_MM / 2:.3f}" '
        f'width="{LABEL_WIDTH_MM}" height="{SEPARATOR_COVER_HEIGHT_MM}" fill="#000"/>'
        f'<line x1="{SEPARATOR_X_START_MM}" y1="{separator_y}" '
        f'x2="{SEPARATOR_X_END_MM}" y2="{separator_y}" stroke="#9b9d96" '
        f'stroke-width="{SEPARATOR_WIDTH_MM}" stroke-linecap="round"/>'
        for separator_y in (UPPER_SEPARATOR_Y_MM, LOWER_SEPARATOR_Y_MM)
    )
    tech_section = (
        f'<svg x="{TECH_SECTION_X_MM:.4f}" y="{UPPER_SEPARATOR_Y_MM:.4f}" '
        f'width="{TECH_SECTION_SIZE_MM:.4f}" height="{TECH_SECTION_SIZE_MM:.4f}" '
        'viewBox="0 0 810 809.999993" preserveAspectRatio="xMidYMid meet">'
        f'{TECH_SVG_INNER}</svg>'
    )
    reliance_logo = (
        f'<svg x="{RELIANCE_X_MM:.4f}" y="{RELIANCE_TOP_MM:.4f}" '
        f'width="{RELIANCE_WIDTH_MM:.4f}" height="{RELIANCE_HEIGHT_MM:.4f}" '
        'viewBox="0 0 1600 1203" preserveAspectRatio="xMidYMid meet">'
        f'{RELIANCE_SVG_INNER}</svg>'
    )
    make_in_india = (
        f'<svg x="{MAKE_IN_INDIA_X_MM:.4f}" y="{MAKE_IN_INDIA_TOP_MM:.4f}" '
        f'width="{MAKE_IN_INDIA_WIDTH_MM:.4f}" height="{MAKE_IN_INDIA_HEIGHT_MM:.4f}" '
        'viewBox="0 0 1983 903" preserveAspectRatio="xMidYMid meet">'
        f'{MAKE_IN_INDIA_SVG_INNER}</svg>'
    )
    return (
        _svg_contours(VECTOR_ART["white"], "#fff")
        + _svg_contours(VECTOR_ART["green"], "#37ad4e")
        + reliance_logo
        + tech_section
        + make_in_india
        + separators
    )


def _svg_barcode(value: str, top_mm: float, printer_dpi: int) -> str:
    pattern, modules = _barcode_pattern(value)
    total = modules + 20
    module_w = LAYOUT["barcode_w"] / total
    minimum_mm = (2 * 25.4) / printer_dpi
    if module_w + 1e-9 < minimum_mm:
        raise ValueError(f"Barcode {value!r} is too dense for {printer_dpi} DPI.")
    x = LAYOUT["barcode_x"] + 10 * module_w
    pieces = [
        f'<rect x="{LAYOUT["barcode_x"]:.4f}" y="{top_mm:.4f}" width="{LAYOUT["barcode_w"]:.4f}" '
        f'height="{LAYOUT["barcode_h"]:.4f}" fill="#fff"/>'
    ]
    is_bar = True
    for char in pattern:
        units = ord(char.upper()) - ord("A") + 1
        width = units * module_w
        if is_bar:
            pieces.append(
                f'<rect x="{x:.4f}" y="{top_mm:.4f}" width="{width:.4f}" '
                f'height="{LAYOUT["barcode_h"]:.4f}" fill="#000"/>'
            )
        x += width
        is_bar = not is_bar
    return "".join(pieces)


def render_to_svg(row: dict[str, str], printer_dpi: int = 600) -> str:
    errors = validate_rows([row])
    if errors:
        raise ValueError(errors[0]["message"])
    vector_art = _svg_vector_art()
    model = html.escape(row.get("Model", "") or "4G Dongle")
    device = html.escape(row["Device ID"])
    ccid = html.escape(row["CCID"])
    imei = html.escape(row["IMEI"])
    qr_matrix = _qr_matrix(_qr_payload(row))
    qr_cell = LAYOUT["qr_size"] / len(qr_matrix)
    qr_parts = [
        f'<rect x="{LAYOUT["qr_x"]}" y="{LAYOUT["qr_top"]}" width="{LAYOUT["qr_size"]}" '
        f'height="{LAYOUT["qr_size"]}" fill="#fff"/>'
    ]
    for row_idx, cells in enumerate(qr_matrix):
        for col_idx, dark in enumerate(cells):
            if dark:
                qr_parts.append(
                    f'<rect x="{LAYOUT["qr_x"] + col_idx * qr_cell:.4f}" '
                    f'y="{LAYOUT["qr_top"] + row_idx * qr_cell:.4f}" '
                    f'width="{qr_cell:.4f}" height="{qr_cell:.4f}" fill="#000"/>'
                )
    return (
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{LABEL_WIDTH_MM}mm" height="{LABEL_HEIGHT_MM}mm" '
        f'viewBox="0 0 {LABEL_WIDTH_MM} {LABEL_HEIGHT_MM}">'
        '<rect width="100%" height="100%" fill="#000"/>'
        f'{vector_art}'
        '<g fill="#fff" font-family="Arial,Helvetica,sans-serif" font-weight="700">'
        f'<text x="{LAYOUT["info_label_x"]}" y="{LAYOUT["model_top"]}" font-size="1.686">Model</text>'
        f'<text x="{LAYOUT["info_colon_x"]}" y="{LAYOUT["model_top"]}" font-size="1.686">:</text>'
        f'<text x="{LAYOUT["model_value_x"]}" y="{LAYOUT["model_top"]}" font-size="1.686">{model}</text>'
        f'<text x="{LAYOUT["info_label_x"]}" y="{LAYOUT["device_top"]}" font-size="1.686">Device ID</text>'
        f'<text x="{LAYOUT["info_colon_x"]}" y="{LAYOUT["device_top"]}" font-size="1.686">:</text>'
        f'<text x="{LAYOUT["device_value_x"]}" y="{LAYOUT["device_top"]}" font-size="1.686">{device}</text>'
        f'<text x="{LAYOUT["barcode_text_x"]}" y="{LAYOUT["ccid_text_top"]}" font-size="1.460">CCID {ccid}</text>'
        f'<text x="{LAYOUT["barcode_text_x"]}" y="{LAYOUT["imei_text_top"]}" font-size="1.460">IMEI {imei}</text>'
        '</g>'
        f'{_svg_barcode(row["CCID"], LAYOUT["ccid_barcode_top"], printer_dpi)}'
        f'{_svg_barcode(row["IMEI"], LAYOUT["imei_barcode_top"], printer_dpi)}'
        f'{"".join(qr_parts)}'
        '</svg>'
    )


def _raise_validation(errors: list[dict[str, Any]], message: str = "Spreadsheet validation failed.") -> None:
    raise HTTPException(status_code=422, detail={"message": message, "errors": errors[:200]})


@app.post("/api/upload")
async def upload_file(file: UploadFile = File(...)):
    data = await file.read()
    try:
        rows = read_device_file(data, file.filename or "devices.xlsx")
    except ValueError as exc:
        _raise_validation([{"row": 0, "field": "File", "message": str(exc)}])
    errors = validate_rows(rows)
    if errors:
        _raise_validation(errors)
    return JSONResponse({
        "columns": [*REQUIRED_COLUMNS, "Model"],
        "rows": rows,
        "count": len(rows),
        "label": {"widthMm": LABEL_WIDTH_MM, "heightMm": LABEL_HEIGHT_MM},
    })


class PreviewRequest(BaseModel):
    row: dict[str, Any]
    printer_dpi: int = Field(default=600, alias="printerDpi")


@app.post("/api/preview")
async def preview_svg(request: PreviewRequest):
    row = {key: "" if value is None else str(value).strip() for key, value in request.row.items()}
    try:
        svg = render_to_svg(row, request.printer_dpi)
    except ValueError as exc:
        _raise_validation([{"row": 0, "field": "Preview", "message": str(exc)}])
    encoded = base64.b64encode(svg.encode("utf-8")).decode("ascii")
    return JSONResponse({"svg": f"data:image/svg+xml;base64,{encoded}"})


class GenerateRequest(BaseModel):
    rows: list[dict[str, Any]]
    start: int = 1
    end: int | None = None
    printer_dpi: int = Field(default=600, alias="printerDpi")
    page_format: str = Field(default="label", alias="pageFormat")


@app.post("/api/generate")
async def generate_pdf(request: GenerateRequest):
    rows = [{key: "" if value is None else str(value).strip() for key, value in row.items()} for row in request.rows]
    errors = validate_rows(rows)
    if errors:
        _raise_validation(errors)
    end = len(rows) if request.end is None else request.end
    if request.start < 1 or end < request.start or end > len(rows):
        _raise_validation([{"row": 0, "field": "Range", "message": f"Choose a range between 1 and {len(rows)}."}])
    selected = rows[request.start - 1:end]
    try:
        pdf = render_to_pdf(selected, request.printer_dpi, request.page_format)
    except ValueError as exc:
        _raise_validation([{"row": 0, "field": "Printer", "message": str(exc)}])
    suffix = "_a4" if request.page_format == "a4" else ""
    filename = f"stickers_{request.start}-{end}{suffix}.pdf"
    page_count = len(selected) if request.page_format == "label" else (len(selected) + A4_LABELS_PER_PAGE - 1) // A4_LABELS_PER_PAGE
    return StreamingResponse(
        pdf,
        media_type="application/pdf",
        headers={
            "Content-Disposition": f'attachment; filename="{filename}"',
            "X-Sticker-Count": str(len(selected)),
            "X-Page-Count": str(page_count),
        },
    )


@app.get("/api/health")
async def health():
    return {"ok": True, "labelMm": [LABEL_WIDTH_MM, LABEL_HEIGHT_MM]}


@app.get("/api/reference.svg")
async def reference_svg():
    svg = (
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{LABEL_WIDTH_MM}mm" height="{LABEL_HEIGHT_MM}mm" '
        f'viewBox="0 0 {LABEL_WIDTH_MM} {LABEL_HEIGHT_MM}">'
        '<rect width="100%" height="100%" fill="#000"/>'
        f'{_svg_vector_art()}'
        '</svg>'
    )
    return Response(svg, media_type="image/svg+xml", headers={"Cache-Control": "no-store"})


app.mount("/", StaticFiles(directory=str(STATIC_DIR), html=True), name="static")


if __name__ == "__main__":
    app_host = os.getenv("APP_HOST", "127.0.0.1")
    try:
        app_port = int(os.getenv("APP_PORT", "8502"))
    except ValueError as exc:
        raise RuntimeError("APP_PORT must be a valid integer.") from exc
    if not 1 <= app_port <= 65535:
        raise RuntimeError("APP_PORT must be between 1 and 65535.")
    uvicorn.run(app, host=app_host, port=app_port, reload=False)
