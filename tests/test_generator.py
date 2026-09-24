from __future__ import annotations

import io
import shutil
import subprocess
from pathlib import Path

import pytest
from openpyxl import Workbook
from PIL import Image
from pypdf import PdfReader
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics

import server


ROWS = [
    {
        "Device ID": "K34632721",
        "IMEI": "356789012345678",
        "CCID": "89148000001234567890",
        "QR Data": "46796433678985434",
        "Model": "4G Dongle",
    },
    {
        "Device ID": "K34632722",
        "IMEI": "356789012345679",
        "CCID": "89148000001234567891",
        "QR Data": "46796433678985435",
        "Model": "4G Dongle",
    },
]


def test_csv_preserves_exact_identifier_text():
    csv_bytes = (
        "Device ID,IMEI,CCID,QR Data\n"
        "DEV-01,035678901234567,08914800000123456789,000123456789\n"
    ).encode()
    rows = server.read_device_file(csv_bytes, "devices.csv")
    assert rows[0]["IMEI"] == "035678901234567"
    assert rows[0]["CCID"] == "08914800000123456789"
    assert set(rows[0]) == {"Device ID", "Model", "CCID", "IMEI"}


def test_qr_data_column_is_not_required():
    csv_bytes = (
        "Device ID,IMEI,CCID\n"
        "DEV-01,035678901234567,08914800000123456789\n"
    ).encode()
    rows = server.read_device_file(csv_bytes, "devices.csv")
    assert server.validate_rows(rows) == []


def test_empty_unnamed_trailing_csv_column_is_ignored():
    csv_bytes = (
        "S/N,IMEI,CCID,Model,\n"
        "I26A108241,866224083845882,89918640507061557069,KRIG42ADAAI26,\n"
        ",866224083793777,,KRIG42ADAAI26,\n"
    ).encode()
    result = server._read_device_file_result(csv_bytes, "devices.csv")
    assert len(result.rows) == 1
    assert result.incomplete_rows == 1
    assert result.rows[0]["Device ID"] == "I26A108241"


def test_unnamed_csv_column_with_data_is_rejected():
    csv_bytes = (
        "S/N,IMEI,CCID,\n"
        "I26A108241,866224083845882,89918640507061557069,unexpected\n"
    ).encode()
    with pytest.raises(ValueError, match="Column 4 contains data but has no header"):
        server.read_device_file(csv_bytes, "devices.csv")


def test_serial_number_header_is_accepted_as_device_id():
    csv_bytes = (
        "S/N,IMEI,CCID\n"
        "5753120024723,866224084206712,89918640507061277734\n"
    ).encode()
    rows = server.read_device_file(csv_bytes, "devices.csv")
    assert rows[0]["Device ID"] == "5753120024723"


def test_qr_payload_uses_serial_number_label():
    payload = server._qr_payload(ROWS[0])
    assert payload.startswith("S/N: K34632721\n")
    assert "Device ID:" not in payload


def test_blank_and_incomplete_rows_are_skipped_and_extra_columns_are_ignored():
    csv_bytes = (
        "Tester Name,Device ID,IMEI,CCID,QR Data,Model\n"
        "A,5753120024723,866224084206712,89918640507061277734,1,KRIG42ACAAI26\n"
        "B,5753120024724,866224084206713,,2,KRIG42ACAAI26\n"
        ",,,,,\n"
    ).encode()
    result = server._read_device_file_result(csv_bytes, "devices.csv")
    assert len(result.rows) == 1
    assert result.blank_rows == 1
    assert result.incomplete_rows == 1
    assert result.ignored_columns == ("Tester Name", "QR Data")
    assert result.rows[0] == {
        "Device ID": "5753120024723",
        "Model": "KRIG42ACAAI26",
        "CCID": "89918640507061277734",
        "IMEI": "866224084206712",
    }


def test_long_model_and_serial_number_use_fixed_size_compact_lines():
    row = {
        "Device ID": "5753120024723",
        "Model": "KRIG42ACAAI26",
        "CCID": "89918640507061277734",
        "IMEI": "866224084206712",
    }
    svg = server.render_to_svg(row)
    assert "Model: KRIG42ACAAI26" in svg
    assert "S/N: 5753120024723" in svg
    assert "Device ID:" not in svg
    assert svg.count('font-size="1.6863"') == 2
    assert "StickerDynamic" in svg
    assert 'x="10.12"' not in svg
    assert f'transform="translate(0 -{server.CONTENT_UP_MM:.4f})"' in svg


def test_content_shift_balances_top_and_bottom_margins():
    top_margin = server.RELIANCE_TOP_MM - server.CONTENT_UP_MM
    bottom_margin = (
        server.LABEL_HEIGHT_MM
        - (server.LAYOUT["qr_top"] + server.LAYOUT["qr_size"] - server.CONTENT_UP_MM)
    )
    assert top_margin == pytest.approx(bottom_margin, abs=0.01)
    assert top_margin > 2.4


def test_larger_qr_keeps_the_right_edge_and_balanced_bottom_margin():
    assert server.LAYOUT["qr_size"] == pytest.approx(7.2)
    assert server.LAYOUT["qr_x"] + server.LAYOUT["qr_size"] == pytest.approx(22.2)
    svg = server.render_to_svg(ROWS[0])
    assert 'width="7.2" height="7.2" fill="#fff"' in svg


def test_technical_section_is_tightened_vertically_only():
    assert server.TECH_SECTION_WIDTH_MM == pytest.approx(22.67)
    assert server.TECH_SECTION_HEIGHT_MM == pytest.approx(21.27)
    assert server.LOWER_SEPARATOR_Y_MM == pytest.approx(43.32)
    assert 'preserveAspectRatio="none"' in server._svg_vector_art()


def test_fixed_size_serial_number_line_fits_the_print_font():
    text = "S/N: 5753120024723"
    for font in server.FONT_OPTIONS.values():
        width_mm = pdfmetrics.stringWidth(text, font.pdf_name, 4.78) / mm
        assert server.LAYOUT["info_x"] + width_mm <= server.LAYOUT["info_right"]


def test_each_registered_font_renders_svg_and_pdf():
    for font_key in server.FONT_OPTIONS:
        assert "dynamic-text" in server.render_to_svg(ROWS[0], font_key=font_key)
        assert server.render_to_pdf([ROWS[0]], font_key=font_key).getvalue().startswith(b"%PDF")


def test_unknown_print_font_is_rejected():
    with pytest.raises(ValueError, match="unavailable"):
        server.render_to_pdf([ROWS[0]], font_key="missing-font")


def test_numeric_excel_identifier_cells_are_rejected():
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["Device ID", "IMEI", "CCID", "QR Data"])
    sheet.append(["DEV-01", 356789012345678, 89148000001234567890, "QR-01"])
    buffer = io.BytesIO()
    workbook.save(buffer)
    with pytest.raises(ValueError, match="stored as Text"):
        server.read_device_file(buffer.getvalue(), "devices.xlsx")


def test_duplicate_identifiers_are_rejected():
    rows = [dict(ROWS[0]), dict(ROWS[1])]
    rows[1]["IMEI"] = rows[0]["IMEI"]
    errors = server.validate_rows(rows)
    assert any(error["field"] == "IMEI" and "Duplicate" in error["message"] for error in errors)


def test_pdf_has_one_physical_size_page_per_row():
    data = server.render_to_pdf(ROWS, printer_dpi=600).getvalue()
    reader = PdfReader(io.BytesIO(data))
    assert len(reader.pages) == 2
    width_mm = float(reader.pages[0].mediabox.width) / 72 * 25.4
    height_mm = float(reader.pages[0].mediabox.height) / 72 * 25.4
    assert width_mm == pytest.approx(server.LABEL_WIDTH_MM, abs=0.02)
    assert height_mm == pytest.approx(server.LABEL_HEIGHT_MM, abs=0.02)
    assert "K34632721" in (reader.pages[0].extract_text() or "")
    assert "K34632722" in (reader.pages[1].extract_text() or "")


def test_fixed_artwork_is_vector_not_an_embedded_image():
    data = server.render_to_pdf([ROWS[0]], printer_dpi=600).getvalue()
    page = PdfReader(io.BytesIO(data)).pages[0]
    xobjects = page.get("/Resources", {}).get("/XObject", {})
    assert all(obj.get_object().get("/Subtype") != "/Image" for obj in xobjects.values())


def test_a4_output_places_21_stickers_per_sheet():
    rows = [
        {
            "Device ID": f"K{i:08d}",
            "IMEI": str(356789012345000 + i),
            "CCID": str(89148000001234567000 + i),
            "QR Data": f"4679643367898{i:04d}",
            "Model": "4G Dongle",
        }
        for i in range(22)
    ]
    data = server.render_to_pdf(rows, printer_dpi=600, page_format="a4").getvalue()
    reader = PdfReader(io.BytesIO(data))
    assert len(reader.pages) == 2
    width_mm = float(reader.pages[0].mediabox.width) / 72 * 25.4
    height_mm = float(reader.pages[0].mediabox.height) / 72 * 25.4
    assert width_mm == pytest.approx(210, abs=0.02)
    assert height_mm == pytest.approx(297, abs=0.02)
    assert "K00000000" in (reader.pages[0].extract_text() or "")
    assert "K00000020" in (reader.pages[0].extract_text() or "")
    assert "K00000021" in (reader.pages[1].extract_text() or "")


def test_dense_ccid_is_rejected_at_300_dpi():
    with pytest.raises(ValueError, match="too dense"):
        server.render_to_pdf([ROWS[0]], printer_dpi=300)


def _find_pdftoppm() -> str | None:
    direct = shutil.which("pdftoppm")
    if direct:
        return direct
    bundled = Path.home() / ".cache/codex-runtimes/codex-primary-runtime/dependencies/native/poppler/Library/bin/pdftoppm.exe"
    return str(bundled) if bundled.exists() else None


def test_exported_codes_decode_at_600_dpi(tmp_path: Path):
    pdftoppm = _find_pdftoppm()
    if not pdftoppm:
        pytest.skip("pdftoppm is not installed")
    zxingcpp = pytest.importorskip("zxingcpp")
    pdf_path = tmp_path / "sticker.pdf"
    image_prefix = tmp_path / "sticker"
    pdf_path.write_bytes(server.render_to_pdf([ROWS[0]], printer_dpi=600).getvalue())
    subprocess.run(
        [pdftoppm, "-png", "-r", "600", "-f", "1", "-singlefile", str(pdf_path), str(image_prefix)],
        check=True,
        capture_output=True,
    )
    decoded = {result.text for result in zxingcpp.read_barcodes(Image.open(image_prefix.with_suffix(".png")))}
    assert ROWS[0]["CCID"] in decoded
    assert ROWS[0]["IMEI"] in decoded
    assert server._qr_payload(ROWS[0]) in decoded
    assert ROWS[0]["QR Data"] not in decoded
