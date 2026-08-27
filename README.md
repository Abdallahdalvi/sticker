# Bulk Sticker Generator

Local web application for producing scanner-ready Reliance 4G dongle labels from Excel or CSV data.

The supported label is **24.08 mm × 74.08 mm portrait**, matching the supplied reference PDF. Each accepted spreadsheet row becomes one physical-size sticker with:

- dynamic Device ID text
- dynamic CCID Code 128 and readable text
- dynamic IMEI Code 128 and readable text
- a dynamic QR code containing Device ID, CCID, and IMEI

## Run

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python server.py
```

Open <http://127.0.0.1:8502/>.

The bind address and port can be overridden with `APP_HOST` and `APP_PORT`.

The older Streamlit prototype in `app.py` is retained only for historical reference. It is not the supported application.

## CasaOS deployment on port 8877

Copy this project folder to the CasaOS server, for example under
`/DATA/AppData/bulk-sticker-generator`, then run:

```bash
cd /DATA/AppData/bulk-sticker-generator
docker compose -f docker-compose.casaos.yml up -d --build
```

Open `http://CASAOS_SERVER_IP:8877/`. The Compose service:

- publishes host port **8877** to container port **8877**
- restarts automatically unless explicitly stopped
- runs as a non-root user
- includes an `/api/health` Docker health check
- supports both `amd64` and `arm64` CasaOS hosts

To update after copying a newer project version:

```bash
docker compose -f docker-compose.casaos.yml up -d --build
```

To view status or logs:

```bash
docker compose -f docker-compose.casaos.yml ps
docker compose -f docker-compose.casaos.yml logs -f --tail=100
```

## Input workbook

Download `device_data_template.xlsx` from the application. Required columns are Device ID, IMEI, and CCID. Existing workbooks may keep the legacy QR Data column, but its value is not encoded.

| Column | Rule |
| --- | --- |
| Device ID | required, maximum 32 characters, unique |
| IMEI | required, exactly 15 digits, unique |
| CCID | required, 18–22 digits, unique |
| QR Data | optional legacy column; accepted but not encoded |
| Model | optional; defaults to `4G Dongle` |

Identifier cells in `.xlsx` files must be formatted as **Text**. Numeric or formula cells are rejected instead of being silently rounded, converted to scientific notation, or stripped of leading zeros. UTF-8 CSV is also supported. Legacy `.xls` is intentionally rejected.

## Printing

- Use a **600 DPI** thermal printer for 18–22 digit CCID labels at this narrow width.
- Print at **100% / Actual Size**.
- Disable “Fit to page”, scaling, and browser headers/footers.
- **Individual sticker pages** contain one sticker per physical-size page for roll/thermal printing.
- **A4 sticker sheets** contain 21 stickers per page in a centered 7 × 3 grid with 4 mm cutting gutters.

The app calculates the Code 128 narrow-module width before export. A 300 DPI export is rejected when the selected values cannot maintain two printer dots per narrow module.

## Validation

The complete upload is validated before preview or export. PDF generation stops for missing fields, invalid lengths, duplicates, unsafe numeric Excel cells, invalid record ranges, and barcodes that are too dense for the selected printer DPI.

## Tests

```powershell
python -m pip install -r requirements-dev.txt
pytest -q
```

The integration scan test renders the generated PDF with Poppler and decodes both Code 128 symbols and the QR code with ZXing. If `pdftoppm` is not on `PATH`, that single scan test is skipped.

## Rebuilding the vector reference artwork

The checked-in raster is retained only as the source for the traced vector contours. Generated PDFs and live previews use `static/reference_vector_paths.json`, so the logo, specifications, icons, dividers, and lion remain sharp at any zoom.

```powershell
python scripts/build_reference_asset.py "C:\path\to\reference.pdf" static/reference_static.png
python scripts/build_vector_art.py static/reference_static.png static/reference_vector_paths.json
```
