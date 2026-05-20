# How To Add a Simple Web GUI for Supplier IČO Lookup

This guide adds a very simple web GUI to the existing Python scraper.

The GUI has one input field:

```text
[IČO] [Vyhľadať]
```

It lets the user enter an IČO such as:

```text
36785512
```

or:

```text
36 785 512
```

After search, the app shows a browser summary and creates downloadable:

- JSON
- CSV
- XLSX

---

## 1. Current project assumption

The current project already has:

```text
main.py
Dockerfile
docker-compose.yaml
requirements.txt
```

The existing `main.py` already contains a callable function:

```python
scrape_subject(ico: str) -> dict
```

The web GUI will import and call that function directly:

```python
from main import scrape_subject
```

The hardcoded value inside the existing `main()` function does not matter for the web GUI, because the GUI does not run `python main.py`. It runs FastAPI through Uvicorn.

---

## 2. Final project structure

After applying this guide, the folder should look like this:

```text
project/
  app.py
  main.py
  requirements.txt
  Dockerfile
  docker-compose.yaml
  cli.py                 # optional
  output/                # created automatically or mounted by Docker Compose
```

---

## 3. Create `app.py`

Create a new file named `app.py` in the same folder as `main.py`.

```python
import csv
import html
import json
import re
from datetime import datetime
from pathlib import Path
from typing import Any

from fastapi import FastAPI, Form, HTTPException
from fastapi.responses import FileResponse, HTMLResponse

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from main import scrape_subject


APP_TITLE = "Supplier IČO Lookup"
OUTPUT_DIR = Path("output")
OUTPUT_DIR.mkdir(exist_ok=True)

app = FastAPI(title=APP_TITLE)


# ============================================================
# INPUT VALIDATION
# ============================================================

def normalize_ico(value: str) -> str:
    """
    Accepts:
    - 36785512
    - 36 785 512
    - 36-785-512

    Returns:
    - 36785512

    Raises:
    - ValueError if the value is not exactly 8 digits after cleanup.
    """
    ico = re.sub(r"\D", "", value or "")

    if len(ico) != 8:
        raise ValueError("IČO musí mať presne 8 číslic.")

    return ico


# ============================================================
# RESULT NORMALIZATION
# ============================================================

def stringify(value: Any) -> str:
    """
    Converts nested dict/list values into readable strings for CSV/XLSX cells.
    """
    if value is None:
        return ""

    if isinstance(value, (dict, list)):
        return json.dumps(value, ensure_ascii=False)

    return str(value)


def find_first_value(source: dict, candidate_keys: list[str]) -> str:
    """
    Finds the first matching value from a dictionary by candidate key fragments.

    Useful because FinStat/RÚZ labels can vary slightly.
    """
    if not isinstance(source, dict):
        return ""

    lowered_candidates = [x.lower() for x in candidate_keys]

    for key, value in source.items():
        key_lower = str(key).lower()

        if any(candidate in key_lower for candidate in lowered_candidates):
            return stringify(value)

    return ""


def build_summary_row(result: dict) -> dict:
    """
    Creates one flat summary row for CSV/XLSX.
    Keeps this intentionally conservative: only uses values that are present.
    """
    orsr = result.get("orsr", {}) or {}
    rpvs = result.get("rpvs", {}) or {}
    finstat = result.get("finstat", {}) or {}
    ruz = result.get("ruz", {}) or {}

    finstat_basic = finstat.get("zakladne_udaje", {}) or {}
    finstat_financials = finstat.get("financne_ukazovatele", {}) or {}

    rpvs_partner = rpvs.get("partner_verejneho_sektora", {}) or {}
    rpvs_ubos = rpvs.get("konecni_uzivatelia_vyhod", []) or []

    return {
        "ico": result.get("ico", ""),
        "obchodne_meno": (
            orsr.get("obchodne_meno", "")
            or ruz.get("nazov", "")
            or rpvs_partner.get("Obchodné meno", "")
        ),
        "sidlo": orsr.get("sidlo", "") or finstat_basic.get("Sídlo", ""),
        "pravna_forma": orsr.get("pravna_forma", ""),
        "den_zapisu": orsr.get("den_zapisu", ""),
        "zakladne_imanie": orsr.get("vyska_zakladneho_imania", ""),
        "sk_nace": finstat_basic.get("SK NACE", "") or ruz.get("SK NACE", ""),
        "kategoria_zamestnancov": finstat_basic.get("Kategória zamestnancov", ""),
        "trzby_predaj_sluzieb": find_first_value(
            finstat_financials,
            ["tržby z predaja služieb", "trzby z predaja sluzieb", "tržby"],
        ),
        "vynosy": find_first_value(
            finstat_financials,
            ["výnosy", "vynosy"],
        ),
        "zisk_strata": find_first_value(
            finstat_financials,
            ["zisk", "strata", "výsledok hospodárenia", "vysledok hospodarenia"],
        ),
        "rpvs_pocet_kuv": len(rpvs_ubos),
        "orsr_pocet_statutarov": len(orsr.get("statutarny_organ", []) or []),
        "zdroj_orsr": "áno" if bool(orsr) else "nie",
        "zdroj_rpvs": "áno" if bool(rpvs) else "nie",
        "zdroj_finstat": "áno" if bool(finstat) else "nie",
        "zdroj_ruz": "áno" if bool(ruz) else "nie",
    }


def build_risk_flags(result: dict) -> list[dict]:
    """
    Simple initial risk rules.
    Keep these transparent and easy to audit.
    """
    flags = []

    orsr = result.get("orsr", {}) or {}
    rpvs = result.get("rpvs", {}) or {}
    finstat = result.get("finstat", {}) or {}
    ruz = result.get("ruz", {}) or {}

    if not orsr:
        flags.append({
            "severity": "yellow",
            "flag": "ORSR data missing",
            "detail": "ORSR did not return parsed data or scraping failed.",
        })

    if not finstat:
        flags.append({
            "severity": "yellow",
            "flag": "FinStat data missing",
            "detail": "FinStat did not return parsed data or scraping failed.",
        })

    if not ruz:
        flags.append({
            "severity": "yellow",
            "flag": "RÚZ data missing",
            "detail": "RÚZ did not return parsed data or scraping failed.",
        })

    if rpvs and not rpvs.get("konecni_uzivatelia_vyhod"):
        flags.append({
            "severity": "yellow",
            "flag": "No RPVS beneficial owners parsed",
            "detail": "RPVS was reached, but no KUV records were parsed.",
        })

    return flags


# ============================================================
# OUTPUT WRITERS
# ============================================================

def output_base_name(ico: str) -> str:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"{ico}_{timestamp}"


def write_json_output(result: dict, base_name: str) -> Path:
    path = OUTPUT_DIR / f"{base_name}.json"

    with path.open("w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    return path


def write_csv_output(summary_row: dict, base_name: str) -> Path:
    path = OUTPUT_DIR / f"{base_name}_summary.csv"

    with path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=list(summary_row.keys()))
        writer.writeheader()
        writer.writerow(summary_row)

    return path


def autosize_sheet_columns(ws) -> None:
    for column_cells in ws.columns:
        max_length = 0
        column_letter = get_column_letter(column_cells[0].column)

        for cell in column_cells:
            value = stringify(cell.value)
            max_length = max(max_length, len(value))

        ws.column_dimensions[column_letter].width = min(max_length + 2, 80)


def style_header(ws) -> None:
    fill = PatternFill("solid", fgColor="D9EAF7")

    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = fill
        cell.alignment = Alignment(vertical="top")


def add_key_value_sheet(wb: Workbook, title: str, data: dict) -> None:
    ws = wb.create_sheet(title=title[:31])
    ws.append(["key", "value"])

    if isinstance(data, dict):
        for key, value in data.items():
            ws.append([str(key), stringify(value)])
    else:
        ws.append(["value", stringify(data)])

    style_header(ws)
    autosize_sheet_columns(ws)


def write_xlsx_output(
    result: dict,
    summary_row: dict,
    risk_flags: list[dict],
    base_name: str,
) -> Path:
    path = OUTPUT_DIR / f"{base_name}_report.xlsx"

    wb = Workbook()

    # Summary sheet
    ws = wb.active
    ws.title = "Summary"
    ws.append(list(summary_row.keys()))
    ws.append([stringify(value) for value in summary_row.values()])
    style_header(ws)
    autosize_sheet_columns(ws)

    # Risk flags sheet
    ws_flags = wb.create_sheet(title="Risk flags")
    ws_flags.append(["severity", "flag", "detail"])

    for flag in risk_flags:
        ws_flags.append([
            flag.get("severity", ""),
            flag.get("flag", ""),
            flag.get("detail", ""),
        ])

    style_header(ws_flags)
    autosize_sheet_columns(ws_flags)

    # Source-specific sheets
    add_key_value_sheet(wb, "ORSR", result.get("orsr", {}) or {})
    add_key_value_sheet(wb, "RPVS", result.get("rpvs", {}) or {})
    add_key_value_sheet(wb, "FinStat", result.get("finstat", {}) or {})
    add_key_value_sheet(wb, "RUZ", result.get("ruz", {}) or {})

    # Raw JSON sheet
    ws_raw = wb.create_sheet(title="Raw JSON")
    ws_raw.append(["raw_json"])
    ws_raw.append([json.dumps(result, ensure_ascii=False, indent=2)])
    style_header(ws_raw)
    ws_raw.column_dimensions["A"].width = 120
    ws_raw["A2"].alignment = Alignment(wrap_text=True, vertical="top")

    wb.save(path)
    return path


def generate_outputs(result: dict) -> dict:
    ico = result.get("ico", "unknown")
    base_name = output_base_name(ico)

    summary_row = build_summary_row(result)
    risk_flags = build_risk_flags(result)

    json_path = write_json_output(result, base_name)
    csv_path = write_csv_output(summary_row, base_name)
    xlsx_path = write_xlsx_output(result, summary_row, risk_flags, base_name)

    return {
        "summary_row": summary_row,
        "risk_flags": risk_flags,
        "json_file": json_path.name,
        "csv_file": csv_path.name,
        "xlsx_file": xlsx_path.name,
    }


# ============================================================
# HTML RENDERING
# ============================================================

def render_page(title: str, body: str) -> HTMLResponse:
    return HTMLResponse(f"""
    <!doctype html>
    <html lang="sk">
    <head>
        <meta charset="utf-8">
        <title>{html.escape(title)}</title>
        <style>
            body {{
                font-family: Arial, sans-serif;
                max-width: 1100px;
                margin: 40px auto;
                padding: 0 20px;
                line-height: 1.45;
            }}
            input {{
                padding: 10px;
                width: 260px;
                font-size: 16px;
            }}
            button {{
                padding: 10px 16px;
                font-size: 16px;
                cursor: pointer;
            }}
            table {{
                border-collapse: collapse;
                width: 100%;
                margin-top: 20px;
            }}
            th, td {{
                border: 1px solid #ddd;
                padding: 8px;
                vertical-align: top;
                text-align: left;
            }}
            th {{
                background: #f2f2f2;
            }}
            pre {{
                background: #f7f7f7;
                padding: 16px;
                overflow-x: auto;
                border: 1px solid #ddd;
            }}
            .hint {{
                color: #666;
                margin-top: 8px;
            }}
            .downloads a {{
                display: inline-block;
                margin-right: 12px;
                margin-bottom: 12px;
            }}
            .error {{
                color: #9b1c1c;
                background: #fff0f0;
                padding: 12px;
                border: 1px solid #e0b4b4;
            }}
        </style>
    </head>
    <body>
        {body}
    </body>
    </html>
    """)


@app.get("/", response_class=HTMLResponse)
def home():
    return render_page(
        APP_TITLE,
        """
        <h1>Vyhľadanie dodávateľa podľa IČO</h1>

        <form method="post" action="/lookup">
            <input
                name="ico"
                placeholder="napr. 36 785 512"
                required
                autofocus
            >
            <button type="submit">Vyhľadať</button>
        </form>

        <div class="hint">
            Zadajte IČO s medzerami alebo bez medzier. Výstup bude dostupný ako JSON, CSV a XLSX.
        </div>
        """,
    )


@app.post("/lookup", response_class=HTMLResponse)
def lookup(ico: str = Form(...)):
    try:
        normalized_ico = normalize_ico(ico)

        result = scrape_subject(normalized_ico)
        outputs = generate_outputs(result)

        summary_row = outputs["summary_row"]
        risk_flags = outputs["risk_flags"]

        summary_rows_html = "\n".join(
            f"<tr><th>{html.escape(str(key))}</th><td>{html.escape(stringify(value))}</td></tr>"
            for key, value in summary_row.items()
        )

        if risk_flags:
            risk_rows_html = "\n".join(
                "<tr>"
                f"<td>{html.escape(flag.get('severity', ''))}</td>"
                f"<td>{html.escape(flag.get('flag', ''))}</td>"
                f"<td>{html.escape(flag.get('detail', ''))}</td>"
                "</tr>"
                for flag in risk_flags
            )
        else:
            risk_rows_html = """
            <tr>
                <td>green</td>
                <td>No basic risk flags</td>
                <td>No initial rule-based flags were generated.</td>
            </tr>
            """

        pretty_json = html.escape(json.dumps(result, ensure_ascii=False, indent=2))

        body = f"""
        <a href="/">← Nové vyhľadanie</a>

        <h1>Výsledok pre IČO {html.escape(normalized_ico)}</h1>

        <div class="downloads">
            <a href="/download/{html.escape(outputs["json_file"])}">Stiahnuť JSON</a>
            <a href="/download/{html.escape(outputs["csv_file"])}">Stiahnuť CSV</a>
            <a href="/download/{html.escape(outputs["xlsx_file"])}">Stiahnuť XLSX</a>
        </div>

        <h2>Súhrn</h2>
        <table>
            {summary_rows_html}
        </table>

        <h2>Risk flags</h2>
        <table>
            <tr>
                <th>severity</th>
                <th>flag</th>
                <th>detail</th>
            </tr>
            {risk_rows_html}
        </table>

        <h2>Raw JSON</h2>
        <pre>{pretty_json}</pre>
        """

        return render_page(f"Výsledok {normalized_ico}", body)

    except Exception as e:
        error_message = f"{type(e).__name__}: {str(e)}"

        body = f"""
        <a href="/">← Späť</a>
        <h1>Chyba</h1>
        <div class="error">{html.escape(error_message)}</div>
        """

        return render_page("Chyba", body)


@app.get("/download/{filename}")
def download_file(filename: str):
    """
    Safe download endpoint.
    Prevents path traversal by stripping directory parts from the filename.
    """
    safe_name = Path(filename).name
    path = OUTPUT_DIR / safe_name

    if not path.exists() or not path.is_file():
        raise HTTPException(status_code=404, detail="File not found")

    suffix = path.suffix.lower()

    if suffix == ".json":
        media_type = "application/json"
    elif suffix == ".csv":
        media_type = "text/csv"
    elif suffix == ".xlsx":
        media_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    else:
        media_type = "application/octet-stream"

    return FileResponse(
        path=path,
        filename=safe_name,
        media_type=media_type,
    )
```

---

## 4. Update `requirements.txt`

Replace or update `requirements.txt` so it includes at least:

```txt
beautifulsoup4
fastapi
lxml
openpyxl
python-multipart
requests
selenium
urllib3
uvicorn[standard]
```

---

## 5. Replace `Dockerfile`

Use this:

```dockerfile
FROM python:3.13-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

RUN mkdir -p /app/output

CMD ["uvicorn", "app:app", "--host", "0.0.0.0", "--port", "8000"]
```

---

## 6. Replace `docker-compose.yaml`

Use this:

```yaml
services:
  selenium:
    image: selenium/standalone-chrome:latest
    container_name: selenium-chrome
    shm_size: 2gb
    ports:
      - "4444:4444"
      - "7900:7900"
    healthcheck:
      test: ["CMD-SHELL", "wget -q -O - http://localhost:4444/wd/hub/status || wget -q -O - http://localhost:4444/status || exit 1"]
      interval: 5s
      timeout: 5s
      retries: 20
      start_period: 10s

  scraper:
    build: .
    container_name: py-scraper
    depends_on:
      selenium:
        condition: service_healthy
    environment:
      SELENIUM_URL: http://selenium:4444/wd/hub
    ports:
      - "8000:8000"
    volumes:
      - ./output:/app/output
```

---

## 7. Optional: add `cli.py`

This is optional. It provides command-line input in addition to the web GUI.

```python
import argparse
import json
import re
from pathlib import Path

from main import scrape_subject


def normalize_ico(value: str) -> str:
    ico = re.sub(r"\D", "", value or "")

    if len(ico) != 8:
        raise ValueError("IČO musí mať presne 8 číslic.")

    return ico


def main():
    parser = argparse.ArgumentParser(description="Supplier lookup by IČO")
    parser.add_argument("--ico", required=True, help="IČO, napr. 36785512 alebo '36 785 512'")
    parser.add_argument("--out", default="output", help="Output folder")

    args = parser.parse_args()

    ico = normalize_ico(args.ico)
    output_dir = Path(args.out)
    output_dir.mkdir(exist_ok=True)

    result = scrape_subject(ico)

    output_path = output_dir / f"{ico}_result.json"

    with output_path.open("w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    print(json.dumps(result, ensure_ascii=False, indent=2))
    print(f"\nVýsledok uložený do: {output_path}")


if __name__ == "__main__":
    main()
```

Run:

```bash
docker compose run --rm scraper python cli.py --ico "36 785 512"
```

---

## 8. Start the web GUI

Run:

```bash
docker compose up --build
```

Open:

```text
http://localhost:8000
```

Expected page:

```text
Vyhľadanie dodávateľa podľa IČO

[ napr. 36 785 512 ] [ Vyhľadať ]
```

---

## 9. Output files

Generated files are saved to:

```text
output/
```

Example:

```text
output/
  36785512_20260520_143012.json
  36785512_20260520_143012_summary.csv
  36785512_20260520_143012_report.xlsx
```

The browser page also provides direct download links for:

- JSON
- CSV
- XLSX

---

## 10. Output design

Use all three formats.

| Output | Purpose |
|---|---|
| JSON | full raw evidence/debug output |
| CSV | one-row summary for simple import/filtering |
| XLSX | practical business report with multiple sheets |

CSV alone is not enough because supplier data contains one-to-many structures:

- statutory persons
- beneficial owners
- shareholders
- contracts
- accounting statements

---

## 11. Current implemented sources

The current scraper collects from:

| Source | Purpose |
|---|---|
| ORSR | commercial register details, statutory body, shareholders, legal facts |
| RPVS | public-sector partner data and beneficial owners |
| FinStat | financial and company summary data |
| RÚZ | accounting entity data |

---

## 12. Recommended future public registries

Add these later as separate adapters.

| Source | Purpose | Priority |
|---|---|---:|
| RPO — Register právnických osôb | canonical identity, address, legal form, status | High |
| CRZ — Centrálny register zmlúv | supplier contracts, contract values, public counterparties | High |
| Insolvency / liquidation register | insolvency, restructuring, liquidation red flags | High |
| Obchodný vestník | older legal notices and events | Medium |
| Financial Administration / tax-debtor datasets | tax/debt indicators where available | Medium |
| ÚVO / UVOstat | public procurement exposure | Medium |
| EU sanctions list | sanctions screening for company names and persons | Medium |

---

## 13. What not to add in MVP

Avoid these in the MVP:

- court case mining
- general web search
- social media checks
- automated negative-news screening
- OCR-heavy PDF processing
- complex ownership graphing

These add noise and false positives before the source adapters are stable.

---

## 14. Known limitations

This implementation is intentionally simple.

Limitations:

1. The request blocks until scraping finishes.
2. If one source fails inside `scrape_subject()`, the whole lookup may fail.
3. Selenium scraping is fragile compared with official APIs.
4. CRZ, RPO, insolvency, tax-debt, ÚVO and sanctions are not yet implemented.
5. No user authentication is included.

---

## 15. Recommended next technical improvement

Refactor `scrape_subject()` so one failed source does not break the whole lookup.

Target structure:

```python
{
    "ico": "36785512",
    "sources": {
        "orsr": {"status": "ok", "data": {}},
        "rpvs": {"status": "ok", "data": {}},
        "finstat": {"status": "error", "error": "..."},
        "ruz": {"status": "ok", "data": {}}
    }
}
```

This should be done before adding CRZ, RPO, insolvency, tax-debt, ÚVO or sanctions adapters.

---

## 16. Minimal validation checklist

After implementing the files:

1. Run:

```bash
docker compose up --build
```

2. Open:

```text
http://localhost:8000
```

3. Enter:

```text
36 785 512
```

4. Confirm that:

```text
output/
```

contains:

```text
.json
_summary.csv
_report.xlsx
```

5. Confirm that invalid input such as:

```text
123
```

returns an IČO validation error.

---

## 17. Production hardening checklist

Before real use:

- isolate failures per source
- add timeouts per source
- add source URLs to output
- add scrape timestamps
- add versioned output schema
- add audit log
- add retries with backoff
- replace Selenium with official APIs where possible
- add CRZ source adapter
- add RPO source adapter
- add insolvency/liquidation adapter
- add basic authentication if exposed outside localhost
```
