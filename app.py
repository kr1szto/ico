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
    mode = result.get("mode", "full")

    def source_error_detail(source_key: str, fallback: str) -> str:
        error = result.get(f"{source_key}_error") or {}

        if isinstance(error, dict):
            error_type = error.get("type", "Error")
            message = error.get("message", fallback)
            current_url = error.get("current_url")

            if current_url:
                return f"{error_type}: {message} URL: {current_url}"

            return f"{error_type}: {message}"

        return fallback

    if mode != "fast" and not orsr:
        flags.append({
            "severity": "yellow",
            "flag": "ORSR data missing",
            "detail": source_error_detail(
                "orsr",
                "ORSR did not return parsed data or scraping failed.",
            ),
        })

    if not finstat:
        flags.append({
            "severity": "yellow",
            "flag": "FinStat data missing",
            "detail": source_error_detail(
                "finstat",
                "FinStat did not return parsed data or scraping failed.",
            ),
        })

    if mode != "fast" and not ruz:
        flags.append({
            "severity": "yellow",
            "flag": "RÚZ data missing",
            "detail": source_error_detail(
                "ruz",
                "RÚZ did not return parsed data or scraping failed.",
            ),
        })

    if mode != "fast" and not rpvs:
        flags.append({
            "severity": "yellow",
            "flag": "RPVS data missing",
            "detail": source_error_detail(
                "rpvs",
                "RPVS did not return parsed data or scraping failed.",
            ),
        })

    selenium_error = result.get("selenium_error")
    if isinstance(selenium_error, dict):
        flags.append({
            "severity": "red",
            "flag": "Selenium setup failed",
            "detail": (
                f"{selenium_error.get('type', 'Error')}: "
                f"{selenium_error.get('message', 'Selenium could not complete browser automation.')}"
            ),
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
            label {{
                display: inline-flex;
                align-items: center;
                gap: 6px;
                margin: 14px 18px 0 0;
                color: #333;
            }}
            input[type="radio"] {{
                width: auto;
            }}
            button {{
                padding: 10px 16px;
                font-size: 16px;
                cursor: pointer;
            }}
            button:disabled {{
                cursor: wait;
                opacity: 0.7;
            }}
            .actions {{
                margin-top: 14px;
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
            .status {{
                color: #4b5563;
                display: none;
                margin-top: 14px;
            }}
            .summary-note {{
                background: #eef6ff;
                border: 1px solid #bfdbfe;
                color: #1e3a8a;
                padding: 12px;
                margin: 16px 0;
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
        <script>
            function markSubmitting(form) {{
                const button = form.querySelector("button[type='submit']");
                const status = document.getElementById("lookup-status");
                const mode = form.querySelector("input[name='mode']:checked").value;

                if (button) {{
                    button.disabled = true;
                    button.textContent = mode === "fast" ? "Vyhľadávam..." : "Kontrolujem registre...";
                }}

                if (status) {{
                    status.style.display = "block";
                    status.textContent = mode === "fast"
                        ? "Rýchle vyhľadanie trvá zvyčajne pár sekúnd."
                        : "Úplná kontrola môže trvať dlhšie, pretože používa viac verejných registrov.";
                }}
            }}
        </script>
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

        <form method="post" action="/lookup" onsubmit="markSubmitting(this)">
            <input
                name="ico"
                placeholder="napr. 36 785 512"
                required
                autofocus
            >

            <div>
                <label>
                    <input type="radio" name="mode" value="fast" checked>
                    Rýchlo
                </label>
                <label>
                    <input type="radio" name="mode" value="full">
                    Úplná kontrola
                </label>
            </div>

            <div class="actions">
                <button type="submit">Vyhľadať</button>
            </div>
        </form>

        <div class="hint">
            Rýchle vyhľadanie používa FinStat. Úplná kontrola pridá ORSR, RPVS a RÚZ a môže trvať dlhšie.
        </div>
        <div id="lookup-status" class="status"></div>
        """,
    )


@app.get("/health")
def health():
    return {"status": "ok"}


@app.post("/lookup", response_class=HTMLResponse)
def lookup(ico: str = Form(...), mode: str = Form("fast")):
    try:
        normalized_ico = normalize_ico(ico)
        lookup_mode = "full" if mode == "full" else "fast"

        result = scrape_subject(
            normalized_ico,
            include_deep_sources=(lookup_mode == "full"),
        )
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
        mode_note = (
            "Rýchly výsledok používa FinStat. Pre ORSR, RPVS a RÚZ spustite úplnú kontrolu."
            if lookup_mode == "fast"
            else "Úplná kontrola zahŕňa FinStat, ORSR, RPVS a RÚZ podľa dostupnosti registrov."
        )

        body = f"""
        <a href="/">← Nové vyhľadanie</a>

        <h1>Výsledok pre IČO {html.escape(normalized_ico)}</h1>

        <div class="summary-note">{html.escape(mode_note)}</div>

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
