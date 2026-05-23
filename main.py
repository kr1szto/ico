import os
import sys
import time
import json
import traceback
import re
import requests
import urllib3

from bs4 import BeautifulSoup

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

ORSR_URL = "https://www.orsr.sk/search_ico.asp"
RPVS_URL = "https://rpvs.gov.sk/rpvs"
RUZ_URL = "https://www.registeruz.sk/cruz-public/domain/accountingentity/simplesearch"

SELENIUM_URL = os.getenv("SELENIUM_URL", "").strip()
CHROME_BIN = os.getenv("CHROME_BIN", "/usr/bin/chromium")
CHROMEDRIVER_PATH = os.getenv("CHROMEDRIVER_PATH", "/usr/bin/chromedriver")
SCRAPER_WAIT_SECONDS = int(os.getenv("SCRAPER_WAIT_SECONDS", "45"))
PAGE_LOAD_TIMEOUT_SECONDS = int(os.getenv("PAGE_LOAD_TIMEOUT_SECONDS", "60"))


# ============================================================
# HELPERS
# ============================================================

def build_chrome_options(use_local_binary: bool) -> Options:
    options = Options()

    # workaround na privacy error
    options.set_capability("acceptInsecureCerts", True)
    options.add_argument("--ignore-certificate-errors")
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")

    if use_local_binary and CHROME_BIN:
        options.binary_location = CHROME_BIN

    return options


def create_remote_driver(options: Options, max_attempts: int, delay: int):
    for attempt in range(1, max_attempts + 1):
        try:
            print(f"[INFO] Pokus {attempt}/{max_attempts} o pripojenie na Selenium: {SELENIUM_URL}")
            driver = webdriver.Remote(
                command_executor=SELENIUM_URL,
                options=options
            )
            configure_driver_timeouts(driver)
            print("[INFO] Selenium session vytvorená úspešne.")
            return driver
        except Exception as e:
            print(f"[WARN] Selenium ešte nie je ready: {e}")
            if attempt == max_attempts:
                raise
            time.sleep(delay)


def create_local_driver(options: Options):
    print(f"[INFO] Spúšťam lokálny Chrome driver: {CHROMEDRIVER_PATH}")
    service = Service(executable_path=CHROMEDRIVER_PATH)
    driver = webdriver.Chrome(service=service, options=options)
    configure_driver_timeouts(driver)
    print("[INFO] Lokálna Selenium session vytvorená úspešne.")
    return driver


def configure_driver_timeouts(driver) -> None:
    driver.set_page_load_timeout(PAGE_LOAD_TIMEOUT_SECONDS)
    driver.set_script_timeout(PAGE_LOAD_TIMEOUT_SECONDS)


def create_driver(max_attempts=10, delay=3):
    use_local_driver = not SELENIUM_URL
    options = build_chrome_options(use_local_binary=use_local_driver)

    if use_local_driver:
        return create_local_driver(options)

    return create_remote_driver(options, max_attempts, delay)


def save_debug(driver, prefix="debug"):
    try:
        with open(f"{prefix}.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        driver.save_screenshot(f"{prefix}.png")
        print(f"[INFO] Uložené {prefix}.html a {prefix}.png")
    except Exception as e:
        print("[WARN] Nepodarilo sa uložiť debug artefakty:", e)


def format_source_error(error: Exception, driver=None) -> dict:
    if type(error).__name__ == "TimeoutException":
        message = "Timed out waiting for the expected registry page element."
    else:
        message = str(error).strip()

        if not message:
            message = "Timed out waiting for the expected registry page element."

    result = {
        "type": type(error).__name__,
        "message": message,
    }

    if driver:
        try:
            result["current_url"] = driver.current_url
        except Exception:
            pass

    return result


def run_source(subjekt: dict, source_key: str, label: str, callback, driver=None, ico: str | None = None) -> None:
    try:
        subjekt[source_key] = callback()
    except Exception as error:
        print(f"[ERROR] {label} zlyhal.")
        traceback.print_exc()
        subjekt[f"{source_key}_error"] = format_source_error(error, driver=driver)

        if driver and ico:
            save_debug(driver, prefix=f"{ico}_{source_key}_debug")


def normalize_text(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def compact_statutory_person(lines: list[str]) -> list[str]:
    if not lines:
        return []

    # ak je to len typ orgánu alebo príliš krátky záznam, vráť ako je
    if len(lines) < 4:
        return lines

    result = []

    # 1. Meno + funkcia
    first_line = " ".join(lines[:4]).strip()
    result.append(normalize_text(first_line))

    # 2. Ulica + číslo
    if len(lines) >= 6:
        address_1 = f"{lines[4]} {lines[5]}"
        result.append(normalize_text(address_1))

    # 3. Mesto + PSČ
    if len(lines) >= 8:
        address_2 = f"{lines[6]} {lines[7]}"
        result.append(normalize_text(address_2))

    # 4. Vznik funkcie + (od: ...)
    if len(lines) >= 10:
        function_line = f"{lines[8]} {lines[9]}"
        result.append(normalize_text(function_line))
    elif len(lines) >= 9:
        result.append(normalize_text(lines[8]))

    # ak by bolo riadkov viac než 10, pridaj zvyšok
    if len(lines) > 10:
        for extra in lines[10:]:
            result.append(normalize_text(extra))

    return result

def parse_orsr_section_by_label(soup: BeautifulSoup, section_label: str) -> list:
    """
    Nájde ORSR sekciu podľa labelu v span.tl, napr.:
    - Štatutárny orgán
    - Dozorná rada
    - Akcie
    - Akcionár
    - Ďalšie právne skutočnosti

    Vráti zoznam blokov (každý blok = list riadkov).
    """
    label = None
    for span in soup.select("span.tl"):
        text = span.get_text(" ", strip=True)
        if section_label in text:
            label = span
            break

    if not label:
        return []

    row = label.find_parent("tr")
    if not row:
        return []

    cells = row.find_all("td", recursive=False)
    if len(cells) < 2:
        return []

    content_cell = cells[1]

    # ak sú vnorené tabuľky, ber ich po blokoch
    nested_tables = content_cell.find_all("table", recursive=False)

    results = []

    if nested_tables:
        for tbl in nested_tables:
            text = tbl.get_text("\n", strip=True)
            lines = [normalize_text(line) for line in text.splitlines() if normalize_text(line)]
            if lines:
                results.append(lines)
    else:
        # fallback: sekcia môže byť aj len textovo bez nested tables
        text = content_cell.get_text("\n", strip=True)
        lines = [normalize_text(line) for line in text.splitlines() if normalize_text(line)]
        if lines:
            results.append(lines)

    return results


def compact_raw_section(section_rows: list[list[str]]) -> list[list[str]]:
    cleaned_rows = []

    for row in section_rows:
        cleaned = [
            normalize_text(item.replace("\r", " ").replace("\n", " "))
            for item in row
            if normalize_text(item.replace("\r", " ").replace("\n", " "))
        ]
        cleaned_rows.append(cleaned)

    return cleaned_rows

def clean_value(text: str) -> str:
    if not text:
        return text

    # odstráň copy hlášku
    text = text.replace("Údaj bol úspešne skopírovaný", "")

    # odstráň whitespace znaky
    text = text.replace("\n", " ").replace("\t", " ")

    # zjednoť medzery
    text = " ".join(text.split())

    return text.strip()



# ============================================================
# ORSR
# ============================================================

def find_orsr_ico_input(driver):
    driver.switch_to.default_content()

    elems = driver.find_elements(By.NAME, "ICO")
    if elems:
        return elems[0]

    frames = driver.find_elements(By.CSS_SELECTOR, "frame, iframe")

    for frame in frames:
        try:
            driver.switch_to.default_content()
            driver.switch_to.frame(frame)
            elems = driver.find_elements(By.NAME, "ICO")
            if elems:
                return elems[0]
        except Exception:
            continue

    driver.switch_to.default_content()
    return None


def orsr_search_company(driver, wait, ico: str):
    print("[INFO] ORSR: otváram stránku...")
    driver.get(ORSR_URL)
    wait.until(lambda d: d.execute_script("return document.readyState") == "complete")

    search_input = find_orsr_ico_input(driver)
    if not search_input:
        raise RuntimeError("ORSR: Nenašiel som input[name='ICO'].")

    search_input.clear()
    search_input.send_keys(ico)

    search_button = wait.until(
        EC.element_to_be_clickable((
            By.XPATH,
            "//input[@type='submit' and contains(normalize-space(@value), 'Hľadaj')]"
        ))
    )

    print("[INFO] ORSR: klikám na Hľadaj...")
    driver.execute_script("arguments[0].click();", search_button)
    wait.until(lambda d: d.execute_script("return document.readyState") == "complete")


def orsr_open_first_result(driver, wait):
    first_result_xpath = "//tbody/tr[td][1]/td[2]//a[contains(@href, 'vypis.asp')]"

    print("[INFO] ORSR: čakám na prvý výsledok...")
    first_result = wait.until(
        EC.element_to_be_clickable((By.XPATH, first_result_xpath))
    )

    print("[INFO] ORSR: klikám na prvý výsledok...")
    driver.execute_script("arguments[0].click();", first_result)

    wait.until(lambda d: "vypis.asp" in d.current_url)
    wait.until(lambda d: d.execute_script("return document.readyState") == "complete")


def parse_orsr_basic_info(soup: BeautifulSoup) -> dict:
    result = {
        "obchodne_meno": None,
        "sidlo": None,
        "den_zapisu": None,
        "pravna_forma": None,
        "predmet_podnikania": [],        
        "vyska_zakladneho_imania": None,
        "datum_aktualizacie_dat": None

    }

    field_map = {
        "Obchodné meno": "obchodne_meno",
        "Sídlo": "sidlo",
        "Deň zápisu": "den_zapisu",
        "Právna forma": "pravna_forma",
        "Predmet podnikania": "predmet_podnikania",
        "Výška základného imania": "vyska_zakladneho_imania",
        "Dátum aktualizácie dát": "datum_aktualizacie_dat"
    }

    for row in soup.select("tr"):
        cells = row.find_all("td", recursive=False)
        if len(cells) != 2:
            continue

        label = cells[0].get_text(" ", strip=True).replace(":", "")
        value_cell = cells[1]

        if label not in field_map:
            continue

        key = field_map[label]

        if key == "predmet_podnikania":
            items = [normalize_text(x) for x in value_cell.stripped_strings if normalize_text(x)]
            result["predmet_podnikania"].extend(items)
        else:
            result[key] = normalize_text(value_cell.get_text(" ", strip=True))

    return result


def parse_orsr_statutarny_organ(soup: BeautifulSoup) -> dict:
    result = {
        "typ_organu": None,
        "statutarny_organ": []
    }

    label = None
    for span in soup.select("span.tl"):
        text = span.get_text(" ", strip=True)
        if "Štatutárny orgán" in text:
            label = span
            break

    if not label:
        return result

    row = label.find_parent("tr")
    if not row:
        return result

    cells = row.find_all("td", recursive=False)
    if len(cells) < 2:
        return result

    content_cell = cells[1]
    nested_tables = content_cell.find_all("table", recursive=False)

    for idx, tbl in enumerate(nested_tables, start=1):
        text = tbl.get_text("\n", strip=True)
        lines = [normalize_text(line) for line in text.splitlines() if normalize_text(line)]

        if not lines:
            continue

        if idx == 1:
            result["typ_organu"] = " ".join(lines)
        else:
            compact_lines = compact_statutory_person(lines)
            result["statutarny_organ"].append(compact_lines)
        

    return result


def parse_orsr_detail(driver) -> dict:
    print("[INFO] ORSR: parsujem detail firmy...")
    html = driver.page_source
    soup = BeautifulSoup(html, "lxml")

    result = {}
    result.update(parse_orsr_basic_info(soup))

    # Štatutárny orgán (špeciálne spracovanie)
    statutar = parse_orsr_statutarny_organ(soup)
    result.update(statutar)

    # Dozorná rada
    dozorna_rada_raw = parse_orsr_section_by_label(soup, "Dozorná rada")
    result["dozorna_rada"] = [
        compact_statutory_person(lines) if len(lines) >= 4 else lines
        for lines in dozorna_rada_raw
    ]

    # Akcie
    result["akcie"] = parse_orsr_section_by_label(soup, "Akcie")

    # Akcionár
    akcionar_raw = parse_orsr_section_by_label(soup, "Akcionár")
    result["akcionar"] = compact_raw_section(akcionar_raw)

    # Ďalšie právne skutočnosti
    result["dalsie_pravne_skutocnosti"] = parse_orsr_section_by_label(
        soup, "Ďalšie právne skutočnosti"
    )

    return result


# ============================================================
# RPVS
# ============================================================

def rpvs_search_company(driver, wait, ico: str):
    print("[INFO] RPVS: otváram stránku...")
    driver.get(RPVS_URL)

    wait.until(lambda d: d.execute_script("return document.readyState") == "complete")
    time.sleep(2)

    rpvs_input = wait.until(
        EC.presence_of_element_located((By.ID, "partner_hladat_text"))
    )

    rpvs_input.clear()
    rpvs_input.send_keys(ico)
    rpvs_input.send_keys(Keys.ENTER)

    wait.until(
        EC.presence_of_element_located((By.ID, "table-VyhladavaniePartnera"))
    )
    time.sleep(2)


def rpvs_open_first_result(driver, wait):
    first_result_xpath = "//table[@id='table-VyhladavaniePartnera']//tbody/tr[1]/td[4]//a"

    print("[INFO] RPVS: klikám na prvý výsledok...")
    first_result = wait.until(
        EC.element_to_be_clickable((By.XPATH, first_result_xpath))
    )

    driver.execute_script("arguments[0].click();", first_result)

    wait.until(lambda d: "/Detail/" in d.current_url)
    wait.until(lambda d: d.execute_script("return document.readyState") == "complete")


def parse_rpvs_detail(driver) -> dict:
    print("[INFO] RPVS: parsujem detail...")
    html = driver.page_source
    soup = BeautifulSoup(html, "lxml")

    result = {
        "partner_verejneho_sektora": {},
        "opravnena_osoba": {},
        "konecni_uzivatelia_vyhod": []
    }

    for panel in soup.select("div.panel.panel-default"):
        title_el = panel.select_one("h2.panel-title")
        if not title_el:
            continue

        title = normalize_text(title_el.get_text(strip=True))

        # ========================================================
        # PARTNER VEREJNÉHO SEKTORA
        # ========================================================
        if title == "Partner verejného sektora":
            body = panel.select_one("div.panel-body")
            if not body:
                continue

            for group in body.select("div.form-group"):
                label_el = group.select_one("label")
                value_el = group.select_one("p.form-control-static")

                if not label_el or not value_el:
                    continue

                key = normalize_text(label_el.get_text(" ", strip=True))
                value = normalize_text(value_el.get_text(" ", strip=True))

                result["partner_verejneho_sektora"][key] = value

        # ========================================================
        # OPRÁVNENÁ OSOBA
        # ========================================================
        elif title == "Oprávnená osoba":
            body = panel.select_one("div.panel-body")
            if not body:
                continue

            for group in body.select("div.form-group"):
                label_el = group.select_one("label")
                value_el = group.select_one("p.form-control-static")

                if not label_el or not value_el:
                    continue

                key = normalize_text(label_el.get_text(" ", strip=True))
                value = normalize_text(value_el.get_text(" ", strip=True))

                result["opravnena_osoba"][key] = value

        # ========================================================
        # KONEČNÍ UŽÍVATELIA VÝHOD
        # ========================================================
        elif title == "Koneční užívatelia výhod":
            table = panel.select_one("table")
            if not table:
                continue

            rows = table.select("tbody tr")

            for row in rows:
                cells = row.select("th.hidden-xs, td.hidden-xs")

                if len(cells) < 5:
                    continue

                kuvy = {
                    "Meno a priezvisko": normalize_text(cells[0].get_text(strip=True)),
                    "Dátum narodenia": normalize_text(cells[1].get_text(strip=True)),
                    "Štátna príslušnosť": normalize_text(cells[2].get_text(strip=True)),
                    "Adresa": normalize_text(cells[3].get_text(" ", strip=True)),
                    "Verejný funkcionár": normalize_text(cells[4].get_text(strip=True)),
                }

                result["konecni_uzivatelia_vyhod"].append(kuvy)

    return result

# ============================================================
# FINSTAT
# ============================================================

def finstat_scrape(input_ico: str) -> dict:
    print("[INFO] FinStat: scraping...")
    url = f"https://finstat.sk/vyhladavanie?query={input_ico}"

    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/124.0.0.0 Safari/537.36"
        )
    }

    result = {
        "zakladne_udaje": {},
        "financne_ukazovatele": {}
    }

    response = requests.get(url, headers=headers, timeout=15, verify=False)
    response.raise_for_status()

    soup = BeautifulSoup(response.text, "lxml")

    # Základné info
    element = soup.select_one("div.col-md-8.detail-company-info-side.col-xs-12")
    if element:
        for lead in element.select("div.lead"):
            lead.decompose()

        labels = [
            "IČO",
            "DIČ",
            "IČ DPH",
            "Sídlo",
            "Historický názov",
            "Dátum vzniku",
            "SK NACE",
            "Kategória zamestnancov"
        ]

        pattern = r'(?=\b(?:' + "|".join(re.escape(label) for label in labels) + r')\b)'

        for child in element.find_all(recursive=False):
            text = child.get_text(" ", strip=True)
            if not text:
                continue

            parts = re.split(pattern, text)

            for part in parts:
                part = normalize_text(part)
                if not part:
                    continue

                for label in labels:
                    if part.startswith(label):
                        value = part[len(label):].strip()
                        result["zakladne_udaje"][label] = value
                        break

    # Finančné ukazovatele
    table = soup.select_one("table.table.table-lined.table-condensed.detail-company-financial")
    if table:
        tbody = table.select_one("tbody")

        if tbody:
            for row in tbody.find_all("tr"):
                cells = row.find_all(["th", "td"])

                if len(cells) >= 2:
                    name = normalize_text(cells[0].get_text(" ", strip=True))
                    value = normalize_text(cells[1].get_text(" ", strip=True))

                    if name:
                        result["financne_ukazovatele"][name] = value

    return result
# ============================================================
# RUZ
# ============================================================
def ruz_search_company(driver, wait, ico: str):
    print("[INFO] RUZ: otváram stránku...")
    driver.get(RUZ_URL)

    wait.until(lambda d: d.execute_script("return document.readyState") == "complete")
    time.sleep(2)

    print("[INFO] RUZ: hľadám input...")

    search_input = wait.until(
        EC.presence_of_element_located((By.ID, "input_search"))
    )

    search_input.clear()
    search_input.send_keys(ico)

    # počkaj na dropdown suggestions
    print("[INFO] RUZ: čakám na autocomplete výsledky...")

    first_result = wait.until(
        EC.element_to_be_clickable((
            By.XPATH,
            "//ul[contains(@class,'ui-autocomplete')]//li[1]//a"
        ))
    )

    print("[INFO] RUZ: klikám prvý výsledok...")
    driver.execute_script("arguments[0].click();", first_result)

    # počkaj na detail
    wait.until(lambda d: "/show/" in d.current_url)
    wait.until(lambda d: d.execute_script("return document.readyState") == "complete")

    print("[SUCCESS] RUZ detail otvorený.")





# ============================================================
# RUZ PARSE
# ============================================================

def parse_ruz_detail(driver) -> dict:
    print("[INFO] RUZ: parsujem detail...")

    html = driver.page_source
    soup = BeautifulSoup(html, "lxml")

    result = {}

    container = soup.select_one('div[data-tab-container="1"]')
    if not container:
        print("[WARN] RUZ: nenašiel sa hlavný container")
        return result
        
    # odstráni UI prvky (napr. copy icony)
    for icon in container.select(".icon-duplicate, .btn-link, i.icon"):
        icon.extract()
    
    # názov
    title = container.select_one("h3, h1")
    if title:
        result["nazov"] = normalize_text(title.get_text())

    # info bloky
    for block in container.select("div.b-content div.fs-14.text-gray"):

        # iba text uzlov (bez vnorených UI vecí)
        texts = list(block.stripped_strings)

        if len(texts) < 2:
            continue

        label = normalize_text(texts[0].replace(":", ""))

        # špeciálne prípady
        if label == "SK NACE":
            value = clean_value(" ".join(block.stripped_strings))

        elif label == "Adresa":
            value = clean_value(", ".join([normalize_text(x) for x in block.stripped_strings]))


        else:
            value = clean_value(normalize_text(" ".join(texts[1:])))

        result[label] = value

    return result

# ============================================================
# ORCHESTRATION
# ============================================================

def normalize_selected_sources(
    selected_sources: list[str] | None = None,
    include_deep_sources: bool | None = None,
) -> set[str]:
    if selected_sources is None:
        if include_deep_sources is False:
            return {"finstat"}
        return {"finstat", "orsr", "rpvs", "ruz"}

    allowed_sources = {"finstat", "orsr", "rpvs", "ruz"}
    normalized_sources = {
        source.strip().lower()
        for source in selected_sources
        if isinstance(source, str)
    }

    return normalized_sources & allowed_sources


def scrape_subject(
    ico: str,
    selected_sources: list[str] | None = None,
    include_deep_sources: bool | None = None,
) -> dict:
    sources = normalize_selected_sources(
        selected_sources=selected_sources,
        include_deep_sources=include_deep_sources,
    )

    subjekt = {
        "ico": ico,
        "orsr": {},
        "rpvs": {},
        "finstat": {},
        "ruz": {},
        "selected_sources": sorted(sources),
    }

    driver = None

    try:
        if "finstat" in sources:
            run_source(
                subjekt,
                "finstat",
                "FinStat",
                lambda: finstat_scrape(ico),
                ico=ico,
            )

        browser_sources = sources & {"orsr", "rpvs", "ruz"}
        if not browser_sources:
            return subjekt

        driver = create_driver()
        wait = WebDriverWait(driver, SCRAPER_WAIT_SECONDS)

        def scrape_orsr():
            orsr_search_company(driver, wait, ico)
            orsr_open_first_result(driver, wait)
            return parse_orsr_detail(driver)

        def scrape_rpvs():
            rpvs_search_company(driver, wait, ico)
            rpvs_open_first_result(driver, wait)
            return parse_rpvs_detail(driver)

        def scrape_ruz():
            ruz_search_company(driver, wait, ico)
            return parse_ruz_detail(driver)

        if "orsr" in browser_sources:
            run_source(subjekt, "orsr", "ORSR", scrape_orsr, driver=driver, ico=ico)
        if "rpvs" in browser_sources:
            run_source(subjekt, "rpvs", "RPVS", scrape_rpvs, driver=driver, ico=ico)
        if "ruz" in browser_sources:
            run_source(subjekt, "ruz", "RÚZ", scrape_ruz, driver=driver, ico=ico)

        return subjekt

    except Exception as error:
        print("[ERROR] Selenium orchestration zlyhala.")
        traceback.print_exc()
        subjekt["selenium_error"] = format_source_error(error, driver=driver)
        return subjekt

    finally:
        if driver:
            driver.quit()


def main():
    ico = "12345678"

    try:
        subjekt = scrape_subject(ico)

        print("\n=== VÝSLEDNÝ SUBJEKT ===\n")
        print(json.dumps(subjekt, ensure_ascii=False, indent=2))

        with open(f"{ico}_result.json", "w", encoding="utf-8") as f:
            json.dump(subjekt, f, ensure_ascii=False, indent=2)

        print(f"[INFO] Výsledok uložený do {ico}_result.json")

    except Exception as e:
        print("[FATAL] Program zlyhal.")
        print(type(e).__name__, str(e))
        sys.exit(1)

if __name__ == "__main__":
    main()
