from pathlib import Path
import csv
import glob
import json
import os
import re
import shutil
import time
import unicodedata

import pandas as pd
import pytesseract
from pdf2image import convert_from_path
from PIL import ImageFilter, ImageOps
try:
    from pypdf import PdfReader
except ImportError:
    from PyPDF2 import PdfReader

import mysql.connector

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException, NoAlertPresentException
from selenium.webdriver.firefox.options import Options


# ═══════════════════════════════════════════════════════════════════════════
#  CONFIGURATION
# ═══════════════════════════════════════════════════════════════════════════

download_folder = r"C:\Users\Admin\Downloads"

# OCR
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
POPPLER_PATH = r"C:\Program Files\poppler-26.02.0\Library\bin"
DPI = 250

# Static selection targets
STATIC_STATE_NAME = "Meghalaya"
STATIC_YEAR = "2025"
STATIC_ROLL_TYPE = "Final Roll - 2025"

# MySQL
DB_CONFIG = {
    "host": "localhost",
    "user": "root",
    "password": "root",
    "database": "election_db",
}

VOTER_COLUMNS = [
    "Voterids", "Name", "GuardianType", "GuardianName", "House Number", "Age", "Gender",
]
METADATA_COLUMNS = [
    "MainTown/Village", "PostOffice", "PoliceStation", "Mandal",
    "Sub Division", "District", "PinCode", "AssemblyConstituency",
    "ParliamentaryConstituency", "PollingStation", "PollingAddress",
]


# ═══════════════════════════════════════════════════════════════════════════
#  SECTION 1 — OCR / CSV PIPELINE  (Read Part Number happens on the site;
#  this section turns a downloaded PDF into a CSV of voter records)
# ═══════════════════════════════════════════════════════════════════════════

def clean(value):
    """Normalise OCR text and discard characters that cannot belong to a field."""
    value = unicodedata.normalize("NFKC", value or "")
    value = "".join(
        char for char in value
        if char.isalnum() or char.isspace() or char in " .,'()&/#-"
    )
    return re.sub(r"\s+", " ", value).strip(" :-")


def csv_house_number(value):
    """Protect numeric slash/hyphen house numbers from Excel date conversion."""
    value = clean(value)
    if re.fullmatch(r"\d{1,4}\s*[/-]\s*\d{1,4}", value):
        return f'="{value}"'
    return value


def ocr(image, psm=6, scale=1.5):
    """OCR a card after a portable contrast/noise cleanup."""
    gray = ImageOps.grayscale(image)
    if scale != 1:
        gray = gray.resize((int(gray.width * scale), int(gray.height * scale)))
    gray = ImageOps.autocontrast(gray).filter(ImageFilter.MedianFilter(3))
    try:
        return pytesseract.image_to_string(gray, lang="eng", config=f"--psm {psm}")
    finally:
        gray.close()


def value_after_label(text, labels):
    for label in labels:
        match = re.search(rf"{label}\s*[:\-]?\s*([^\n]+)", text, re.I)
        if match:
            return clean(match.group(1))
    return ""


def multiline_value(text, label, stop_labels):
    stop = "|".join(stop_labels)
    match = re.search(
        rf"{label}\s*[:\-]?\s*(.*?)(?=\s*(?:{stop})\s*[:\-]?|$)",
        text,
        re.I | re.S,
    )
    return clean(match.group(1)) if match else ""


def polling_station_value(text):
    labels = [
        r"No\.?\s*(?:and|&)\s*Name\s*of\s*Part",
        r"No\.?\s*(?:and|&)\s*Name\s*of\s*Polling\s*(?:Station|Booth)",
        r"Name\s*(?:and\s*address\s*)?of\s*Polling\s*(?:Station|Booth)",
    ]
    stop_labels = (
        r"Address\s*of\s*Polling\s*(?:Station|Booth)"
        r"|Type\s*of\s*Polling\s*(?:Station|Booth)"
        r"|Number\s*of\s*Auxiliary"
        r"|4\.\s*NUMBER\s*OF\s*ELECTORS"
    )
    for label in labels:
        m = re.search(
            rf"{label}\s*(?:[:\-]\s*)?(.*?)(?=\s*(?:{stop_labels})\s*[:\-]?|$)",
            text,
            re.I | re.S,
        )
        if m:
            value = clean(m.group(1))
            if value:
                return value
    return ""


def parse_cover_page(image):
    text = ocr(image, psm=6, scale=1)
    metadata = {
        "MainTown/Village": value_after_label(text, [r"Main Town\s*(?:or|/)\s*Village"]),
        "PostOffice": value_after_label(text, [r"Post Office"]),
        "PoliceStation": value_after_label(text, [r"Police Station"]),
        "Mandal": value_after_label(text, [r"Block", r"Mandal", r"Tehsil", r"Tehsil\s*/\s*Mandal", r"Panchayat", r"Patwari", r"Taluk", r"Taluka", r"Circle", r"Sub\s*District", r"Revenue\s*Circle"]),
        "Sub Division": value_after_label(text, [r"Subdivision", r"Revenue Division"]),
        "District": value_after_label(text, [r"District"]),
        "PinCode": value_after_label(text, [r"Pin\s*code", r"Pincode", r"Pin", r"PinCode", r"Postal\s*Code"]),
        "AssemblyConstituency": "",
        "ParliamentaryConstituency": "",
        "PollingStation": polling_station_value(text),
        "PollingAddress": multiline_value(
            text,
            r"Address of Polling Station",
            [r"Type of Polling Station", r"Number of Auxiliary", r"4\.\s*NUMBER OF ELECTORS"],
        ),
    }

    if not metadata["PollingStation"]:
        metadata["PollingStation"] = polling_station_value(ocr(image, psm=11, scale=1))

    assembly = re.search(
        r"No\.\s*Name\s*and\s*Reservation\s*Status\s*of\s*Assembly\s*Constituency\s*:\s*(.+?)(?=\n|Part\s*No)",
        text,
        re.I,
    )
    metadata["AssemblyConstituency"] = clean(assembly.group(1)) if assembly else ""

    parliament = re.search(r"located\s*:\s*([^\n]+)", text, re.I)
    if parliament:
        metadata["ParliamentaryConstituency"] = clean(parliament.group(1))
    else:
        parliament = re.search(
            r"No\.\s*Name\s*and\s*Reservation\s*Status\s*of\s*Parliamentary\s*Constituency\s*:\s*(.+?)(?=\n|1\.\s*Details)",
            text,
            re.I | re.S,
        )
        metadata["ParliamentaryConstituency"] = clean(parliament.group(1)) if parliament else ""

    metadata["PollingStation"] = re.sub(
        r"\b(?:Type of Polling Station|Address of Polling Station)\b.*", "",
        metadata["PollingStation"], flags=re.I,
    ).strip(" :-")
    metadata["PollingAddress"] = re.sub(
        r"\b(?:Type of Polling Station|Number of Auxiliary|4\.\s*NUMBER OF ELECTORS)\b.*",
        "", metadata["PollingAddress"], flags=re.I,
    ).strip(" :-")
    metadata["PollingAddress"] = re.sub(
        r"^Stations? in this part\s*:\s*", "", metadata["PollingAddress"], flags=re.I
    )
    return metadata


def guardian_values(text):
    relationships = {
        "Father": r"Father'?s?\s*Name",
        "Mother": r"Mother'?s?\s*Name",
        "Husband": r"Husband'?s?\s*Name",
        "Wife": r"Wife'?s?\s*Name",
        "Other": r"Other'?s?\s*Name",
    }
    for relationship, label in relationships.items():
        value = value_after_label(text, [label])
        if value:
            return relationship, value
    name = value_after_label(text, [r"Name"])
    return "Unknown", name


def epic_match(text):
    compact = re.sub(r"[^A-Za-z0-9]", "", text or "").upper()
    for match in re.finditer(r"[A-Z0-9]{3}[A-Z0-9]{7}", compact):
        candidate = match.group(0)
        prefix = candidate[:3].translate(str.maketrans({"0": "O", "1": "I", "5": "S", "8": "B"}))
        number = candidate[3:].translate(str.maketrans({"O": "0", "Q": "0", "D": "0", "I": "1", "L": "1", "Z": "2", "S": "5", "B": "8"}))
        if prefix.isalpha() and number.isdigit():
            return prefix + number
    return ""


def parse_voter_card(text):
    text = text.replace("\x0c", " ")
    text = re.sub(r"Photo\s*(?:is\s*)?Available", "", text, flags=re.I)
    text = re.sub(r"\b(?:Narne|Nane)\b", "Name", text, flags=re.I)

    voter_id_match = epic_match(text)
    name = value_after_label(text, [r"Name"])
    age_match = re.search(r"\bA(?:g|q|c)e?\s*[:\-]?\s*(\d{1,3})\b", text, re.I)
    age = age_match.group(1) if age_match else ""
    gender = value_after_label(text, [r"Gender"])
    house_number = value_after_label(text, [r"House\s*(?:No\.?|Number)"])
    house_number = re.sub(r"Photo.*", "", house_number, flags=re.I)
    house_number = re.sub(r"Available.*", "", house_number, flags=re.I).strip()

    if re.fullmatch(r"\d{2}-\d{4}", house_number):
        house_number = ""

    guardian_type, guardian_name = guardian_values(text)

    if not voter_id_match or not name:
        return None
    return {
        "Voterids": voter_id_match,
        "Name": name,
        "GuardianType": guardian_type,
        "GuardianName": guardian_name,
        "House Number": csv_house_number(house_number),
        "Age": age,
        "Gender": clean(gender),
    }


def voter_card_boxes(image):
    width, height = image.size
    left, right = 0.022 * width, 0.985 * width
    top, bottom = 0.034 * height, 0.972 * height
    column_width = (right - left) / 3
    row_height = (bottom - top) / 10
    for row in range(10):
        for column in range(3):
            yield image.crop((
                left + column * column_width + 2,
                top + row * row_height + 2,
                left + (column + 1) * column_width - 2,
                top + (row + 1) * row_height - 2,
            ))


def valid_age(value):
    try:
        return 0 <= int(value) <= 120
    except (TypeError, ValueError):
        return False


def elector_age(value):
    try:
        return 18 <= int(value) <= 120
    except (TypeError, ValueError):
        return False


def merge_card_records(sparse, block):
    if not sparse:
        return block
    if not block:
        return sparse

    merged = sparse.copy()
    for field in ("Name", "House Number", "Gender"):
        if not merged.get(field) and block.get(field):
            merged[field] = block[field]
    if merged.get("GuardianType") == "Unknown" and block.get("GuardianType") != "Unknown":
        merged["GuardianType"] = block["GuardianType"]
        merged["GuardianName"] = block["GuardianName"]
    elif not merged.get("GuardianName") and block.get("GuardianName"):
        merged["GuardianName"] = block["GuardianName"]
    if not elector_age(merged.get("Age")) and elector_age(block.get("Age")):
        merged["Age"] = block["Age"]
    return merged


def card_needs_block_ocr(record):
    return not record or (
        record.get("GuardianType") == "Unknown"
        or not record.get("GuardianName")
        or not elector_age(record.get("Age"))
        or not record.get("Gender")
    )


def age_from_card_crop(card):
    width, height = card.size
    crop = card.crop((0.02 * width, 0.48 * height, 0.78 * width, 0.82 * height))
    try:
        texts = [ocr(crop, psm=6, scale=2), ocr(crop, psm=11, scale=2), ocr(crop, psm=7, scale=2)]
    finally:
        crop.close()
    for text in texts:
        match = re.search(r"\bA(?:g|q|c)e?\s*[:\-]?\s*(\d{1,3})\b", text, re.I)
        value = match.group(1) if match else ""
        if valid_age(value):
            return value
    return ""


def identity_from_card_crop(card):
    width, height = card.size
    crop = card.crop((0.01 * width, 0.02 * height, 0.98 * width, 0.42 * height))
    try:
        return [ocr(crop, psm=6, scale=2), ocr(crop, psm=11, scale=2)]
    finally:
        crop.close()


def guardian_from_card_crop(card):
    width, height = card.size
    crop = card.crop((0.02 * width, 0.20 * height, 0.80 * width, 0.58 * height))
    try:
        text = ocr(crop, psm=6, scale=2)
    finally:
        crop.close()
    text = re.sub(r"\b(?:Narne|Nane)\b", "Name", text, flags=re.I)
    return guardian_values(text)


def card_record(card):
    try:
        sparse = parse_voter_card(ocr(card, psm=11))
        if not card_needs_block_ocr(sparse):
            return sparse
        block = parse_voter_card(ocr(card, psm=6))
        record = merge_card_records(sparse, block)
        if not record:
            for text in identity_from_card_crop(card):
                record = parse_voter_card(text)
                if record:
                    break
        if record and record.get("GuardianType") == "Unknown":
            guardian_type, guardian_name = guardian_from_card_crop(card)
            if guardian_type != "Unknown":
                record["GuardianType"] = guardian_type
                record["GuardianName"] = guardian_name
        if record and not elector_age(record.get("Age")):
            retry_age = age_from_card_crop(card)
            if retry_age:
                record["Age"] = retry_age
        return record
    except (OSError, pytesseract.TesseractError) as error:
        print(f"Skipped one unreadable card: {error}")
        return None


from PIL import ImageOps

def card_has_print(card):
    thumbnail = ImageOps.grayscale(card).resize((60, 80))
    try:
        pixels = thumbnail.getchannel(0).getdata()
        return sum(pixel < 180 for pixel in pixels) > 70
    finally:
        thumbnail.close()

def render_page(file_path, page_number):
    return convert_from_path(
        str(file_path), dpi=DPI, fmt="jpeg", poppler_path=POPPLER_PATH,
        first_page=page_number, last_page=page_number,
        thread_count=1,
    )[0]


def expected_elector_count(cover_text):
    match = re.search(r"\bTotal\s*\n?\s*(\d{1,5})\b", cover_text, re.I)
    return int(match.group(1)) if match else None


def process_pdf_file(file_path):
    """OCR -> Create CSV. Returns the CSV path.

    The CSV is written via file_path.with_suffix('.csv'), so it is created
    directly next to the PDF that is passed in — no separate input/output
    staging folder is used.
    """
    file_path = Path(file_path)
    pdf = PdfReader(str(file_path))
    total_pages = len(pdf.pages)
    if total_pages < 3:
        raise ValueError("Expected electoral roll with at least three pages")

    print(f"Processing {file_path.name} ({total_pages} pages)")
    cover = render_page(file_path, 1)
    try:
        cover_text = ocr(cover, psm=6, scale=1)
        metadata = parse_cover_page(cover)
        expected_count = expected_elector_count(cover_text)
    finally:
        cover.close()

    records = []
    unresolved_cards = []

    for page_number in range(3, total_pages + 1):
        image = render_page(file_path, page_number)
        try:
            for card_number, card in enumerate(voter_card_boxes(image), start=1):
                try:
                    record = card_record(card)
                    if record:
                        records.append(record)
                        if not elector_age(record.get("Age")):
                            unresolved_cards.append({
                                "PDF Page": page_number,
                                "Card Position": card_number,
                                "Issue": "Age needs review",
                                "OCR Age": record.get("Age", ""),
                            })
                    elif card_has_print(card):
                        unresolved_cards.append({
                            "PDF Page": page_number,
                            "Card Position": card_number,
                            "Issue": "Voter card not recognised",
                            "OCR Age": "",
                        })
                finally:
                    card.close()
        finally:
            image.close()
        print(f"Completed page {page_number}: {len(records)} elector records so far")

    if not records:
        raise RuntimeError("No elector cards were recognised; check Tesseract and the PDF layout")

    raw_df = pd.DataFrame(records, columns=VOTER_COLUMNS)
    duplicate_rows = raw_df[raw_df.duplicated("Voterids", keep=False)]
    df = raw_df.drop_duplicates("Voterids")
    for column in METADATA_COLUMNS:
        df[column] = metadata.get(column, "")

    csv_path = file_path.with_suffix(".csv")
    df.to_csv(csv_path, index=False, encoding="utf-8-sig", quoting=csv.QUOTE_ALL)
    print(f"Saved {len(df)} records to: {csv_path}")

    if unresolved_cards:
        review_path = file_path.with_name(file_path.stem + "_unresolved_cards.csv")
        pd.DataFrame(unresolved_cards).to_csv(review_path, index=False, encoding="utf-8-sig")
        print(f"WARNING: {len(unresolved_cards)} printed cards were not recognised. Review: {review_path}")
    if not duplicate_rows.empty:
        duplicate_path = file_path.with_name(file_path.stem + "_duplicate_epics.csv")
        duplicate_rows.to_csv(duplicate_path, index=False, encoding="utf-8-sig", quoting=csv.QUOTE_ALL)
        print(f"WARNING: duplicate EPIC reads were kept for review: {duplicate_path}")
    if expected_count and len(df) < expected_count:
        print(
            f"WARNING: cover page reports {expected_count} electors but {len(df)} unique "
            "EPICs were extracted. Review this PDF before treating it as complete."
        )
    return csv_path


# ═══════════════════════════════════════════════════════════════════════════
#  SECTION 2 — MYSQL  (Create Database / Create Table on first run,
#  then Insert Records into MySQL from each CSV)
# ═══════════════════════════════════════════════════════════════════════════

_db_initialized = False


def get_mysql_connection():
    """Connect to the server, creating the target database if it's missing."""
    server_config = {k: v for k, v in DB_CONFIG.items() if k != "database"}
    server_conn = mysql.connector.connect(**server_config)
    cursor = server_conn.cursor()
    cursor.execute(f"CREATE DATABASE IF NOT EXISTS `{DB_CONFIG['database']}`")
    server_conn.commit()
    cursor.close()
    server_conn.close()

    return mysql.connector.connect(**DB_CONFIG)


def ensure_table_exists(conn):
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS voters (
            Voterids VARCHAR(20) PRIMARY KEY,
            Name VARCHAR(255),
            GuardianType VARCHAR(20),
            GuardianName VARCHAR(255),
            HouseNumber VARCHAR(100),
            Age VARCHAR(10),
            Gender VARCHAR(20),
            MainTownVillage VARCHAR(255),
            PostOffice VARCHAR(255),
            PoliceStation VARCHAR(255),
            Mandal VARCHAR(255),
            SubDivision VARCHAR(255),
            District VARCHAR(255),
            PinCode VARCHAR(20),
            AssemblyConstituency VARCHAR(255),
            ParliamentaryConstituency VARCHAR(255),
            PollingStation VARCHAR(255),
            PollingAddress TEXT,
            SourceFile VARCHAR(255)
        )
    """)
    conn.commit()
    cursor.close()


def insert_csv_to_mysql(csv_path):
    """Read CSV -> Create Database/Table (first run only) -> Insert Records."""
    global _db_initialized

    conn = get_mysql_connection()
    try:
        if not _db_initialized:
            ensure_table_exists(conn)
            _db_initialized = True

        df = pd.read_csv(csv_path, dtype=str).fillna("")
        cursor = conn.cursor()

        insert_sql = """
            INSERT INTO voters (
                Voterids, Name, GuardianType, GuardianName, HouseNumber, Age, Gender,
                MainTownVillage, PostOffice, PoliceStation, Mandal, SubDivision, District,
                PinCode, AssemblyConstituency, ParliamentaryConstituency, PollingStation,
                PollingAddress, SourceFile
            ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            ON DUPLICATE KEY UPDATE
                Name = VALUES(Name),
                GuardianType = VALUES(GuardianType),
                GuardianName = VALUES(GuardianName),
                HouseNumber = VALUES(HouseNumber),
                Age = VALUES(Age),
                Gender = VALUES(Gender)
        """

        rows = []
        for _, row in df.iterrows():
            # Strip the ="..."  Excel-date-protection wrapper before storing.
            house_number = str(row.get("House Number", "")).strip()
            if house_number.startswith('="') and house_number.endswith('"'):
                house_number = house_number[2:-1]

            rows.append((
                row.get("Voterids", ""),
                row.get("Name", ""),
                row.get("GuardianType", ""),
                row.get("GuardianName", ""),
                house_number,
                row.get("Age", ""),
                row.get("Gender", ""),
                row.get("MainTown/Village", ""),
                row.get("PostOffice", ""),
                row.get("PoliceStation", ""),
                row.get("Mandal", ""),
                row.get("Sub Division", ""),
                row.get("District", ""),
                row.get("PinCode", ""),
                row.get("AssemblyConstituency", ""),
                row.get("ParliamentaryConstituency", ""),
                row.get("PollingStation", ""),
                row.get("PollingAddress", ""),
                Path(csv_path).name,
            ))

        cursor.executemany(insert_sql, rows)
        conn.commit()
        print(f"  ✔ Inserted/updated {len(rows)} records into MySQL from {Path(csv_path).name}")
        cursor.close()
    finally:
        conn.close()


# ═══════════════════════════════════════════════════════════════════════════
#  SECTION 3 — SELENIUM SCRAPER  (Open Website -> ... -> Download PDF ->
#  Rename PDF -> Create Folder -> Move PDF -> OCR/CSV/MySQL per PDF)
# ═══════════════════════════════════════════════════════════════════════════

options = Options()
options.set_preference("browser.download.folderList", 2)
options.set_preference("browser.download.dir", download_folder)
options.set_preference(
    "browser.helperApps.neverAsk.saveToDisk",
    "application/pdf,application/octet-stream"
)
options.set_preference("pdfjs.disabled", True)

driver = webdriver.Firefox(options=options)
driver.maximize_window()

wait = WebDriverWait(driver, 30)

driver.get("https://www.eci.gov.in/")

try:
    wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, "ngx-spinner-overlay")))
except Exception:
    pass

download_link = wait.until(
    EC.presence_of_element_located((By.XPATH, "//a[contains(text(),'Download Electoral Roll')]"))
)
driver.execute_script("arguments[0].click();", download_link)

time.sleep(1)


def accept_ok_alert_once(timeout=8):
    end_time = time.time() + timeout
    while time.time() < end_time:
        try:
            alert = driver.switch_to.alert
            alert_text = alert.text
            alert.accept()
            print(f"Alert accepted: {alert_text}")
            return True
        except Exception:
            time.sleep(0.3)
    print("No alert appeared within timeout - continuing.")
    return False


accept_ok_alert_once()
time.sleep(1.5)

if len(driver.window_handles) > 1:
    driver.switch_to.window(driver.window_handles[-1])

state_dropdown_el = None
for attempt in range(3):
    try:
        state_dropdown_el = wait.until(EC.presence_of_element_located((By.NAME, "stateCode")))
        break
    except TimeoutException:
        print(f"  ⚠ Timed out waiting for state dropdown (attempt {attempt + 1}). Retrying...")
        try:
            driver.switch_to.alert.accept()
        except Exception:
            pass
        time.sleep(1)

if state_dropdown_el is None:
    raise TimeoutException("Could not locate state dropdown after retries.")


def make_safe(name):
    return "".join(c for c in name if c not in r'<>:"/\|?*').strip()


def clean_option_text(raw_text):
    raw_text = raw_text.strip()
    if " - " in raw_text:
        return raw_text.split(" - ", 1)[1].strip()
    return raw_text


REACT_SELECT_OPTION_XPATH = (
    "//div[contains(@id,'react-select') and contains(@id,'option')]"
    " | //div[@role='option']"
    " | //div[contains(@class,'option') and not(contains(@class,'control'))]"
)
CONSTITUENCY_INPUT_XPATHS = [
    "//label[contains(normalize-space(.),'Assembly Constituency')]/following::input[contains(@id,'react-select') and @role='combobox'][1]",
    "//*[contains(normalize-space(.),'Assembly Constituency')]/following::input[contains(@id,'react-select') and @role='combobox'][1]",
    "//input[contains(@id,'react-select') and @role='combobox']",
]
CONSTITUENCY_CONTROL_XPATHS = [
    "//label[contains(normalize-space(.),'Assembly Constituency')]/following::div[contains(@class,'control')][1]",
    "//*[contains(normalize-space(.),'Assembly Constituency')]/following::div[contains(@class,'control')][1]",
    "//div[contains(@class,'control') and .//input[contains(@id,'react-select') and @role='combobox']]",
]


def clear_blocking_alert():
    try:
        alert = driver.switch_to.alert
        print(f"  Alert closed while switching dropdowns: {alert.text}")
        alert.accept()
        time.sleep(0.5)
        return True
    except NoAlertPresentException:
        return False
    except Exception:
        return False


def wait_for_page_ready(timeout=20):
    try:
        WebDriverWait(driver, timeout).until(
            EC.invisibility_of_element_located((By.CLASS_NAME, "ngx-spinner-overlay"))
        )
    except Exception:
        pass
    clear_blocking_alert()


def is_usable_element(element):
    try:
        return element.is_displayed() and element.is_enabled()
    except Exception:
        return False


def visible_options():
    opts = driver.find_elements(By.XPATH, REACT_SELECT_OPTION_XPATH)
    visible = []
    seen = set()
    for option in opts:
        try:
            text = option.text.strip()
            key = (text, option.location.get("x"), option.location.get("y"))
            if option.is_displayed() and text and key not in seen:
                visible.append(option)
                seen.add(key)
        except Exception:
            pass
    return visible


def click_constituency_target(target):
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", target)
    time.sleep(0.2)

    try:
        control = target.find_element(By.XPATH, "./ancestor-or-self::div[contains(@class,'control')][1]")
    except Exception:
        control = target

    click_candidates = [control, target]
    try:
        indicator = control.find_element(
            By.XPATH,
            ".//*[contains(@class,'indicator') or contains(@class,'Indicators') or contains(@class,'dropdown')]"
        )
        click_candidates.insert(0, indicator)
    except Exception:
        pass

    for candidate in click_candidates:
        try:
            driver.execute_script("arguments[0].click();", candidate)
            time.sleep(0.7)
            opts = visible_options()
            if opts:
                return opts
        except Exception:
            pass

    for key in (None, Keys.ENTER):
        try:
            actions = ActionChains(driver).move_to_element(target).click()
            if key:
                actions.send_keys(key)
            actions.perform()
            time.sleep(0.8)
            opts = visible_options()
            if opts:
                return opts
        except Exception:
            pass

    return visible_options()


def find_constituency_targets():
    inputs = []
    controls = []

    for xpath in CONSTITUENCY_INPUT_XPATHS:
        try:
            for element in driver.find_elements(By.XPATH, xpath):
                if is_usable_element(element) and element not in inputs:
                    inputs.append(element)
        except Exception:
            pass

    for input_el in inputs:
        try:
            control = input_el.find_element(By.XPATH, "./ancestor::div[contains(@class,'control')][1]")
            if is_usable_element(control) and control not in controls:
                controls.append(control)
        except Exception:
            pass

    for xpath in CONSTITUENCY_CONTROL_XPATHS:
        try:
            for element in driver.find_elements(By.XPATH, xpath):
                if is_usable_element(element) and element not in controls:
                    controls.append(element)
        except Exception:
            pass

    return inputs, controls


def open_constituency_options(timeout=35):
    last_error = None
    end_time = time.time() + timeout

    while time.time() < end_time:
        try:
            wait_for_page_ready(timeout=5)
            driver.execute_script("window.scrollTo({top: 0, behavior: 'instant'});")
            time.sleep(0.5)

            already_open = visible_options()
            if already_open:
                return already_open

            inputs, controls = find_constituency_targets()
            print(f"  Assembly Constituency targets found: {len(inputs)} inputs, {len(controls)} controls")

            click_targets = controls + inputs
            if not click_targets:
                last_error = "No visible Assembly Constituency React-Select input/control found"
                time.sleep(0.8)
                continue

            for target in click_targets:
                try:
                    opts = click_constituency_target(target)
                    if opts:
                        print(f"  Assembly Constituency options opened: {len(opts)}")
                        return opts
                except Exception as e:
                    last_error = e
                    clear_blocking_alert()

            time.sleep(0.8)
        except Exception as e:
            last_error = e
            clear_blocking_alert()
            time.sleep(0.8)

    raise TimeoutException(f"Could not open Assembly Constituency options: {last_error}")


def get_constituency_option_texts():
    selected_text = get_selected_constituency_text()
    opts = open_constituency_options()
    option_texts = [option.text.strip() for option in opts if option.text.strip()]

    try:
        ActionChains(driver).send_keys(Keys.ESCAPE).perform()
    except Exception:
        pass

    if selected_text and selected_text.casefold() not in {t.casefold() for t in option_texts}:
        option_texts = [selected_text] + option_texts

    print(f"  Constituency processing plan ({len(option_texts)}): {option_texts}")
    return option_texts


def get_selected_constituency_text():
    _, controls = find_constituency_targets()
    for control in controls:
        try:
            values = control.find_elements(
                By.XPATH, ".//*[contains(@class,'singleValue') or contains(@class,'single-value')]"
            )
            for value in values:
                text = value.text.strip()
                if text:
                    return text

            control_text = control.get_attribute("innerText") or control.text or ""
            for line in control_text.splitlines():
                candidate = line.strip()
                if re.match(r"^\d+\s*-\s*.+$", candidate):
                    return candidate
        except Exception:
            pass
    return ""


def select_constituency_by_text(constituency_text):
    current_text = get_selected_constituency_text()
    if current_text and current_text.strip().casefold() == constituency_text.strip().casefold():
        print(f"  Constituency already selected: {current_text}")
        return current_text

    opts = open_constituency_options()
    wanted = constituency_text.strip().casefold()

    for const_option in opts:
        raw_text = const_option.text.strip()
        if raw_text.casefold() == wanted:
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", const_option)
            time.sleep(0.3)
            driver.execute_script("arguments[0].click();", const_option)
            wait_for_page_ready()
            return raw_text

    available = ", ".join(option.text.strip() for option in opts if option.text.strip())
    raise ValueError(
        f"Constituency '{constituency_text}' is not currently available. Visible options: {available}"
    )


def select_constituency_with_retry(constituency_text, attempts=3):
    last_error = None
    for attempt in range(1, attempts + 1):
        try:
            return select_constituency_by_text(constituency_text)
        except Exception as error:
            last_error = error
            print(f"  Could not select '{constituency_text}' (attempt {attempt}/{attempts}): {error}")
            try:
                ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                driver.execute_script("if (document.activeElement) document.activeElement.blur();")
            except Exception:
                pass
            wait_for_page_ready(timeout=10)
            time.sleep(2)

    raise TimeoutException(
        f"Could not select constituency '{constituency_text}' after {attempts} attempts: {last_error}"
    )


def get_select_options(select_name):
    dropdown = Select(wait.until(EC.presence_of_element_located((By.NAME, select_name))))
    return [
        (option.get_attribute("value"), option.text.strip())
        for option in dropdown.options
        if option.get_attribute("value") and "--Select--" not in option.text
    ]


def select_dropdown_value(select_name, value):
    dropdown = Select(wait.until(EC.presence_of_element_located((By.NAME, select_name))))
    dropdown.select_by_value(value)
    return dropdown.first_selected_option.text.strip()


def select_year_and_roll_type():
    year_dropdown = wait.until(EC.presence_of_element_located((By.NAME, "revyear")))
    Select(year_dropdown).select_by_visible_text(STATIC_YEAR)
    time.sleep(1.5)
    wait_for_page_ready()

    roll_dropdown = wait.until(EC.presence_of_element_located((By.NAME, "roleType")))
    roll_select = Select(roll_dropdown)
    try:
        roll_select.select_by_visible_text(STATIC_ROLL_TYPE)
    except Exception as error:
        available = ", ".join(option.text.strip() for option in roll_select.options)
        raise ValueError(
            f"Configured roll type '{STATIC_ROLL_TYPE}' is unavailable. Available options: {available}"
        ) from error
    print(f"Roll type selected: {STATIC_ROLL_TYPE}")

    time.sleep(1.5)
    wait_for_page_ready()


def select_language():
    language_dropdown = wait.until(EC.presence_of_element_located((By.NAME, "langCd")))
    Select(language_dropdown).select_by_visible_text("ENGLISH")
    print("Language selected")
    time.sleep(1)


def wait_for_file_ready(filepath, stable_seconds=2, timeout=120):
    elapsed = 0
    last_size = -1
    stable_count = 0

    while elapsed < timeout:
        try:
            current_size = os.path.getsize(filepath)
        except OSError:
            time.sleep(0.5)
            elapsed += 0.5
            continue

        if current_size == last_size and current_size > 0:
            stable_count += 1
        else:
            stable_count = 0

        last_size = current_size

        if stable_count >= stable_seconds:
            try:
                with open(filepath, "rb"):
                    pass
                return True
            except (PermissionError, OSError):
                stable_count = 0

        time.sleep(0.5)
        elapsed += 0.5

    return False


def get_table_signature():
    try:
        rows = driver.find_elements(By.XPATH, "//table/tbody/tr")
        if not rows:
            return None

        def row_part_text(row):
            try:
                return row.find_element(By.XPATH, ".//td[@role='cell'][2]").text.strip()
            except Exception:
                return row.text.strip()

        first_text = row_part_text(rows[0])
        last_text = row_part_text(rows[-1])
        return (len(rows), first_text, last_text)
    except Exception:
        return None


def wait_for_table_to_change(previous_signature, timeout=45):
    wait_for_page_ready(timeout=10)
    if previous_signature is None:
        wait.until(EC.presence_of_element_located((By.XPATH, "//table/tbody/tr[1]")))
        return get_table_signature()

    end_time = time.time() + timeout
    while time.time() < end_time:
        wait_for_page_ready(timeout=5)
        current_signature = get_table_signature()
        if current_signature and current_signature != previous_signature:
            print(f"  Table refreshed: {previous_signature} -> {current_signature}")
            return current_signature
        time.sleep(0.8)

    raise TimeoutException("Next page was clicked, but the table did not refresh to new rows.")


def has_next_page():
    try:
        buttons = driver.find_elements(By.XPATH, "//button[contains(@class,'control-btn')]")
        for btn in buttons:
            if ">" in btn.text and "<" not in btn.text:
                disabled = btn.get_attribute("disabled")
                aria_disabled = btn.get_attribute("aria-disabled")
                class_name = btn.get_attribute("class") or ""
                if not disabled and aria_disabled != "true" and "disabled" not in class_name.lower():
                    return True
        return False
    except Exception:
        return False


def click_next_page(previous_signature=None):
    buttons = driver.find_elements(By.XPATH, "//button[contains(@class,'control-btn')]")
    for btn in buttons:
        if ">" in btn.text and "<" not in btn.text:
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
            time.sleep(0.3)
            driver.execute_script("arguments[0].click();", btn)
            print("  Clicked Next Page button; waiting for table refresh...")
            wait_for_table_to_change(previous_signature)
            return True
    return False


def get_enabled_previous_page_button():
    for btn in driver.find_elements(By.XPATH, "//button[contains(@class,'control-btn')]"):
        try:
            text = btn.text.strip()
            disabled = btn.get_attribute("disabled")
            aria_disabled = btn.get_attribute("aria-disabled")
            class_name = btn.get_attribute("class") or ""
            is_enabled = not disabled and aria_disabled != "true" and "disabled" not in class_name.lower()
            if "<" in text and ">" not in text and is_enabled:
                return btn
        except StaleElementReferenceException:
            continue
    return None


def reset_to_first_table_page():
    pages_moved = 0
    while True:
        previous_button = get_enabled_previous_page_button()
        if previous_button is None:
            if pages_moved:
                print(f"  Reset table to page 1 ({pages_moved} previous-page clicks).")
            return

        previous_signature = get_table_signature()
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", previous_button)
        time.sleep(0.3)
        driver.execute_script("arguments[0].click();", previous_button)
        wait_for_table_to_change(previous_signature)
        pages_moved += 1


def parse_part_text(part_text):
    if " - " in part_text:
        part_no, part_name = part_text.split(" - ", 1)
    else:
        part_no = part_text
        part_name = "Unknown"
    return part_no.strip(), part_name.strip()


def rename_batch(new_pdfs, parts_in_order):
    """Rename PDF -> Create Folder -> Move PDF. Returns the list of final paths.

    Each PDF ends up at:
        save_folder / "<PartNo> - <PartName>.pdf"
    where save_folder is download_folder/State/District/Constituency.
    """
    part_lookup = {str(part_no).strip(): part_name for part_no, part_name in parts_in_order}
    renamed_paths = []

    for pdf_path in new_pdfs:
        filename = os.path.basename(pdf_path)

        try:
            part_no = filename.split("-ENG-")[1].split("-WI")[0]
        except Exception:
            print(f"  ⚠ Could not determine part number from {filename}")
            continue

        if part_no not in part_lookup:
            print(f"  ⚠ Part {part_no} not found in current page selection.")
            continue

        part_name = part_lookup[part_no]
        safe_part_name = make_safe(part_name)
        base_name = f"{part_no} - {safe_part_name}"

        max_base_len = 150
        if len(base_name) > max_base_len:
            base_name = base_name[:max_base_len].rstrip()

        try:
            os.makedirs(save_folder, exist_ok=True)
        except Exception as e:
            print(f"  ⚠ Could not ensure save folder exists: {e}")

        new_path = os.path.join(save_folder, base_name + ".pdf")

        counter = 1
        while os.path.exists(new_path):
            new_path = os.path.join(save_folder, f"{base_name}_{counter}.pdf")
            counter += 1

        if not os.path.exists(pdf_path):
            print(f"  ⚠ Source file missing (possibly still syncing/quarantined): {pdf_path}")
            continue

        try:
            os.rename(pdf_path, new_path)
        except (FileNotFoundError, OSError) as e:
            print(f"  ⚠ os.rename failed ({e}), retrying with shutil.move...")
            try:
                os.makedirs(save_folder, exist_ok=True)
                shutil.move(pdf_path, new_path)
            except Exception as e2:
                print(f"  ✖ Failed to move {filename} -> {new_path}: {e2}")
                continue

        print(f"  ✔ Saved: {os.path.basename(new_path)}")
        renamed_paths.append(new_path)

    return renamed_paths


def portal_reports_file_not_found():
    try:
        warnings = driver.find_elements(
            By.XPATH,
            "//*[contains(translate(normalize-space(.), "
            "'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'file not found')]"
        )
        return any(warning.is_displayed() for warning in warnings)
    except Exception:
        return False


def process_current_page(page_num):
    """Read Part Number -> user CAPTCHA -> Download PDF -> Rename/Move ->
    process_pdf_file (OCR->CSV, saved in the same constituency folder) ->
    insert_csv_to_mysql (DB)."""
    wait_for_page_ready(timeout=10)
    driver.execute_script("window.scrollTo({top: 0, behavior: 'smooth'});")
    time.sleep(1)

    try:
        wait.until(EC.presence_of_element_located((By.XPATH, "//table/tbody/tr[1]")))
    except TimeoutException:
        print(f"  ⚠ No table found on page {page_num}.")
        return 0

    print(f"\n  {'-' * 46}")
    print(f"  Page {page_num} | Checking for missing parts...")
    print(f"  {'-' * 46}")

    js_extract_text = """
    return Array.from(document.querySelectorAll('table tbody tr')).map(row => {
        return row.cells.length > 1 ? row.cells[1].innerText.trim() : '';
    });
    """
    raw_part_texts = driver.execute_script(js_extract_text)
    total_rows = len(raw_part_texts)

    parts_to_download = []
    indices_to_select = []

    for i, part_text in enumerate(raw_part_texts):
        if part_text:
            part_no, part_name = parse_part_text(part_text)
            safe_part_name = make_safe(part_name)

            expected_filename = f"{part_no} - {safe_part_name}.pdf"
            expected_filepath = os.path.join(save_folder, expected_filename)

            if os.path.exists(expected_filepath):
                print(f"  ✔ Skipping Part {part_no} (Already exists)")
            else:
                parts_to_download.append((part_no, part_name))
                indices_to_select.append(i)

    print(f"\n  Total parts on page: {total_rows}")
    print(f"  Parts needing download: {len(parts_to_download)}")

    if not parts_to_download:
        print(f"  All parts on page {page_num} are already downloaded. Moving on...")
        return total_rows

    js_select_specific = f"""
    const indices = {json.dumps(indices_to_select)};
    const rows = document.querySelectorAll('table tbody tr');
    indices.forEach(index => {{
        if (rows[index]) {{
            const cb = rows[index].querySelector('input[type="checkbox"]');
            if (cb && !cb.checked) {{
                cb.click();
            }}
        }}
    }});
    """
    driver.execute_script(js_select_specific)
    print("  ✔ Targeted checkboxes clicked.")
    time.sleep(1)

    pdfs_before = set(glob.glob(os.path.join(download_folder, "*.pdf")))
    expected_count = len(parts_to_download)

    while True:
        input(f"\n  [Page {page_num}] {expected_count} parts selected for download. "
              f"Solve the CAPTCHA, then press ENTER...")

        try:
            download_button = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Download Selected PDFs')]"))
            )
            driver.execute_script("arguments[0].click();", download_button)
            print("  Download button clicked. Checking if files are downloading...")
        except Exception:
            print("  ⚠ Could not click the download button. Please make sure you are on the right page.")
            continue

        download_started = False
        for _ in range(15):
            if portal_reports_file_not_found():
                print(
                    "  ⚠ Portal reported 'File not found' for this selection. "
                    "Skipping this page instead of retrying the CAPTCHA."
                )
                return total_rows

            part_files = glob.glob(os.path.join(download_folder, "*.part"))
            pdfs_now = set(glob.glob(os.path.join(download_folder, "*.pdf")))
            truly_new = pdfs_now - pdfs_before

            if part_files or truly_new:
                download_started = True
                break
            time.sleep(1)

        if download_started:
            print("  ✔ Download started successfully!")
            break
        else:
            print("  ⚠ No download detected. The CAPTCHA might have been incorrect or expired.")
            print("  Please refresh the CAPTCHA on the page, enter it again, and press ENTER.")

    timeout_seconds = max(120, expected_count * 20)
    elapsed = 0
    new_pdfs = set()

    while elapsed < timeout_seconds:
        part_files = glob.glob(os.path.join(download_folder, "*.part"))
        if part_files:
            time.sleep(0.5)
            elapsed += 0.5
            continue

        pdfs_now = set(glob.glob(os.path.join(download_folder, "*.pdf")))
        truly_new = pdfs_now - pdfs_before
        new_pdfs = truly_new

        if len(truly_new) >= expected_count:
            print(f"  All {len(truly_new)} new PDFs detected.")
            break

        time.sleep(0.5)
        elapsed += 0.5

    if not new_pdfs:
        print(f"  ⚠ Timeout: no new PDFs detected on page {page_num}. Skipping rename.")
        return total_rows

    if len(new_pdfs) < expected_count:
        print(f"  ⚠ Only {len(new_pdfs)} of {expected_count} expected PDFs were detected "
              f"before timeout. Proceeding with what was found.")

    ready_pdfs = []
    for pdf in new_pdfs:
        print(f"  Waiting for file to finish writing: {os.path.basename(pdf)}")
        if wait_for_file_ready(pdf, stable_seconds=2, timeout=120):
            ready_pdfs.append(pdf)
        else:
            print(f"  ⚠ File still locked, skipping: {os.path.basename(pdf)}")

    # Rename PDF -> Create Folder -> Move PDF
    # (rename_batch already places each PDF inside its constituency's
    # save_folder: download_folder/State/District/Constituency/part.pdf)
    renamed_paths = rename_batch(ready_pdfs, parts_to_download)

    for pdf_path in renamed_paths:
        try:
            # process_pdf_file() writes the CSV via file_path.with_suffix(".csv"),
            # so the CSV is created directly alongside the PDF in save_folder.
            # No separate "input_pdfs" staging folder is used.
            csv_path = process_pdf_file(pdf_path)

            insert_csv_to_mysql(csv_path)

        except Exception as e:
            print(f"Processing failed for {pdf_path}: {e}")

    return total_rows


def process_selected_constituency(state_name, district_name, constituency_name):
    global save_folder

    safe_state = make_safe(state_name)
    safe_district = make_safe(district_name)
    safe_constituency = make_safe(constituency_name)

    save_folder = os.path.join(download_folder, safe_state, safe_district, safe_constituency)
    os.makedirs(save_folder, exist_ok=True)
    print(f"Save folder created: {save_folder}")

    reset_to_first_table_page()

    page_number = 1
    total_parts = 0

    while True:
        print(f"\n{'#' * 60}")
        print(f"  PROCESSING PAGE {page_number}")
        print(f"{'#' * 60}")

        rows_done = process_current_page(page_number)
        total_parts += rows_done

        print(f"\n  Page {page_number} complete. Parts processed so far: {total_parts}")

        driver.execute_script("window.scrollTo({top: document.body.scrollHeight, behavior: 'smooth'});")
        time.sleep(1)
        if has_next_page():
            previous_signature = get_table_signature()
            print(f"  Moving to page {page_number + 1}...")
            click_next_page(previous_signature)
            page_number += 1
        else:
            print("\n  No more pages for this constituency.")
            break

    print(f"\nCONSTITUENCY COMPLETE: {state_name} / {district_name} / {constituency_name}")
    print(f"   Total parts processed : {total_parts}")
    print(f"   PDFs saved in         : {save_folder}")


def select_static_state(state_name_to_select):
    wanted = state_name_to_select.strip().casefold()
    state_options = get_select_options("stateCode")

    for state_value, raw_state_text in state_options:
        display_name = clean_option_text(raw_state_text)
        if display_name.casefold() == wanted:
            selected_text = select_dropdown_value("stateCode", state_value)
            return state_value, clean_option_text(selected_text)

    available_states = ", ".join(clean_option_text(text) for _, text in state_options)
    raise ValueError(f"Static state '{state_name_to_select}' was not found. Available states: {available_states}")


# ═══════════════════════════════════════════════════════════════════════════
#  MAIN — Select State -> Select District -> Select Constituency -> ...
#  -> Next Constituency -> Next District -> Finish
# ═══════════════════════════════════════════════════════════════════════════

try:
    _, state_name = select_static_state(STATIC_STATE_NAME)
    print(f"\nSTATIC STATE selected: {state_name}")
    time.sleep(1.5)
    wait_for_page_ready()

    select_year_and_roll_type()

    wait.until(lambda d: len(Select(d.find_element(By.NAME, "district")).options) > 1)
    district_options = get_select_options("district")
    print(f"Found {len(district_options)} districts in {state_name}.")
except Exception as e:
    driver.quit()
    raise RuntimeError(f"Could not load static state '{STATIC_STATE_NAME}': {e}") from e


for district_value, raw_district_text in district_options:
    try:
        selected_district_text = select_dropdown_value("district", district_value)
        district_name = clean_option_text(selected_district_text)
        print(f"\nDistrict selected: {district_name}")
        time.sleep(1.5)
        wait_for_page_ready()

        constituency_options = get_constituency_option_texts()
        print(f"Found {len(constituency_options)} constituencies in {district_name}.")
    except Exception as e:
        print(f"  Error loading district {raw_district_text}: {e}")
        continue

    for raw_constituency_text in constituency_options:
        try:
            selected_constituency_text = select_constituency_with_retry(raw_constituency_text)
            constituency_name = clean_option_text(selected_constituency_text or raw_constituency_text) or "Constituency"
            print(f"Constituency selected: {constituency_name}")
            time.sleep(1)

            select_language()

            process_selected_constituency(state_name, district_name, constituency_name)
        except Exception as e:
            print(f"  Error processing constituency {raw_constituency_text}: {e}")
    continue

print(f"\nALL DISTRICTS AND CONSTITUENCIES IN {state_name} PROCESSED!")
driver.quit()
