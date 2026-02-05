import streamlit as st
import pandas as pd
import requests
import re
import time
import hashlib
from bs4 import BeautifulSoup
from io import BytesIO
from unidecode import unidecode
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

# -------------------------------------------------
# CONFIG
# -------------------------------------------------
HEADERS = {"User-Agent": "Mozilla/5.0"}
REQUEST_DELAY = 0.15  # Google Maps throttle

try:
    GOOGLE_MAPS_API_KEY = st.secrets.get("GOOGLE_MAPS_API_KEY", "")
except:
    GOOGLE_MAPS_API_KEY = ""

SHORT_FORMS = {
    "RD": "ROAD", "ST": "STREET", "AVE": "AVENUE",
    "BLVD": "BOULEVARD", "DR": "DRIVE",
    "LN": "LANE", "PL": "PLACE"
}

STANDARD_COUNTRIES = {
    "USA": "UNITED STATES OF AMERICA",
    "US": "UNITED STATES OF AMERICA",
    "UNITED STATES": "UNITED STATES OF AMERICA",
    "UK": "UNITED KINGDOM OF GREAT BRITAIN AND NORTHERN IRELAND",
    "UNITED KINGDOM": "UNITED KINGDOM OF GREAT BRITAIN AND NORTHERN IRELAND"
}

# -------------------------------------------------
# HELPERS
# -------------------------------------------------
def normalize_text(text: str) -> str:
    text = unidecode(text).upper()
    text = re.sub(r"[^\w\s]", "", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()

def hash_address(addr: dict) -> str:
    key = "|".join([
        addr.get("STREET ADDRESS 1", ""),
        addr.get("CITY", ""),
        addr.get("STATE", ""),
        addr.get("PIN CODE", ""),
        addr.get("COUNTRY", "")
    ])
    return hashlib.md5(key.encode()).hexdigest()

# -------------------------------------------------
# ADDRESS EXTRACTION
# -------------------------------------------------
def extract_address(website: str) -> str:
    try:
        if not website.startswith("http"):
            website = "https://" + website

        res = requests.get(website, headers=HEADERS, timeout=10)
        soup = BeautifulSoup(res.text, "html.parser")

        tag = soup.find("address")
        if tag:
            return tag.get_text(" ", strip=True)

        for el in soup.find_all(["footer", "p", "div"]):
            txt = el.get_text(" ", strip=True)
            if re.search(r'\b(STREET|ROAD|AVE|AVENUE|BOULEVARD|DRIVE|LANE)\b', txt.upper()):
                return txt

    except Exception:
        pass

    return ""

# -------------------------------------------------
# STANDARDIZATION
# -------------------------------------------------
def standardize_address(raw: str) -> str:
    if not raw:
        return ""

    addr = normalize_text(raw)

    for k, v in SHORT_FORMS.items():
        addr = re.sub(rf"\b{k}\b", v, addr)

    return addr

# -------------------------------------------------
# GOOGLE MAPS ENRICHMENT
# -------------------------------------------------
def enrich_google_maps(record: dict) -> dict:
    if not GOOGLE_MAPS_API_KEY or not record["STREET ADDRESS 1"]:
        return record

    try:
        query = record["STREET ADDRESS 1"]
        url = (
            "https://maps.googleapis.com/maps/api/geocode/json"
            f"?address={query}&key={GOOGLE_MAPS_API_KEY}"
        )

        res = requests.get(url, timeout=10).json()

        if res.get("status") != "OK":
            return record

        for comp in res["results"][0]["address_components"]:
            if "locality" in comp["types"]:
                record["CITY"] = comp["long_name"].upper()
            if "administrative_area_level_1" in comp["types"]:
                record["STATE"] = comp["short_name"].upper()
            if "postal_code" in comp["types"]:
                record["PIN CODE"] = comp["long_name"]
            if "country" in comp["types"]:
                c = comp["long_name"].upper()
                record["COUNTRY"] = STANDARD_COUNTRIES.get(c, c)

        time.sleep(REQUEST_DELAY)

    except Exception:
        pass

    return record

# -------------------------------------------------
# CONFIDENCE SCORE
# -------------------------------------------------
def calculate_confidence(addr: dict) -> int:
    score = 0

    if addr["STREET ADDRESS 1"]:
        score += 40
    if addr["CITY"]:
        score += 15
    if addr["STATE"]:
        score += 15
    if addr["PIN CODE"]:
        score += 15
    if addr["COUNTRY"]:
        score += 15

    return min(score, 100)

# -------------------------------------------------
# EXCEL EXPORT
# -------------------------------------------------
def generate_excel(data: list) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "SiteIntel Output"

    ws.merge_cells("A1:J1")
    ws["A1"] = "SiteIntel – By Kishor"
    ws["A1"].font = Font(size=16, bold=True)
    ws["A1"].alignment = Alignment(horizontal="center")

    headers = [
        "STREET ADDRESS 1", "STREET ADDRESS 2",
        "CITY", "STATE", "PIN CODE", "COUNTRY",
        "CONFIDENCE SCORE", "DUPLICATE FLAG",
        "MASTER RECORD ID", "DATA SOURCE LINK"
    ]

    ws.append(headers)

    for c in ws[2]:
        c.font = Font(bold=True)

    for r in data:
        ws.append([r.get(h, "") for h in headers])

    table = Table(
        displayName="AddressTable",
        ref=f"A2:J{len(data)+2}"
    )
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium9",
        showRowStripes=True
    )
    ws.add_table(table)

    for col in range(1, 11):
        ws.column_dimensions[get_column_letter(col)].width = 25

    ws.freeze_panes = "A3"

    out = BytesIO()
    wb.save(out)
    return out.getvalue()

# -------------------------------------------------
# STREAMLIT UI
# -------------------------------------------------
import streamlit as st
import pandas as pd
import requests
import re
import time
import hashlib
from bs4 import BeautifulSoup
from io import BytesIO
from unidecode import unidecode
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

# -------------------------------------------------
# CONFIG
# -------------------------------------------------
HEADERS = {"User-Agent": "Mozilla/5.0"}
REQUEST_DELAY = 0.15  # Google Maps throttle

try:
    GOOGLE_MAPS_API_KEY = st.secrets.get("GOOGLE_MAPS_API_KEY", "")
except:
    GOOGLE_MAPS_API_KEY = ""

SHORT_FORMS = {
    "RD": "ROAD", "ST": "STREET", "AVE": "AVENUE",
    "BLVD": "BOULEVARD", "DR": "DRIVE",
    "LN": "LANE", "PL": "PLACE"
}

STANDARD_COUNTRIES = {
    "USA": "UNITED STATES OF AMERICA",
    "US": "UNITED STATES OF AMERICA",
    "UNITED STATES": "UNITED STATES OF AMERICA",
    "UK": "UNITED KINGDOM OF GREAT BRITAIN AND NORTHERN IRELAND",
    "UNITED KINGDOM": "UNITED KINGDOM OF GREAT BRITAIN AND NORTHERN IRELAND"
}

# -------------------------------------------------
# HELPERS
# -------------------------------------------------
def normalize_text(text: str) -> str:
    text = unidecode(text).upper()
    text = re.sub(r"[^\w\s]", "", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()

def hash_address(addr: dict) -> str:
    key = "|".join([
        addr.get("STREET ADDRESS 1", ""),
        addr.get("CITY", ""),
        addr.get("STATE", ""),
        addr.get("PIN CODE", ""),
        addr.get("COUNTRY", "")
    ])
    return hashlib.md5(key.encode()).hexdigest()

# -------------------------------------------------
# ADDRESS EXTRACTION
# -------------------------------------------------
def extract_address(website: str) -> str:
    try:
        if not website.startswith("http"):
            website = "https://" + website

        res = requests.get(website, headers=HEADERS, timeout=10)
        soup = BeautifulSoup(res.text, "html.parser")

        tag = soup.find("address")
        if tag:
            return tag.get_text(" ", strip=True)

        for el in soup.find_all(["footer", "p", "div"]):
            txt = el.get_text(" ", strip=True)
            if re.search(r'\b(STREET|ROAD|AVE|AVENUE|BOULEVARD|DRIVE|LANE)\b', txt.upper()):
                return txt

    except Exception:
        pass

    return ""

# -------------------------------------------------
# STANDARDIZATION
# -------------------------------------------------
def standardize_address(raw: str) -> str:
    if not raw:
        return ""

    addr = normalize_text(raw)

    for k, v in SHORT_FORMS.items():
        addr = re.sub(rf"\b{k}\b", v, addr)

    return addr

# -------------------------------------------------
# GOOGLE MAPS ENRICHMENT
# -------------------------------------------------
def enrich_google_maps(record: dict) -> dict:
    if not GOOGLE_MAPS_API_KEY or not record["STREET ADDRESS 1"]:
        return record

    try:
        query = record["STREET ADDRESS 1"]
        url = (
            "https://maps.googleapis.com/maps/api/geocode/json"
            f"?address={query}&key={GOOGLE_MAPS_API_KEY}"
        )

        res = requests.get(url, timeout=10).json()

        if res.get("status") != "OK":
            return record

        for comp in res["results"][0]["address_components"]:
            if "locality" in comp["types"]:
                record["CITY"] = comp["long_name"].upper()
            if "administrative_area_level_1" in comp["types"]:
                record["STATE"] = comp["short_name"].upper()
            if "postal_code" in comp["types"]:
                record["PIN CODE"] = comp["long_name"]
            if "country" in comp["types"]:
                c = comp["long_name"].upper()
                record["COUNTRY"] = STANDARD_COUNTRIES.get(c, c)

        time.sleep(REQUEST_DELAY)

    except Exception:
        pass

    return record

# -------------------------------------------------
# CONFIDENCE SCORE
# -------------------------------------------------
def calculate_confidence(addr: dict) -> int:
    score = 0

    if addr["STREET ADDRESS 1"]:
        score += 40
    if addr["CITY"]:
        score += 15
    if addr["STATE"]:
        score += 15
    if addr["PIN CODE"]:
        score += 15
    if addr["COUNTRY"]:
        score += 15

    return min(score, 100)

# -------------------------------------------------
# EXCEL EXPORT
# -------------------------------------------------
def generate_excel(data: list) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "SiteIntel Output"

    ws.merge_cells("A1:J1")
    ws["A1"] = "SiteIntel – By Kishor"
    ws["A1"].font = Font(size=16, bold=True)
    ws["A1"].alignment = Alignment(horizontal="center")

    headers = [
        "STREET ADDRESS 1", "STREET ADDRESS 2",
        "CITY", "STATE", "PIN CODE", "COUNTRY",
        "CONFIDENCE SCORE", "DUPLICATE FLAG",
        "MASTER RECORD ID", "DATA SOURCE LINK"
    ]

    ws.append(headers)

    for c in ws[2]:
        c.font = Font(bold=True)

    for r in data:
        ws.append([r.get(h, "") for h in headers])

    table = Table(
        displayName="AddressTable",
        ref=f"A2:J{len(data)+2}"
    )
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium9",
        showRowStripes=True
    )
    ws.add_table(table)

    for col in range(1, 11):
        ws.column_dimensions[get_column_letter(col)].width = 25

    ws.freeze_panes = "A3"

    out = BytesIO()
    wb.save(out)
    return out.getvalue()

# -------------------------------------------------
# STREAMLIT UI
# -------------------------------------------------
st.set_page_config(page_title="SiteIntel – By Kishor", layout="wide", page_icon="📍")
st.title("📍 SiteIntel")
st.caption("Enterprise Address Intelligence")

uploaded = st.file_uploader("Upload Excel with company websites", type=["xlsx", "xls"])

if st.button("🚀 Process"):

    if not uploaded:
        st.warning("Upload a file first.")
        st.stop()

    df = pd.read_excel(uploaded)

    def find_website_column(df: pd.DataFrame):
        # 1) Column name heuristics
        for c in df.columns:
            if re.search(r"web|site|url", str(c), re.I):
                return c

        # 2) Values starting with http
        for c in df.columns:
            try:
                s = df[c].astype(str).str.strip().fillna("")
            except Exception:
                continue
            if s.str.startswith("http").any():
                return c

        # 3) Values starting with www.
        for c in df.columns:
            try:
                s = df[c].astype(str).str.strip().fillna("")
            except Exception:
                continue
            if s.str.startswith("www.").any():
                return c

        # 4) Values that look like domains (e.g. example.com)
        domain_re = re.compile(r"\w+\.\w+")
        for c in df.columns:
            try:
                s = df[c].astype(str).str.strip().fillna("")
            except Exception:
                continue
            if s.apply(lambda x: bool(domain_re.search(x))).any():
                return c

        return None

    url_col = find_website_column(df)

    if not url_col:
        st.error("No website column found.")
        st.stop()

    progress = st.progress(0)
    records = []
    seen = {}

    for i, site in enumerate(df[url_col].astype(str)):
        raw = extract_address(site)
        std = standardize_address(raw)

        rec = {
            "STREET ADDRESS 1": std,
            "STREET ADDRESS 2": "",
            "CITY": "",
            "STATE": "",
            "PIN CODE": "",
            "COUNTRY": "",
            "DATA SOURCE LINK": site
        }

        rec = enrich_google_maps(rec)
        rec["CONFIDENCE SCORE"] = calculate_confidence(rec)

        h = hash_address(rec)
        if h in seen:
            rec["DUPLICATE FLAG"] = "YES"
            rec["MASTER RECORD ID"] = seen[h]
        else:
            rec["DUPLICATE FLAG"] = "NO"
            rec["MASTER RECORD ID"] = h[:8]
            seen[h] = h[:8]

        records.append(rec)
        progress.progress((i + 1) / len(df))

    st.success(f"Processed {len(records)} records")

    st.dataframe(pd.DataFrame(records), use_container_width=True)

    st.download_button(
        "📥 Download Excel",
        generate_excel(records),
        "siteintel_output.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
