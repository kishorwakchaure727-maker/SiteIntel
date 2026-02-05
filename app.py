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
            REQUEST_DELAY = 0.15
            CRAWL_TIMEOUT = 6

            SHORT_FORMS = {
                "RD": "ROAD", "ST": "STREET", "AVE": "AVENUE",
                "BLVD": "BOULEVARD", "DR": "DRIVE", "LN": "LANE", "PL": "PLACE"
            }

            STANDARD_COUNTRIES = {
                "USA": "UNITED STATES OF AMERICA",
                "US": "UNITED STATES OF AMERICA",
            }

            STREET_KEYWORDS = r"\b(STREET|ST\.|ROAD|RD\.|AVE|AVENUE|BOULEVARD|BLVD|DR|DRIVE|LANE|LN|WAY|TERRACE|PLAZA|PL|COURT|CT)\b"

            # -------------------------------------------------
            # UTILITIES
            # -------------------------------------------------
            def ensure_scheme(url: str) -> str:
                if not url.startswith("http"):
                    return "https://" + url.lstrip("/")
                return url

            def normalize_text(text: str) -> str:
                if not text:
                    return ""
                text = unidecode(text).strip()
                return re.sub(r"\s+", " ", text)

            def hash_address(addr: dict) -> str:
                key = "|".join([addr.get(k, "") for k in ["STREET ADDRESS 1", "CITY", "STATE", "PIN CODE", "COUNTRY"]])
                return hashlib.md5(key.encode()).hexdigest()

            # -------------------------------------------------
            # PAGE SEARCH & EXTRACTION
            # -------------------------------------------------
            def candidate_paths():
                return ["/contact", "/contact-us", "/about", "/about-us", "/locations", "/location", "/contactus", "/company/contact"]

            def find_pages_from_home(home_url: str, max_pages=10):
                home = ensure_scheme(home_url)
                pages = [home]
                try:
                    r = requests.get(home, headers=HEADERS, timeout=CRAWL_TIMEOUT)
                    soup = BeautifulSoup(r.text, "html.parser")
                    for a in soup.find_all("a", href=True):
                        href = a["href"]
                        if href.startswith("/"):
                            pages.append(home.rstrip("/") + href)
                        elif href.startswith(home) or href.startswith("http") and home in href:
                            pages.append(href)
                        if len(set(pages)) >= max_pages:
                            break
                except Exception:
                    pass

                # add common paths
                base = re.sub(r"/+$", "", home)
                for p in candidate_paths():
                    pages.append(base + p)

                # dedupe
                out = []
                for p in pages:
                    if p not in out:
                        out.append(p)
                return out[:max_pages]

            def ddg_site_search(domain: str, query_terms="contact address", max_results=5):
                # Use DuckDuckGo HTML to search site pages for contact/address clues
                q = f"site:{domain} {query_terms}"
                try:
                    url = "https://html.duckduckgo.com/html/"
                    res = requests.post(url, data={"q": q}, headers=HEADERS, timeout=CRAWL_TIMEOUT)
                    soup = BeautifulSoup(res.text, "html.parser")
                    links = []
                    for a in soup.find_all("a", href=True):
                        href = a["href"]
                        if href.startswith("http") and domain in href:
                            links.append(href)
                        if len(links) >= max_results:
                            break
                    return links
                except Exception:
                    return []

            def find_address_in_html(text: str):
                if not text:
                    return None
                # Search for an address-like block using keywords and number patterns
                for line in text.splitlines():
                    line = line.strip()
                    if len(line) < 10:
                        continue
                    if re.search(STREET_KEYWORDS, line, re.I) and re.search(r"\d{1,5}", line):
                        return normalize_text(line)
                # fallback: search paragraphs
                m = re.search(r"([\w\s,.-]{10,150}" + STREET_KEYWORDS + r"[\w\s,.-]{0,150})", text, re.I)
                if m:
                    return normalize_text(m.group(0))
                return None

            def extract_address(website: str):
                # returns (raw_address_text, found_page_url)
                if not website:
                    return "", ""
                domain = re.sub(r"https?://", "", website).split("/", 1)[0]
                pages = find_pages_from_home(website, max_pages=12)

                # Try pages from site
                for p in pages:
                    try:
                        r = requests.get(ensure_scheme(p), headers=HEADERS, timeout=CRAWL_TIMEOUT)
                        soup = BeautifulSoup(r.text, "html.parser")
                        # check <address>
                        addr_tag = soup.find("address")
                        if addr_tag:
                            txt = addr_tag.get_text(" ", strip=True)
                            found = find_address_in_html(txt)
                            if found:
                                return found, ensure_scheme(p)

                        # check footer and contact sections
                        for selector in ["footer", "p", "div", "section"]:
                            for el in soup.find_all(selector):
                                txt = el.get_text(" ", strip=True)
                                if re.search(STREET_KEYWORDS, txt, re.I) and re.search(r"\d{1,5}", txt):
                                    return normalize_text(txt), ensure_scheme(p)
                    except Exception:
                        continue

                # If not found, try DDG site search for likely pages
                ddg_links = ddg_site_search(domain, query_terms="contact address", max_results=6)
                for link in ddg_links:
                    try:
                        r = requests.get(link, headers=HEADERS, timeout=CRAWL_TIMEOUT)
                        soup = BeautifulSoup(r.text, "html.parser")
                        txt = soup.get_text(" ", strip=True)
                        found = find_address_in_html(txt)
                        if found:
                            return found, link
                    except Exception:
                        continue

                return "", ""

            # -------------------------------------------------
            # PARSING / STANDARDIZATION
            # -------------------------------------------------
            def standardize_address(raw: str) -> dict:
                # parse raw address into components heuristically
                out = {
                    "STREET ADDRESS 1": "",
                    "STREET ADDRESS 2": "",
                    "CITY": "",
                    "STATE": "",
                    "PIN CODE": "",
                    "COUNTRY": "",
                }
                if not raw:
                    return out
                raw = normalize_text(raw)
                # replace short forms
                for k, v in SHORT_FORMS.items():
                    raw = re.sub(rf"\b{k}\b", v, raw, flags=re.I)

                # split by commas
                parts = [p.strip() for p in re.split(r",|;|\n", raw) if p.strip()]
                if parts:
                    out["STREET ADDRESS 1"] = parts[0]
                if len(parts) >= 2:
                    # try to assign city/state/pin from the last part
                    last = parts[-1]
                    # postal code
                    m = re.search(r"(\d{5}(?:-\d{4})?)", last)
                    if m:
                        out["PIN CODE"] = m.group(1)
                        last = last.replace(m.group(1), "").strip()
                    # state (2-letter) or full
                    tokens = [t.strip() for t in last.split() if t.strip()]
                    if tokens:
                        if len(tokens) == 1 and len(tokens[0]) <= 3:
                            out["STATE"] = tokens[0]
                        else:
                            out["COUNTRY"] = tokens[-1]

                # try to extract city from middle part
                if len(parts) >= 3:
                    out["CITY"] = parts[-2]

                return out

            # -------------------------------------------------
            # FREE NOMINATIM ENRICHMENT (optional, free)
            # -------------------------------------------------
            def enrich_with_nominatim(record: dict) -> dict:
                # Only call when we have at least street or city
                q = ", ".join([record.get("STREET ADDRESS 1", ""), record.get("CITY", ""), record.get("STATE", ""), record.get("COUNTRY", "")])
                q = q.strip().strip(",")
                if not q:
                    return record
                try:
                    url = "https://nominatim.openstreetmap.org/search"
                    res = requests.get(url, params={"q": q, "format": "json", "addressdetails": 1, "limit": 1}, headers={**HEADERS, "User-Agent": "SiteIntel/1.0 (mailto:you@example.com)"}, timeout=10)
                    data = res.json()
                    if data:
                        addr = data[0].get("address", {})
                        if not record.get("CITY") and addr.get("city"):
                            record["CITY"] = addr.get("city").upper()
                        if not record.get("STATE") and addr.get("state"):
                            record["STATE"] = addr.get("state").upper()
                        if not record.get("PIN CODE") and addr.get("postcode"):
                            record["PIN CODE"] = addr.get("postcode")
                        if not record.get("COUNTRY") and addr.get("country"):
                            record["COUNTRY"] = addr.get("country").upper()
                        time.sleep(1)
                except Exception:
                    pass
                return record

            # -------------------------------------------------
            # CONFIDENCE SCORE
            # -------------------------------------------------
            def calculate_confidence(addr: dict) -> int:
                score = 0
                if addr.get("STREET ADDRESS 1"):
                    score += 40
                if addr.get("CITY"):
                    score += 15
                if addr.get("STATE"):
                    score += 15
                if addr.get("PIN CODE"):
                    score += 15
                if addr.get("COUNTRY"):
                    score += 15
                return min(score, 100)

            # -------------------------------------------------
            # EXCEL EXPORT
            # -------------------------------------------------
            def generate_excel(data: list) -> bytes:
                wb = Workbook()
                ws = wb.active
                ws.title = "SiteIntel Output"

                ws.merge_cells("A1:K1")
                ws["A1"] = "SiteIntel – By Kishor"
                ws["A1"].font = Font(size=16, bold=True)
                ws["A1"].alignment = Alignment(horizontal="center")

                headers = [
                    "STREET ADDRESS 1", "STREET ADDRESS 2",
                    "CITY", "STATE", "PIN CODE", "COUNTRY",
                    "CONFIDENCE SCORE", "DUPLICATE FLAG",
                    "MASTER RECORD ID", "DATA SOURCE LINK", "FOUND PAGE"
                ]

                ws.append(headers)
                for c in ws[2]:
                    c.font = Font(bold=True)

                for r in data:
                    ws.append([r.get(h, "") for h in headers])

                ref = f"A2:K{len(data)+2}"
                table = Table(displayName="AddressTable", ref=ref)
                table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
                ws.add_table(table)

                for col in range(1, len(headers)+1):
                    ws.column_dimensions[get_column_letter(col)].width = 25

                ws.freeze_panes = "A3"
                out = BytesIO()
                wb.save(out)
                return out.getvalue()

            # -------------------------------------------------
            # STREAMLIT UI
            # -------------------------------------------------
            st.set_page_config(page_title="SiteIntel – By Kishor", layout="wide", page_icon="📍")
            st.markdown("""
            <div style='display:flex;align-items:center;gap:12px'>
              <div style='font-size:36px'>📍</div>
              <div>
                <h1 style='margin:0'>SiteIntel</h1>
                <div style='color:gray'>Enterprise Address Intelligence — Semi-Agentic (free sources)</div>
              </div>
            </div>
            """, unsafe_allow_html=True)

            mode = st.radio("Mode", ["Batch (Excel upload)", "Single Company"])

            if mode == "Batch (Excel upload)":
                uploaded = st.file_uploader("Upload Excel with company websites", type=["xlsx", "xls"])
                if st.button("🚀 Process Batch"):
                    if not uploaded:
                        st.warning("Upload a file first.")
                        st.stop()

                    df = pd.read_excel(uploaded)

                    # reuse the website column finder
                    def find_website_column(df: pd.DataFrame):
                        for c in df.columns:
                            if re.search(r"web|site|url", str(c), re.I):
                                return c
                        for c in df.columns:
                            try:
                                s = df[c].astype(str).str.strip().fillna("")
                            except Exception:
                                continue
                            if s.str.startswith("http").any():
                                return c
                        for c in df.columns:
                            try:
                                s = df[c].astype(str).str.strip().fillna("")
                            except Exception:
                                continue
                            if s.str.startswith("www.").any():
                                return c
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
                        st.error("No website column found. Please include a column with company website URLs.")
                        st.stop()

                    progress = st.progress(0)
                    records = []
                    seen = {}

                    for i, site in enumerate(df[url_col].astype(str)):
                        raw, page = extract_address(site)
                        parsed = standardize_address(raw)
                        parsed["DATA SOURCE LINK"] = site
                        parsed["FOUND PAGE"] = page
                        parsed = enrich_with_nominatim(parsed)
                        parsed["CONFIDENCE SCORE"] = calculate_confidence(parsed)

                        h = hash_address(parsed)
                        if h in seen:
                            parsed["DUPLICATE FLAG"] = "YES"
                            parsed["MASTER RECORD ID"] = seen[h]
                        else:
                            parsed["DUPLICATE FLAG"] = "NO"
                            parsed["MASTER RECORD ID"] = h[:8]
                            seen[h] = h[:8]

                        records.append(parsed)
                        progress.progress((i + 1) / max(1, len(df)))

                    st.success(f"Processed {len(records)} records")
                    st.dataframe(pd.DataFrame(records), use_container_width=True)
                    st.download_button("📥 Download Excel", generate_excel(records), "siteintel_output.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            else:
                st.subheader("Single company check")
                name = st.text_input("Company name (optional)")
                website = st.text_input("Official website (e.g. example.com or https://example.com)")
                if st.button("🔎 Check Single Company"):
                    if not website:
                        st.warning("Enter a website URL first.")
                        st.stop()
                    raw, page = extract_address(website)
                    parsed = standardize_address(raw)
                    parsed["DATA SOURCE LINK"] = website
                    parsed["FOUND PAGE"] = page
                    parsed = enrich_with_nominatim(parsed)
                    parsed["CONFIDENCE SCORE"] = calculate_confidence(parsed)
                    st.write("**Found page:**", page or "(not found)")
                    st.write("**Raw extract:**", raw or "(no extract)")
                    st.table(pd.DataFrame([parsed]))
                    st.download_button("📥 Download Excel", generate_excel([parsed]), "siteintel_single.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

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
