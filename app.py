import streamlit as st
import pandas as pd
import requests
from bs4 import BeautifulSoup
from io import BytesIO
from unidecode import unidecode
import re
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo

# -------------------------------
# Configurations
# -------------------------------
try:
    GOOGLE_MAPS_API_KEY = st.secrets.get("GOOGLE_MAPS_API_KEY", "YOUR_GOOGLE_MAPS_API_KEY")
except:
    GOOGLE_MAPS_API_KEY = "YOUR_GOOGLE_MAPS_API_KEY"  # Fallback for local development

short_forms = {
    "RD": "ROAD", "ST": "STREET", "AVE": "AVENUE", "BLVD": "BOULEVARD",
    "DR": "DRIVE", "LN": "LANE", "PL": "PLACE", "CT": "COURT", "PKWY": "PARKWAY", "SQ": "SQUARE"
}

standard_countries = {
    "USA": "UNITED STATES OF AMERICA", "US": "UNITED STATES OF AMERICA",
    "UNITED STATES": "UNITED STATES OF AMERICA",
    "UK": "UNITED KINGDOM OF GREAT BRITAIN AND NORTHERN IRELAND",
    "UNITED KINGDOM": "UNITED KINGDOM OF GREAT BRITAIN AND NORTHERN IRELAND",
    "CHINA": "CHINA", "RUSSIA": "RUSSIAN FEDERATION",
    "SOUTH KOREA": "KOREA (REPUBLIC OF)", "KOREA": "KOREA (REPUBLIC OF)"
}

us_states = {
    "ALABAMA": "AL", "ALASKA": "AK", "ARIZONA": "AZ", "CALIFORNIA": "CA",
    "NEW YORK": "NY", "TEXAS": "TX", "FLORIDA": "FL", "ILLINOIS": "IL"
    # Add all states as needed
}

# -------------------------------
# Functions
# -------------------------------
def extract_address(website):
    try:
        response = requests.get(website, timeout=10)
        soup = BeautifulSoup(response.text, 'html.parser')
        address_tag = soup.find('address')
        if address_tag:
            return address_tag.get_text(separator=",")
        text = soup.get_text()
        lines = text.split('\n')
        for line in lines:
            for keyword in ["Head Office", "Corporate Office", "Address"]:
                if keyword.lower() in line.lower():
                    return line.strip()
        return ""
    except Exception:
        return ""

def standardize_address(raw_address):
    address = unidecode(raw_address).upper()
    for short, full in short_forms.items():
        address = re.sub(rf"\b{short}\b", full, address)

    parts = [p.strip() for p in address.split(",")]
    street_1 = parts[0] if len(parts) > 0 else ""
    street_2 = parts[1] if len(parts) > 1 else ""
    city = parts[2] if len(parts) > 2 else ""
    state = parts[3] if len(parts) > 3 else ""
    pin_code = parts[4] if len(parts) > 4 else ""
    country = parts[5] if len(parts) > 5 else ""

    for key, value in standard_countries.items():
        if country.startswith(key):
            country = value

    if country == "UNITED STATES OF AMERICA" and state in us_states:
        state = us_states[state]

    return {
        "STREET ADDRESS 1": street_1,
        "STREET ADDRESS 2": street_2,
        "CITY": city,
        "STATE": state,
        "PIN CODE": pin_code,
        "COUNTRY": country
    }

def enrich_with_google_maps(address):
    query = f"{address['STREET ADDRESS 1']} {address['CITY']} {address['STATE']} {address['COUNTRY']}"
    url = f"https://maps.googleapis.com/maps/api/geocode/json?address={query}&key={GOOGLE_MAPS_API_KEY}"
    try:
        response = requests.get(url)
        data = response.json()
        if data['status'] == 'OK':
            components = data['results'][0]['address_components']
            for comp in components:
                if 'locality' in comp['types']:
                    address['CITY'] = comp['long_name'].upper()
                if 'administrative_area_level_1' in comp['types']:
                    state_name = comp['long_name'].upper()
                    address['STATE'] = us_states.get(state_name, state_name)
                if 'country' in comp['types']:
                    address['COUNTRY'] = standard_countries.get(comp['long_name'].upper(), comp['long_name'].upper())
                if 'postal_code' in comp['types']:
                    address['PIN CODE'] = comp['long_name']
    except Exception:
        pass
    return address

def generate_excel(address_list):
    wb = Workbook()
    ws = wb.active
    ws.title = "Standardized Addresses"

    ws.merge_cells('A1:G1')
    ws['A1'] = "SiteIntel – By Kishor"
    ws['A1'].font = Font(size=16, bold=True)
    ws['A1'].alignment = Alignment(horizontal='center')

    headers = ["STREET ADDRESS 1", "STREET ADDRESS 2", "CITY", "STATE", "PIN CODE", "COUNTRY", "DATA SOURCE LINK"]
    ws.append(headers)

    for col in range(1, len(headers)+1):
        cell = ws.cell(row=2, column=col)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')

    for addr in address_list:
        ws.append([addr.get(h, "") for h in headers])

    ref = f"A2:G{len(address_list)+2}"
    table = Table(displayName="AddressTable", ref=ref)
    style = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
    table.tableStyleInfo = style
    ws.add_table(table)

    for col in ws.columns:
        max_length = max(len(str(cell.value)) for cell in col)
        ws.column_dimensions[col[0].column_letter].width = max_length + 2

    ws.freeze_panes = "A3"

    output = BytesIO()
    wb.save(output)
    return output.getvalue()

# -------------------------------
# Streamlit UI
# -------------------------------
st.set_page_config(
    page_title="SiteIntel – By Kishor",
    layout="wide",
    page_icon="📍",
    initial_sidebar_state="expanded"
)

# Add professional background
st.markdown(
    """
    <style>
    .stApp {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        background-attachment: fixed;
    }
    .stApp > div > div > div > div {
        background-color: rgba(255, 255, 255, 0.95);
        border-radius: 15px;
        padding: 25px;
        margin: 20px;
        box-shadow: 0 8px 16px rgba(0, 0, 0, 0.15);
    }
    .logo-container {
        text-align: center;
        margin-bottom: 30px;
    }
    .main-title {
        text-align: center;
        color: #2c3e50;
        font-size: 2.5em;
        font-weight: bold;
        margin-bottom: 10px;
    }
    .subtitle {
        text-align: center;
        color: #7f8c8d;
        font-size: 1.2em;
        margin-bottom: 40px;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# Logo and Title Section
st.markdown('<div class="logo-container">', unsafe_allow_html=True)
st.image("logo.png", width=400)
st.markdown('</div>', unsafe_allow_html=True)

st.markdown('<h1 class="main-title">📍 SiteIntel</h1>', unsafe_allow_html=True)
st.markdown('<p class="subtitle">Company Address Extraction & Standardization Tool</p>', unsafe_allow_html=True)

st.markdown("""
**How to use:**
1. Upload a CSV/Excel file with company data, or enter details manually
2. Click "Process" to extract and standardize addresses
3. View results in the table below
4. Download the Excel file with standardized addresses
""")

uploaded_file = st.file_uploader("Upload Company List (CSV/Excel)", type=["csv", "xlsx"])
company_name = st.text_input("Enter Company Name")
website = st.text_input("Enter Official Website")

# Results section
results_container = st.container()

if st.button("Process"):
    with results_container:
        st.info("Processing started...")
        companies = []
        if uploaded_file:
            df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)
            companies = [{"name": row["COMPANY NAME"], "website": row["OFFICIAL WEBSITE"]} for _, row in df.iterrows()]
        elif company_name and website:
            companies = [{"name": company_name, "website": website}]
        else:
            st.error("Please upload a file or enter company details.")
            st.stop()

        all_addresses = []
        progress = st.progress(0)
        for i, company in enumerate(companies):
            raw_address = extract_address(company["website"])
            standardized = standardize_address(raw_address)
            enriched = enrich_with_google_maps(standardized)
            enriched["DATA SOURCE LINK"] = company["website"]
            all_addresses.append(enriched)
            progress.progress((i+1)/len(companies))

        excel_data = generate_excel(all_addresses)
        st.success("Processing completed!")
        
        # Display results
        st.subheader("📊 Standardized Addresses")
        df_results = pd.DataFrame(all_addresses)
        st.dataframe(df_results, use_container_width=True)
        
        # Download button
        st.download_button(
            label="📥 Download Excel File",
            data=excel_data,
            file_name="SiteIntel_Output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="Click to download the standardized addresses as an Excel file"
        )

# Disclaimer section at the bottom
st.markdown("---")
st.markdown("""
### ⚠️ **Disclaimer**

**SiteIntel** is a tool for extracting and standardizing company address information from public websites. Please be aware of the following:

- **Data Accuracy**: While we strive for accuracy, the extracted information may not always be complete or up-to-date. Always verify critical information from official sources.
- **Web Scraping**: This tool scrapes public websites. Respect website terms of service and robots.txt files. Use responsibly and avoid overloading servers.
- **Google Maps API**: Address enrichment uses Google Maps Geocoding API. Usage is subject to Google's terms of service and may incur costs for high-volume usage.
- **Privacy & Legal**: Ensure you have proper authorization to collect and process company data. Comply with applicable data protection laws (GDPR, CCPA, etc.).
- **No Warranty**: This tool is provided "as is" without warranty of any kind. The developers are not liable for any damages arising from its use.
- **Contact**: For questions or concerns, please contact the developer.

**Last updated: January 2026**
""")

st.markdown('<div style="text-align: center; color: #666; font-size: 0.8em;">© 2026 SiteIntel - By Kishor</div>', unsafe_allow_html=True)
