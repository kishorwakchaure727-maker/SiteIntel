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
    from openpyxl.utils import get_column_letter

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

    # Auto-adjust column widths
    for col_num, col in enumerate(ws.columns, 1):
        if col:  # Check if column has any cells
            try:
                # Get all non-None cell values in this column
                cell_values = [str(cell.value) if cell.value is not None else "" for cell in col]
                if cell_values:  # Check if we have any values
                    max_length = max(len(val) for val in cell_values)
                    # Ensure minimum width of 10 and maximum of 50
                    width = min(max(max_length + 2, 10), 50)
                    # Use column index to get column letter (A=1, B=2, etc.)
                    ws.column_dimensions[get_column_letter(col_num)].width = width
            except (AttributeError, IndexError, TypeError):
                # Fallback: set default width if calculation fails
                ws.column_dimensions[get_column_letter(col_num)].width = 15

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

# Add professional 3D background and effects
st.markdown(
    """
    <style>
    .stApp {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        background-attachment: fixed;
        perspective: 1000px;
        position: relative;
    }

    /* Supply Chain World Map Watermark */
    .stApp::before {
        content: '';
        position: absolute;
        top: 20px;
        left: 50%;
        transform: translateX(-50%);
        width: 400px;
        height: 200px;
        background-image:
            radial-gradient(circle at 30% 40%, rgba(255,255,255,0.1) 1px, transparent 1px),
            radial-gradient(circle at 35% 35%, rgba(255,255,255,0.08) 1px, transparent 1px),
            radial-gradient(circle at 40% 45%, rgba(255,255,255,0.06) 1px, transparent 1px),
            radial-gradient(circle at 45% 30%, rgba(255,255,255,0.05) 1px, transparent 1px),
            radial-gradient(circle at 50% 50%, rgba(255,255,255,0.07) 1px, transparent 1px),
            radial-gradient(circle at 55% 35%, rgba(255,255,255,0.06) 1px, transparent 1px),
            radial-gradient(circle at 60% 40%, rgba(255,255,255,0.08) 1px, transparent 1px),
            radial-gradient(circle at 65% 45%, rgba(255,255,255,0.05) 1px, transparent 1px),
            radial-gradient(circle at 70% 35%, rgba(255,255,255,0.06) 1px, transparent 1px);
        background-size: 50px 50px, 60px 60px, 45px 45px, 55px 55px, 40px 40px, 65px 65px, 35px 35px, 70px 70px, 50px 50px;
        background-position:
            10% 20%, 15% 15%, 20% 25%, 25% 10%, 30% 30%, 35% 15%, 40% 20%, 45% 25%, 50% 15%;
        opacity: 0.3;
        z-index: 0;
        pointer-events: none;
        filter: blur(0.5px);
    }

    /* Supply Chain Text Overlay */
    .stApp::after {
        content: 'SUPPLY CHAIN WORLD MAP';
        position: absolute;
        top: 40px;
        left: 50%;
        transform: translateX(-50%);
        font-size: 14px;
        font-weight: 300;
        color: rgba(255, 255, 255, 0.4);
        letter-spacing: 2px;
        text-shadow: 0 1px 2px rgba(0, 0, 0, 0.3);
        z-index: 0;
        pointer-events: none;
        font-family: 'Arial', sans-serif;
    }

    .stApp > div > div > div > div {
        background-color: rgba(255, 255, 255, 0.95);
        border-radius: 15px;
        padding: 25px;
        margin: 20px;
        box-shadow:
            0 10px 30px rgba(0, 0, 0, 0.2),
            0 1px 8px rgba(0, 0, 0, 0.1),
            inset 0 1px 0 rgba(255, 255, 255, 0.6);
        transform: translateZ(0);
        transition: all 0.3s ease;
        position: relative;
        z-index: 1;
    }

    .stApp > div > div > div > div:hover {
        transform: translateY(-5px) rotateX(2deg);
        box-shadow:
            0 20px 40px rgba(0, 0, 0, 0.25),
            0 5px 15px rgba(0, 0, 0, 0.15),
            inset 0 1px 0 rgba(255, 255, 255, 0.8);
    }

    /* 3D Button Effects */
    .stButton > button {
        background: linear-gradient(145deg, #667eea, #764ba2);
        border: none;
        border-radius: 10px;
        padding: 12px 24px;
        color: white;
        font-weight: bold;
        text-shadow: 0 1px 2px rgba(0, 0, 0, 0.3);
        box-shadow:
            0 4px 15px rgba(102, 126, 234, 0.4),
            0 2px 8px rgba(0, 0, 0, 0.1),
            inset 0 1px 0 rgba(255, 255, 255, 0.2);
        transition: all 0.3s ease;
        transform: translateZ(0);
    }

    .stButton > button:hover {
        transform: translateY(-2px) scale(1.05);
        box-shadow:
            0 8px 25px rgba(102, 126, 234, 0.6),
            0 4px 12px rgba(0, 0, 0, 0.2),
            inset 0 1px 0 rgba(255, 255, 255, 0.3);
    }

    .stButton > button:active {
        transform: translateY(0) scale(0.98);
        box-shadow:
            0 2px 8px rgba(102, 126, 234, 0.4),
            inset 0 2px 4px rgba(0, 0, 0, 0.1);
    }

    /* 3D Card Effects for DataFrames */
    .stDataFrame {
        border-radius: 12px;
        overflow: hidden;
        box-shadow:
            0 8px 25px rgba(0, 0, 0, 0.15),
            0 2px 10px rgba(0, 0, 0, 0.1);
        transform: translateZ(0);
        transition: all 0.3s ease;
    }

    .stDataFrame:hover {
        transform: translateY(-3px);
        box-shadow:
            0 15px 35px rgba(0, 0, 0, 0.2),
            0 5px 15px rgba(0, 0, 0, 0.15);
    }

    /* 3D Progress Bar */
    .stProgress > div > div {
        background: linear-gradient(90deg, #667eea, #764ba2);
        border-radius: 10px;
        box-shadow: inset 0 2px 4px rgba(0, 0, 0, 0.2);
    }

    /* 3D Logo Effect */
    .logo-container {
        text-align: center;
        margin-bottom: 30px;
        transform: translateZ(0);
    }

    .logo-container img {
        transition: all 0.3s ease;
        filter: drop-shadow(0 4px 8px rgba(0, 0, 0, 0.2));
    }

    .logo-container img:hover {
        transform: scale(1.05) rotateY(5deg);
        filter: drop-shadow(0 8px 16px rgba(0, 0, 0, 0.3));
    }

    /* 3D Input Fields */
    .stTextInput > div > div > input,
    .stFileUploader > div > div > div {
        border-radius: 8px;
        border: 2px solid #e1e5e9;
        box-shadow:
            inset 0 2px 4px rgba(0, 0, 0, 0.1),
            0 2px 8px rgba(0, 0, 0, 0.05);
        transition: all 0.3s ease;
    }

    .stTextInput > div > div > input:focus,
    .stFileUploader > div > div > div:focus-within {
        border-color: #667eea;
        box-shadow:
            inset 0 2px 4px rgba(0, 0, 0, 0.1),
            0 0 0 3px rgba(102, 126, 234, 0.1),
            0 4px 12px rgba(102, 126, 234, 0.2);
        transform: translateY(-1px);
    }

    /* 3D Success/Error Messages */
    .stSuccess, .stError, .stInfo {
        border-radius: 10px;
        border: none;
        box-shadow:
            0 4px 12px rgba(0, 0, 0, 0.1),
            inset 0 1px 0 rgba(255, 255, 255, 0.3);
        transform: translateZ(0);
        transition: all 0.3s ease;
    }

    .stSuccess:hover, .stError:hover, .stInfo:hover {
        transform: translateY(-2px);
        box-shadow:
            0 8px 20px rgba(0, 0, 0, 0.15),
            inset 0 1px 0 rgba(255, 255, 255, 0.4);
    }

    /* 3D Download Button Special Effect */
    [data-testid="stDownloadButton"] > button {
        background: linear-gradient(145deg, #28a745, #20c997);
        position: relative;
        overflow: hidden;
    }

    [data-testid="stDownloadButton"] > button::before {
        content: '';
        position: absolute;
        top: 0;
        left: -100%;
        width: 100%;
        height: 100%;
        background: linear-gradient(90deg, transparent, rgba(255, 255, 255, 0.3), transparent);
        transition: left 0.5s;
    }

    [data-testid="stDownloadButton"] > button:hover::before {
        left: 100%;
    }

    /* Subtle animations */
    @keyframes float {
        0%, 100% { transform: translateY(0px); }
        50% { transform: translateY(-5px); }
    }

    .logo-container img {
        animation: float 3s ease-in-out infinite;
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
