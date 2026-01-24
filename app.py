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
from openpyxl.utils import get_column_letter

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
            return address_tag.get_text(separator=", ")
        text = soup.get_text()
        # Look for address patterns in the text
        lines = text.split('\n')
        for line in lines:
            if any(keyword in line.upper() for keyword in ['STREET', 'ROAD', 'AVENUE', 'DRIVE', 'LANE', 'PLACE']):
                return line.strip()
        return "Address not found"
    except Exception as e:
        return f"Error extracting address: {str(e)}"

def standardize_address(address, city="", state="", country=""):
    address = unidecode(address).upper()

    # Standardize street abbreviations
    for abbr, full in short_forms.items():
        address = re.sub(r'\b' + abbr + r'\b', full, address)

    # Clean up extra spaces and punctuation
    address = re.sub(r'[^\w\s,.-]', '', address)
    address = re.sub(r'\s+', ' ', address).strip()

    return address

def enrich_with_google_maps(address):
    if not GOOGLE_MAPS_API_KEY or GOOGLE_MAPS_API_KEY == "YOUR_GOOGLE_MAPS_API_KEY":
        return address

    try:
        url = f"https://maps.googleapis.com/maps/api/geocode/json?address={address}&key={GOOGLE_MAPS_API_KEY}"
        response = requests.get(url)
        data = response.json()
        if data['status'] == 'OK':
            components = data['results'][0]['address_components']
            enriched = dict(address)

            for comp in components:
                if 'locality' in comp['types']:
                    enriched['CITY'] = comp['long_name'].upper()
                if 'administrative_area_level_1' in comp['types']:
                    state_name = comp['long_name'].upper()
                    enriched['STATE'] = us_states.get(state_name, state_name)
                if 'country' in comp['types']:
                    enriched['COUNTRY'] = standard_countries.get(comp['long_name'].upper(), comp['long_name'].upper())
                if 'postal_code' in comp['types']:
                    enriched['PIN CODE'] = comp['long_name']

            return enriched
    except Exception:
        pass
    return address

def generate_excel(address_list):
    wb = Workbook()
    ws = wb.active
    ws.title = "Standardized Addresses"

    # Title
    ws.merge_cells('A1:G1')
    ws['A1'] = "SiteIntel – By Kishor"
    ws['A1'].font = Font(size=16, bold=True)
    ws['A1'].alignment = Alignment(horizontal='center')

    # Headers
    headers = ["STREET ADDRESS 1", "STREET ADDRESS 2", "CITY", "STATE", "PIN CODE", "COUNTRY", "DATA SOURCE LINK"]
    ws.append(headers)

    # Style headers
    for col in range(1, len(headers)+1):
        cell = ws.cell(row=2, column=col)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')

    # Add data
    for addr in address_list:
        row_data = [
            addr.get("STREET ADDRESS 1", ""),
            addr.get("STREET ADDRESS 2", ""),
            addr.get("CITY", ""),
            addr.get("STATE", ""),
            addr.get("PIN CODE", ""),
            addr.get("COUNTRY", ""),
            addr.get("DATA SOURCE LINK", "")
        ]
        ws.append(row_data)

    # Create table
    ref = f"A2:G{len(address_list)+2}"
    table = Table(displayName="AddressTable", ref=ref)
    style = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
    table.tableStyleInfo = style
    ws.add_table(table)

    # Auto-adjust column widths
    for col_num in range(1, len(headers) + 1):
        col_letter = get_column_letter(col_num)
        max_length = 0
        for row in range(1, len(address_list) + 3):  # Include header rows
            cell = ws.cell(row=row, column=col_num)
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))

        # Set minimum width of 10, maximum of 50
        width = min(max(max_length + 2, 10), 50)
        ws.column_dimensions[col_letter].width = width

    # Freeze panes
    ws.freeze_panes = "A3"

    # Save to BytesIO
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

# Professional CSS with 3D effects and watermark
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

# Main UI
st.markdown("""
<div class="logo-container">
    <h1 style="color: #2c3e50; font-size: 3em; font-weight: bold; margin-bottom: 5px;">SiteIntel</h1>
    <p style="color: #7f8c8d; font-size: 1.2em; font-weight: 300;">By Kishor</p>
</div>
""", unsafe_allow_html=True)

st.markdown("---")

# File upload section
st.header("📤 Upload Company Data")
uploaded_file = st.file_uploader(
    "Upload Excel file with company websites",
    type=['xlsx', 'xls'],
    help="File should contain a column with company website URLs"
)

# Process button
if st.button("🚀 Process Addresses", type="primary"):
    if uploaded_file is not None:
        try:
            # Read the uploaded file
            df = pd.read_excel(uploaded_file)

            # Find column with URLs (look for common column names)
            url_column = None
            for col in df.columns:
                if any(keyword in col.lower() for keyword in ['website', 'url', 'link', 'site']):
                    url_column = col
                    break

            if url_column is None:
                # If no obvious URL column, use the first column
                url_column = df.columns[0]
                st.warning(f"No obvious URL column found. Using '{url_column}' as the website column.")

            st.success(f"Found {len(df)} companies to process")

            # Progress bar
            progress_bar = st.progress(0)
            status_text = st.empty()

            all_addresses = []

            for idx, row in df.iterrows():
                website = str(row[url_column]).strip()
                if website and website != 'nan':
                    status_text.text(f"Processing {idx+1}/{len(df)}: {website}")

                    # Extract address
                    raw_address = extract_address(website)

                    # Standardize address
                    standardized = standardize_address(raw_address)

                    # Create address dictionary
                    address_dict = {
                        "STREET ADDRESS 1": standardized,
                        "STREET ADDRESS 2": "",
                        "CITY": "",
                        "STATE": "",
                        "PIN CODE": "",
                        "COUNTRY": "",
                        "DATA SOURCE LINK": website
                    }

                    # Enrich with Google Maps (if API key available)
                    enriched_address = enrich_with_google_maps(address_dict)
                    all_addresses.append(enriched_address)

                progress_bar.progress((idx + 1) / len(df))

            progress_bar.empty()
            status_text.empty()

            if all_addresses:
                st.success(f"✅ Successfully processed {len(all_addresses)} addresses!")

                # Display results
                st.header("📊 Results Preview")
                result_df = pd.DataFrame(all_addresses)
                st.dataframe(result_df, use_container_width=True)

                # Download button
                excel_data = generate_excel(all_addresses)
                st.download_button(
                    label="📥 Download Excel Report",
                    data=excel_data,
                    file_name="standardized_addresses.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            else:
                st.error("No addresses could be processed. Please check your data.")

        except Exception as e:
            st.error(f"Error processing file: {str(e)}")
            st.info("Please ensure your Excel file contains valid website URLs.")

    else:
        st.warning("Please upload an Excel file first.")

# Footer
st.markdown("---")
st.markdown("""
<div style="text-align: center; color: #95a5a6; font-size: 0.9em;">
    <p><strong>Disclaimer:</strong> This tool extracts and standardizes address information from company websites for business intelligence purposes. Please ensure compliance with applicable laws and website terms of service.</p>
</div>
""", unsafe_allow_html=True)
