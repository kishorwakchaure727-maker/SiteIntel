
# SiteIntel

A Streamlit application for extracting and standardizing company addresses from websites.

## Features

- Extract addresses from company websites
- Standardize address formats
- Enrich with Google Maps API
- Export to Excel

## Installation

1. Clone the repository
2. Install dependencies: `pip install -r requirements.txt`
3. Run the app: `streamlit run app.py`

## Usage

Upload a CSV/Excel file with company names and websites, or enter manually.

## Requirements

- Python 3.11
- Streamlit
- Pandas
- Requests
- BeautifulSoup4
- OpenPyXL
- Unidecode

## Deployment on Streamlit Cloud

1. Push your code to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Connect your GitHub account
4. Select the repository `SiteIntel`
5. Set the main file path to `app.py`
6. Add your Google Maps API key in the app secrets (GOOGLE_MAPS_API_KEY)
7. Deploy

## License

MIT
