
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

## Agentic AI API

SiteIntel can also function as an **Agentic AI tool** through its REST API, allowing autonomous processing without user interaction.

### Running the API Server

```bash
python api.py
```

The API will be available at `http://localhost:8000`

### API Endpoints

- `GET /` - API information
- `GET /health` - Health check
- `POST /process-company` - Process single company
- `POST /process-batch` - Process multiple companies
- `POST /webhook-process` - Automatic file processing
- `POST /agentic-process` - Advanced agentic processing

### Example API Usage

```python
import requests

# Process single company
response = requests.post("http://localhost:8000/process-company", json={
    "name": "Example Corp",
    "website": "https://example.com"
})
print(response.json())

# Process batch
response = requests.post("http://localhost:8000/process-batch", json={
    "companies": [
        {"name": "Company A", "website": "https://companya.com"},
        {"name": "Company B", "website": "https://companyb.com"}
    ]
})
```

### Agentic Integration

The API can be integrated with:
- **AI Assistants** (ChatGPT plugins, Claude tools)
- **Workflow Automation** (Zapier, Make.com)
- **Other AI Agents** (LangChain, AutoGen)
- **Scheduled Tasks** (cron jobs, GitHub Actions)
- **Webhooks** (automatic processing on file upload)

## Requirements

- Python 3.11
- Streamlit
- Pandas
- Requests
- BeautifulSoup4
- OpenPyXL
- Unidecode
- FastAPI
- Uvicorn
- Pydantic

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
