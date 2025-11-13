# CSV Generator

Converts Excel sheets to CSV files. Built with Streamlit.

## Requirements

- Python 3.12+
- uv (or pip)

## Installation

```bash
uv sync
```

## Usage

```bash
uv run streamlit run streamlit_app.py
```

Access at http://localhost:8501

## Features

- Upload Excel file (.xlsx, .xlsm, .xls)
- Select sheets to convert
- Auto-removes unnamed columns from sheets containing "BU POS"
- Download individual CSVs or all as ZIP

## Deployment

### Streamlit Cloud

1. Push to GitHub
2. Connect repo at https://share.streamlit.io
3. Deploy

### Local Network

```bash
streamlit run streamlit_app.py --server.address 0.0.0.0
```

Access via http://[your-ip]:8501

## Output Format

- UTF-8 encoding
- All fields quoted
- No index column
