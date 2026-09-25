<a id="readme-top"></a>

<div align="center">
  <h1>Document Orchestrator</h1>
  <p><strong>From Excel data to verified Word documents.</strong></p>
  <p>A FastAPI workflow that extracts structured spreadsheet data, maps it into a Word template, and coordinates review and revision through a document-editing service.</p>
  <p>
    <a href="#how-it-works"><strong>Explore the workflow »</strong></a>
    <br /><br />
    <a href="#run-locally">Run locally</a>
    &middot;
    <a href="https://github.com/pavan2184/excelwordtool-Orchestrator/issues/new">Report a bug</a>
    &middot;
    <a href="https://github.com/pavan2184/excelwordtool-Orchestrator/issues/new">Request a feature</a>
  </p>
</div>

<p align="center">
  <img src="https://img.shields.io/badge/Python-3776AB?logo=python&amp;logoColor=white" alt="Python" />
  <img src="https://img.shields.io/badge/FastAPI-009688?logo=fastapi&amp;logoColor=white" alt="FastAPI" />
  <img src="https://img.shields.io/badge/Excel-openpyxl-217346?logo=microsoftexcel&amp;logoColor=white" alt="openpyxl" />
  <img src="https://img.shields.io/badge/Word-python--docx-2B579A?logo=microsoftword&amp;logoColor=white" alt="python-docx" />
</p>

## About The Project

Document Orchestrator provides one browser workflow for uploading a `.docx` template and an `.xlsx` data source. It extracts fields and tables, scans the template, sends the structured data to a separate Word-editing service, exposes verification results, and supports revision before download.

## How It Works

1. Upload a Word template, Excel workbook, and filling instruction.
2. Extract normalized fields and tables with `openpyxl`.
3. Scan the template for candidate placeholders.
4. Send the template and structured report data to the Word-editing service.
5. Poll for completion, review verification issues, and submit revisions if needed.
6. Download the resulting document.

## Run Locally

Create a virtual environment and install the dependencies:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Start the companion Word-editing service on `http://localhost:8000`, then run the orchestrator on a different port:

```bash
uvicorn src.app:app --reload --port 8001
```

Open [http://localhost:8001](http://localhost:8001).
