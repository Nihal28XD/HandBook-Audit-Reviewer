# 📊 RBI Handbook Enterprise Audit & Verification Suite

An enterprise-grade cross-validation and auditing engine designed to verify multi-sheet financial Excel workbooks against final published PDF documents (such as RBI Handbook publications) with native PDF highlight annotations and an interactive side-by-side workspace.

---

## 📌 Overview

During publication of large regulatory and statistical handbooks, figures and text transcribed from raw spreadsheets to final layout PDFs can suffer from typesetting errors, rounding shifts, case anomalies, or omitted values. 

The **RBI Handbook Enterprise Audit & Verification Suite** automates this reconciliation process. It extracts and normalizes textual and numeric records across every worksheet in an Excel workbook, cross-references them against the extracted text and character coordinates of the publication PDF, injects native yellow highlight annotations at the exact error coordinates in the PDF, and presents an interactive 50/50 side-by-side comparison workspace.

---

## ✨ Key Features (Based on Implementation)

- **Multi-Sheet Excel Parsing & Coordinate Extraction:** Recursively audits all sheets in an uploaded `.xlsx` workbook, extracting textual labels, financial figures, table headers, and exact cell coordinates (`Sheet: <name> | Row: <idx>`).
- **PDF Text Extraction & Dehyphenation:** Processes publication PDFs page by page with `pdfplumber`, dehyphenating wrapped words and deduplicating characters for accurate visual text alignment.
- **Native PDF Highlight Injection:** Uses `pypdf` annotation primitives (`Highlight`, `ArrayObject`, `FloatObject`) to inject yellow highlights directly into the PDF at exact quad points (`x0`, `top`, `x1`, `bottom`), producing an annotated PDF ready for download or in-browser viewing.
- **Multi-Vector Discrepancy Auditing:**
  - **`#` Character Detection:** Identifies unexpanded Excel column overflows (`###`) rendered in the PDF.
  - **Unrecognized / Extra PDF Words:** Surfaces vocabulary present in the PDF that does not exist in the source spreadsheet (accounting for publication stopwords).
  - **Table / Sheet Title Verification:** Matches sheet and table headers against PDF headings using normalized token overlap.
  - **Case Sensitivity Discrepancies:** Detects case divergences (e.g., uppercase vs lowercase) and maps them to specific PDF pages.
  - **Financial Rounding & Precision Variants:** Evaluates numeric values against exact floats, 2-decimal, 1-decimal, half-up rounded integers (`>= 0.5`), and comma-formatted thousands (`12,500`).
- **Enterprise Audit Dashboard:** Instant executive summary showing Total Cells Audited, Publication Pages, Total Discrepancies Flagged, and an Audit Integrity Score (`100% - errors * 0.5%`).
- **Interactive 50/50 Side-by-Side Comparison Workspace:**
  - **Left Pane:** Embedded viewer showing the natively highlighted PDF.
  - **Right Pane:** Full-featured interactive spreadsheet powered by Handsontable, complete with ribbon tabs, formula bar, cell inspection, and worksheet tab navigation.
  - **Export Options:** Open in a separate browser tab or download as a standalone offline HTML artifact (`RBI_Interactive_Comparison_Workspace.html`).
- **Zero Disk Footprint:** All parsing, coordinate mapping, and PDF annotation occur in-memory using `io.BytesIO` streams, making it safe for ephemeral cloud environments like Render.

---

## 🛠️ Technology Stack

- **Application Framework:** [Streamlit](https://streamlit.io/)
- **Data Processing:** [Pandas](https://pandas.pydata.org/), [OpenPyXL](https://openpyxl.readthedocs.io/)
- **PDF Extraction & Geometry:** [pdfplumber](https://github.com/jsvine/pdfplumber)
- **PDF Annotation Engine:** [pypdf](https://pypdf.readthedocs.io/)
- **Interactive UI Components:** [Streamlit Components API](https://docs.streamlit.io/develop/api-reference/custom-components), [Handsontable](https://handsontable.com/)
- **Deployment Platform:** [Render](https://render.com/) (Web Service with Python runtime)

---

## 📁 Project Structure

```text
RBI_analyser/
├── .streamlit/
│   └── config.toml          # Streamlit headless server and upload size configuration
├── .gitignore               # Comprehensive Git exclusions (caches, venvs, data files)
├── .python-version          # Pinned Python version (3.11.9) for Render/pyenv
├── app.py                   # Production deployment entry point
├── pdf_analyser_Final.py    # Core application engine and Streamlit UI
├── render.yaml              # Render Blueprint deployment specification
├── requirements.txt         # Pinned production runtime dependencies
└── README.md                # Project documentation and setup guide
```

---

## 📋 Prerequisites

- **Python 3.10 or 3.11** recommended (Python 3.8+ supported locally).
- **pip** package installer.
- **Git** installed on your system.

---

## 💻 Local Setup Instructions

### 1. Clone or Open the Project Folder

Open your terminal or IDE (e.g., VS Code) in the root of this project folder (`RBI_analyser`).

### 2. Create and Activate a Virtual Environment

#### On Windows (PowerShell):
```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

> *Note for PowerShell:* If script execution is restricted, run:
> ```powershell
> Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process
> ```

#### On Windows (Command Prompt):
```cmd
python -m venv .venv
.venv\Scripts\activate.bat
```

#### On Linux / macOS:
```bash
python3 -m venv .venv
source .venv/bin/activate
```

### 3. Upgrade pip and Install Dependencies

```bash
python -m pip install --upgrade pip
pip install -r requirements.txt
```

### 4. Run the Application Locally

```bash
streamlit run app.py
```

Streamlit will start and automatically open your default web browser at `http://localhost:8501`.

---

## 🚀 GitHub Repository Initialization & Push Instructions

This folder is configured and ready to be pushed to GitHub. Follow these exact steps:

### Step 1: Initialize Git and Stage Files
```bash
git init
git add .
git commit -m "feat: initial commit for RBI Handbook Enterprise Audit Suite"
```

### Step 2: Set Branch to `main`
```bash
git branch -M main
```

### Step 3: Link Your Remote Repository and Push
Replace `<YOUR_GITHUB_USERNAME>` and `<YOUR_REPOSITORY_NAME>` with your actual repository details:
```bash
git remote add origin https://github.com/<YOUR_GITHUB_USERNAME>/<YOUR_REPOSITORY_NAME>.git
git push -u origin main
```

> **Security Note:** `.gitignore` is pre-configured to exclude virtual environments (`.venv/`), temporary files, secrets, uploaded workbooks (`*.xlsx`, `*.xls`), and publication PDFs (`*.pdf`). Never bypass `.gitignore` to commit confidential audit documents or credentials.

---

## ☁️ Render Deployment Instructions

This repository includes a `render.yaml` Blueprint file for 1-click or automated deployment on [Render](https://render.com/).

### Method A: Deploy with Render Blueprint (Recommended)

1. Log in to your [Render Dashboard](https://dashboard.render.com/).
2. Click **New +** and select **Blueprint**.
3. Connect your GitHub repository.
4. Render will automatically detect `render.yaml` and configure:
   - **Service Name:** `rbi-handbook-analyser`
   - **Environment:** `python`
   - **Python Version:** `3.11.9`
   - **Build Command:** `pip install --upgrade pip && pip install -r requirements.txt`
   - **Start Command:** `streamlit run app.py --server.port $PORT --server.address 0.0.0.0`
   - **Health Check Path:** `/_stcore/health`
5. Click **Apply**. Render will build and deploy your application.

### Method B: Manual Web Service Setup

If you prefer setting up manually without Blueprints:
1. Click **New +** > **Web Service**.
2. Select your connected GitHub repository.
3. Configure settings:
   - **Name:** `rbi-handbook-analyser`
   - **Runtime:** `Python`
   - **Branch:** `main`
   - **Build Command:** `pip install --upgrade pip && pip install -r requirements.txt`
   - **Start Command:** `streamlit run app.py --server.port $PORT --server.address 0.0.0.0`
   - **Health Check Path:** `/_stcore/health`
4. Under **Environment Variables**, add:
   - `PYTHON_VERSION` = `3.11.9`
5. Click **Create Web Service**.

---

## 📖 How to Use the Application

1. Open the application URL in a web browser.
2. In the left sidebar:
   - Click **Upload Excel Workbook** and select your `.xlsx` source file.
   - Click **Upload Publication PDF** and select your target `.pdf` publication.
   - (Optional) Toggle **Strict Case Sensitivity** or **0.5 Financial Rounding Logic**.
3. The engine automatically processes the documents:
   - View top-level audit metrics and integrity percentage.
   - Review categorized discrepancy tables (Case Mismatches, Sheet Name Mismatches, Missing Numbers/Words, Unrecognized PDF Words).
4. Launch the side-by-side workspace:
   - Click **🖥️ Open 50/50 Side-by-Side Workspace in New Tab** to inspect the highlighted PDF next to the live spreadsheet grid.
   - Or click **⬇️ Download Interactive 50/50 Workspace (HTML)** to save the standalone comparison report for offline review.

---

## 📄 Supported Input Formats and Limitations

- **Excel Workbooks:** `.xlsx` format (OpenPyXL engine). Multiple sheets are supported. Password-protected workbooks or legacy `.xls` binary files must be converted to standard `.xlsx`.
- **Publication PDFs:** Native text-based `.pdf` files. Scanned images without an underlying OCR text layer cannot be parsed by character coordinate extractors.
- **Upload File Limits:** Streamlit is configured for file uploads up to 200MB (`maxUploadSize = 200` in `.streamlit/config.toml`). For Render free-tier instances (512MB RAM), keep uploaded files within reasonable memory bounds to avoid OOM limits.
- **Browser Workspace:** The interactive spreadsheet pane in the workspace is an exploratory visual grid (Handsontable); macros and active formula recalculations are not executed in the browser.

---

## 🔧 Troubleshooting

| Issue | Cause | Solution |
| :--- | :--- | :--- |
| `ImportError: DLL load failed while importing cd` | Python 3.8 on Windows loading incompatible `charset-normalizer` wheel | Run `pip install "charset-normalizer>=3.4.0,<3.5.0"`. `requirements.txt` already handles this marker. |
| Popup blocked when opening 50/50 workspace | Browser popup blocker preventing `window.open` | Allow popups for the site or click the **Download Interactive 50/50 Workspace (HTML)** fallback button. |
| Streamlit port error on Render | App attempting to bind to a hardcoded port | Ensure start command uses `--server.port $PORT --server.address 0.0.0.0`. `render.yaml` specifies this automatically. |
| Memory limit exceeded on Render | Very large PDFs (>500 pages) exceeding free-tier RAM | Split large multi-chapter publications into individual chapter PDFs or upgrade instance RAM on Render. |
