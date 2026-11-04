import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import pdfplumber
import pypdf
from pypdf.annotations import Highlight
from pypdf.generic import ArrayObject, FloatObject
import re
import openpyxl
import datetime
import math
import io
import base64
import json
import html

# ─── STREAMLIT PAGE CONFIG & ENTERPRISE LIGHT THEME ───
st.set_page_config(
    page_title="RBI Handbook Enterprise Audit & Verification Suite",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
    /* ─── Modern Light Enterprise Palette ─── */
    :root {
        --main-bg: #F5F7FB;
        --sidebar-bg: #FFFFFF;
        --card-bg: #FFFFFF;
        --primary-btn: #2563EB;
        --primary-btn-text: #FFFFFF;
        --primary-btn-hover: #1D4ED8;
        --secondary-btn: #E8EEF8;
        --secondary-btn-text: #1E293B;
        --secondary-btn-hover: #D9E2EF;
        --text-main: #172B4D;
        --text-secondary: #475569;
        --border-color: #D9E2EF;
        --border-hover: #CBD5E1;
        --table-header: #E2E8F0;
        --table-header-text: #172B4D;
        --table-header-border: #CBD5E1;
        --table-body-text: #1E293B;
        --table-row: #FFFFFF;
        --table-row-alt: #F8FAFC;
        --table-border: #CBD5E1;
        --table-row-hover: #EDF4FF;
        --warning-bg: #FFF0B3;
        --warning-text: #78350F;
        --warning-border: #FDE68A;
        --warning-icon: #B45309;
        --error-bg: #FCE4E4;
        --error-text: #991B1B;
        --error-border: #FECACA;
        --error-icon: #DC2626;
        --success-bg: #DCFCE7;
        --success-text: #166534;
        --info-bg: #EFF6FF;
        --info-text: #1E3A8A;
    }

    /* ─── Global Background & Typography ─── */
    .stApp {
        background-color: var(--main-bg) !important;
        color: var(--text-main) !important;
    }

    /* ─── Sidebar Styling ─── */
    [data-testid="stSidebar"] {
        background-color: var(--sidebar-bg) !important;
        border-right: 1px solid var(--border-color) !important;
    }
    [data-testid="stSidebar"] h1,
    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3,
    [data-testid="stSidebar"] h4 {
        color: var(--text-main) !important;
        font-weight: 700 !important;
    }
    [data-testid="stSidebar"] p,
    [data-testid="stSidebar"] span,
    [data-testid="stSidebar"] label {
        color: var(--text-secondary) !important;
    }

    /* ─── Headings & Body Text ─── */
    h1, h2, h3, h4, h5, h6 {
        color: var(--text-main) !important;
        font-weight: 700 !important;
    }
    p, span, label {
        color: var(--text-secondary);
    }
    hr {
        border-color: var(--border-color) !important;
    }

    /* ─── File Uploader Section ─── */
    [data-testid="stFileUploader"] {
        background-color: var(--card-bg) !important;
        border: 1px solid var(--border-color) !important;
        border-radius: 8px !important;
        padding: 0.85rem !important;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.03) !important;
    }

    [data-testid="stFileUploader"] label {
        color: var(--text-main) !important;
        font-weight: 600 !important;
        font-size: 0.9rem !important;
    }

    /* File Uploader Dropzone */
    [data-testid="stFileUploader"] section,
    [data-testid="stFileUploaderDropzone"],
    [data-testid="stFileUploadDropzone"] {
        background-color: #F8FAFC !important;
        border: 2px dashed var(--border-color) !important;
        border-radius: 8px !important;
        padding: 1rem !important;
        transition: all 0.15s ease !important;
    }

    [data-testid="stFileUploader"] section:hover,
    [data-testid="stFileUploaderDropzone"]:hover,
    [data-testid="stFileUploadDropzone"]:hover {
        border-color: var(--primary-btn) !important;
        background-color: #F1F5F9 !important;
    }

    /* File uploader instructions text */
    [data-testid="stFileUploader"] section span,
    [data-testid="stFileUploader"] section small,
    [data-testid="stFileUploader"] section div {
        color: var(--text-secondary) !important;
        font-weight: 500 !important;
    }

    /* Browse Files Button */
    [data-testid="stFileUploader"] section button,
    [data-testid="stFileUploader"] [data-testid="baseButton-secondary"] {
        background-color: var(--secondary-btn) !important;
        color: var(--secondary-btn-text) !important;
        border: 1px solid var(--border-color) !important;
        border-radius: 6px !important;
        font-weight: 600 !important;
        font-size: 0.875rem !important;
        padding: 0.45rem 1rem !important;
        box-shadow: 0 1px 2px rgba(0, 0, 0, 0.04) !important;
        transition: all 0.15s ease !important;
    }

    [data-testid="stFileUploader"] section button:hover,
    [data-testid="stFileUploader"] [data-testid="baseButton-secondary"]:hover {
        background-color: var(--secondary-btn-hover) !important;
        color: var(--text-main) !important;
        border-color: var(--primary-btn) !important;
    }

    /* Uploaded File Item Details */
    [data-testid="stFileUploader"] ul {
        background-color: var(--card-bg) !important;
        border: 1px solid var(--border-color) !important;
        border-radius: 6px !important;
        margin-top: 0.5rem !important;
        padding: 0.4rem 0.65rem !important;
    }

    [data-testid="stFileUploader"] li,
    [data-testid="stFileUploader"] li div,
    [data-testid="stFileUploader"] li span {
        color: var(--text-main) !important;
        font-weight: 500 !important;
    }

    [data-testid="stFileUploader"] li small {
        color: var(--text-secondary) !important;
    }

    /* Remove File Button */
    [data-testid="stFileUploader"] button[aria-label="Delete file"],
    [data-testid="stFileUploader"] button[data-testid="stFileUploaderDeleteBtn"] {
        background: transparent !important;
        border: none !important;
        color: var(--error-text) !important;
    }

    [data-testid="stFileUploader"] button[aria-label="Delete file"]:hover,
    [data-testid="stFileUploader"] button[data-testid="stFileUploaderDeleteBtn"]:hover {
        background-color: var(--error-bg) !important;
        color: var(--error-text) !important;
        border-radius: 4px !important;
    }

    /* ─── Checkbox Inputs ─── */
    [data-testid="stCheckbox"] label span {
        color: var(--text-main) !important;
        font-weight: 500 !important;
    }

    /* ─── Metric Cards ─── */
    [data-testid="stMetric"] {
        background-color: var(--card-bg) !important;
        border: 1px solid var(--border-color) !important;
        border-radius: 8px !important;
        padding: 1.15rem 1.25rem !important;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.04) !important;
    }

    [data-testid="stMetricLabel"] p,
    [data-testid="stMetricLabel"] div {
        color: var(--text-secondary) !important;
        font-weight: 600 !important;
        font-size: 0.8rem !important;
        text-transform: uppercase !important;
        letter-spacing: 0.03em !important;
    }

    [data-testid="stMetricValue"] div {
        color: var(--text-main) !important;
        font-weight: 800 !important;
        font-size: 1.85rem !important;
    }

    /* ─── Standard Buttons ─── */
    div.stButton > button {
        background-color: var(--primary-btn) !important;
        color: var(--primary-btn-text) !important;
        font-weight: 600 !important;
        border-radius: 6px !important;
        padding: 0.6rem 1.25rem !important;
        border: 1px solid var(--primary-btn) !important;
        transition: all 0.15s ease !important;
        box-shadow: 0 1px 3px rgba(37, 99, 235, 0.25) !important;
    }

    div.stButton > button:hover {
        background-color: var(--primary-btn-hover) !important;
        color: var(--primary-btn-text) !important;
        border-color: var(--primary-btn-hover) !important;
        box-shadow: 0 2px 6px rgba(29, 78, 216, 0.35) !important;
    }

    div.stButton > button:focus {
        outline: 2px solid var(--primary-btn) !important;
        outline-offset: 2px !important;
    }

    /* ─── Download Buttons ─── */
    div.stDownloadButton > button {
        background-color: var(--secondary-btn) !important;
        color: var(--secondary-btn-text) !important;
        font-weight: 600 !important;
        border-radius: 6px !important;
        padding: 0.65rem 1.25rem !important;
        border: 1px solid var(--border-color) !important;
        transition: all 0.15s ease !important;
    }

    div.stDownloadButton > button:hover {
        background-color: var(--secondary-btn-hover) !important;
        color: var(--text-main) !important;
        border-color: var(--primary-btn) !important;
    }

    div.stDownloadButton > button:focus {
        outline: 2px solid var(--primary-btn) !important;
        outline-offset: 2px !important;
    }

    /* ─── Tables & DataFrames ─── */
    .audit-table-wrapper {
        width: 100% !important;
        margin-top: 6px !important;
        margin-bottom: 24px !important;
        border: 1px solid var(--table-border) !important;
        border-radius: 6px !important;
        overflow-x: auto !important;
        background-color: var(--table-row) !important;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.03) !important;
    }

    table,
    table.audit-table,
    [data-testid="stTable"] table {
        color: var(--table-body-text) !important;
        background-color: var(--table-row) !important;
        border-collapse: collapse !important;
        width: 100% !important;
        text-align: left !important;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif !important;
    }

    th,
    table.audit-table th,
    [data-testid="stTable"] th {
        background-color: var(--table-header) !important;
        color: var(--table-header-text) !important;
        font-weight: 700 !important;
        border: 1px solid var(--table-header-border) !important;
        padding: 10px 14px !important;
        font-size: 0.875rem !important;
        letter-spacing: 0.01em !important;
        text-align: left !important;
        white-space: nowrap !important;
    }

    td,
    table.audit-table td,
    [data-testid="stTable"] td {
        color: var(--table-body-text) !important;
        background-color: var(--table-row) !important;
        border: 1px solid var(--table-border) !important;
        padding: 9px 14px !important;
        font-size: 0.875rem !important;
        vertical-align: top !important;
        line-height: 1.45 !important;
    }

    tr:nth-child(even) td,
    table.audit-table tr:nth-child(even) td,
    [data-testid="stTable"] tr:nth-child(even) td {
        background-color: var(--table-row-alt) !important;
    }

    tr:hover td,
    table.audit-table tr:hover td,
    .audit-row:hover td,
    [data-testid="stTable"] tr:hover td {
        background-color: var(--table-row-hover) !important;
    }

    [data-testid="stDataFrame"],
    [data-testid="stDataFrameResizable"] {
        border: 1px solid var(--table-border) !important;
        border-radius: 6px !important;
        background-color: var(--card-bg) !important;
        width: 100% !important;
        margin-top: 6px !important;
        margin-bottom: 24px !important;
    }

    /* ─── Alerts & Status Messages ─── */
    [data-testid="stAlert"] {
        border-radius: 6px !important;
        padding: 0 !important;
        margin-bottom: 0px !important;
    }

    /* Red Discrepancy Heading (Error) */
    div.stAlert:has([data-testid="stNotificationContentError"]),
    div[data-testid="stAlert"]:has([data-testid="stNotificationContentError"]),
    div[data-testid="stAlert"] div[data-baseweb="notification"][kind="negative"],
    div.stAlert div[data-baseweb="notification"][kind="negative"],
    div[data-baseweb="notification"][kind="negative"],
    div.stAlert:has([aria-label="error"]) {
        background-color: var(--error-bg) !important;
        color: var(--error-text) !important;
        border: 1px solid var(--error-border) !important;
        border-radius: 6px !important;
        padding: 10px 16px !important;
        margin: 0 !important;
    }

    div[data-baseweb="notification"][kind="negative"] p,
    div[data-baseweb="notification"][kind="negative"] span,
    div[data-baseweb="notification"][kind="negative"] div,
    div[data-baseweb="notification"][kind="negative"] [data-testid="stMarkdownContainer"] p,
    div[data-testid="stNotificationContentError"] p,
    div[data-testid="stNotificationContentError"] span,
    div[data-testid="stNotificationContentError"] div {
        color: var(--error-text) !important;
        font-weight: 600 !important;
        font-size: 0.95rem !important;
    }

    div[data-baseweb="notification"][kind="negative"] svg,
    div[data-testid="stAlert"]:has([data-testid="stNotificationContentError"]) svg,
    div[data-testid="stNotificationContentError"] svg,
    div[data-testid="stAlertDynamicIcon"] svg,
    div[data-baseweb="notification"][kind="negative"] [data-testid="stAlertDynamicIcon"] {
        color: var(--error-icon) !important;
        fill: var(--error-icon) !important;
    }

    /* Yellow Warning Heading */
    div.stAlert:has([data-testid="stNotificationContentWarning"]),
    div[data-testid="stAlert"]:has([data-testid="stNotificationContentWarning"]),
    div[data-testid="stAlert"] div[data-baseweb="notification"][kind="warning"],
    div.stAlert div[data-baseweb="notification"][kind="warning"],
    div[data-baseweb="notification"][kind="warning"],
    div.stAlert:has([aria-label="warning"]) {
        background-color: var(--warning-bg) !important;
        color: var(--warning-text) !important;
        border: 1px solid var(--warning-border) !important;
        border-radius: 6px !important;
        padding: 10px 16px !important;
        margin: 0 !important;
    }

    div[data-baseweb="notification"][kind="warning"] p,
    div[data-baseweb="notification"][kind="warning"] span,
    div[data-baseweb="notification"][kind="warning"] div,
    div[data-baseweb="notification"][kind="warning"] [data-testid="stMarkdownContainer"] p,
    div[data-testid="stNotificationContentWarning"] p,
    div[data-testid="stNotificationContentWarning"] span,
    div[data-testid="stNotificationContentWarning"] div {
        color: var(--warning-text) !important;
        font-weight: 600 !important;
        font-size: 0.95rem !important;
    }

    div[data-baseweb="notification"][kind="warning"] svg,
    div[data-testid="stAlert"]:has([data-testid="stNotificationContentWarning"]) svg,
    div[data-testid="stNotificationContentWarning"] svg,
    div[data-testid="stAlertDynamicIcon"] svg,
    div[data-baseweb="notification"][kind="warning"] [data-testid="stAlertDynamicIcon"] {
        color: var(--warning-icon) !important;
        fill: var(--warning-icon) !important;
    }

    /* Success alert (pale green background with dark green text) */
    div.stAlert:has([data-testid="stNotificationContentSuccess"]),
    div[data-testid="stAlert"]:has([data-testid="stNotificationContentSuccess"]),
    div[data-baseweb="notification"][kind="positive"] {
        background-color: var(--success-bg) !important;
        color: var(--success-text) !important;
        border: 1px solid #BBF7D0 !important;
        border-radius: 6px !important;
        padding: 10px 16px !important;
    }

    /* Info alert (pale blue background with dark blue text) */
    div.stAlert:has([data-testid="stNotificationContentInfo"]),
    div[data-testid="stAlert"]:has([data-testid="stNotificationContentInfo"]),
    div[data-baseweb="notification"][kind="info"] {
        background-color: var(--info-bg) !important;
        color: var(--info-text) !important;
        border: 1px solid #BFDBFE !important;
        border-radius: 6px !important;
        padding: 10px 16px !important;
    }

    /* Ensure text inside alert boxes inherits correct contrast color */
    [data-testid="stAlert"] p,
    [data-testid="stAlert"] span,
    [data-testid="stAlert"] div {
        color: inherit !important;
        font-weight: 500 !important;
    }

    /* ─── Tooltips ─── */
    div[data-baseweb="tooltip"],
    div[role="tooltip"] {
        background-color: #1E293B !important;
        color: #FFFFFF !important;
        border-radius: 6px !important;
        font-size: 0.8rem !important;
        padding: 6px 10px !important;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.15) !important;
    }

    div[data-baseweb="tooltip"] * {
        color: #FFFFFF !important;
    }

    div[data-baseweb="popover"] {
        background-color: #FFFFFF !important;
        color: var(--text-main) !important;
        border: 1px solid var(--border-color) !important;
        border-radius: 8px !important;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.08) !important;
    }
    </style>
""", unsafe_allow_html=True)

# ─── HEADER SECTION ───
st.title("📊 RBI Handbook Enterprise Audit & Verification Suite")
st.markdown("Automated cross-validation engine for matching multi-sheet Excel records against final publication PDFs with native yellow error highlighting.")

st.markdown("---")

# ─── SIDEBAR UPLOADERS ───
with st.sidebar:
    st.header("📂 Data Source Configuration")
    st.markdown("Upload your source artifacts to initiate verification.")
    excel_file = st.file_uploader("Upload Excel Workbook", type=["xlsx"])
    pdf_file = st.file_uploader("Upload Publication PDF", type=["pdf"])
    
    st.markdown("---")
    st.markdown("### ⚙️ Audit Parameters")
    strict_case = st.checkbox("Strict Case Sensitivity", value=True)
    precision_mode = st.checkbox("0.5 Financial Rounding Logic", value=True)
    st.info("System automatically aligns cell coordinates and injects native PDF highlights.")

def annotate_original_pdf(pdf_bytes, highlights_by_page):
    reader = pypdf.PdfReader(io.BytesIO(pdf_bytes))
    writer = pypdf.PdfWriter()
    writer.append(reader)

    for page_num_1idx, boxes in highlights_by_page.items():
        page_idx = page_num_1idx - 1
        if page_idx < len(writer.pages):
            page_height = float(writer.pages[page_idx].mediabox.height)
            
            for box in boxes:
                x0, top, x1, bottom = box
                y0 = page_height - bottom
                y1 = page_height - top
                
                rect = (max(0, x0 - 2), max(0, y0 - 2), x1 + 2, y1 + 2)
                
                qp = ArrayObject([
                    FloatObject(x0 - 2), FloatObject(y1 + 1),
                    FloatObject(x1 + 2), FloatObject(y1 + 1),
                    FloatObject(x0 - 2), FloatObject(y0 - 1),
                    FloatObject(x1 + 2), FloatObject(y0 - 1)
                ])
                
                try:
                    highlight = Highlight(
                        rect=rect,
                        quad_points=qp,
                        highlight_color="ffff00",  # Native Yellow Highlight
                        printing=True
                    )
                    writer.add_annotation(page_number=page_idx, annotation=highlight)
                except Exception:
                    pass

    output_buffer = io.BytesIO()
    writer.write(output_buffer)
    output_buffer.seek(0)
    return output_buffer.getvalue()

if excel_file and pdf_file:
    with st.spinner("🔄 Initializing enterprise audit pipeline and coordinate mapping..."):
        try:
            hashtag_findings = []
            missing_excel_items = []
            unrecognized_pdf_words = []
            table_name_mismatches = []
            case_mismatches = []
            
            pdf_bytes_data = pdf_file.getvalue()
            
            with pdfplumber.open(io.BytesIO(pdf_bytes_data)) as pdf_check:
                total_pages = len(pdf_check.pages)
                highlights_by_page = {i: [] for i in range(1, total_pages + 1)}

            def dehyphenate_text(text):
                if not text:
                    return ""
                text = re.sub(r'(\b[a-zA-Z]+)-\s*[\r\n]+\s*([a-zA-Z]+\b)', r'\1\2', text)
                text = re.sub(r'(\b[a-zA-Z]{2,})[\r\n]+\s*([a-zA-Z]{2,}\b)', r'\1\2', text)
                return text.replace('\r\n', ' ').replace('\n', ' ').replace('\r', ' ').replace('\t', ' ')

            def clean_visual_text(val):
                if val is None:
                    return ""
                if isinstance(val, (int, float)):
                    if math.isnan(val) or math.isinf(val):
                        return ""
                    return f"{val}"
                if isinstance(val, (datetime.datetime, datetime.date)):
                    return val.strftime("%B")
                val_str = str(val).strip()
                financial_neg = re.match(r'^\(([\d\.,\s]+)\)$', val_str)
                if financial_neg:
                    val_str = f"-{financial_neg.group(1)}"
                if not val_str or val_str in ["nan", "nat"]:
                    return ""
                val_str = val_str.replace('/', ' ').replace('\\', ' ')
                val_str = re.sub(r"['’‘]s\b", "", val_str)
                val_str = val_str.replace("'", " ").replace("’", " ").replace("‘", " ")
                val_str = val_str.replace('\xa0', ' ').replace('\xad', ' ')
                val_str = val_str.replace('–', '-').replace('—', '-')
                val_str = val_str.replace('(', ' ').replace(')', ' ')
                val_str = val_str.replace(",", "").replace("$", "").replace("*", "").replace(":", " ")
                val_str = re.sub(r'[^a-zA-Z0-9\s\.\-]', '', val_str)
                return re.sub(r'\s+', ' ', val_str).strip()

            def round_half_up(val):
                if val >= 0:
                    return math.floor(val + 0.5)
                else:
                    return math.ceil(val - 0.5)

            def generate_precision_variants(val):
                if not isinstance(val, (int, float)) or math.isnan(val) or math.isinf(val):
                    return []
                abs_val = abs(val)
                int_val = int(val)
                exact_str = f"{val}"
                exact_2dec = f"{val:.2f}"
                one_dec = f"{math.floor(val * 10 + 0.5) / 10:.1f}" if val >= 0 else f"{math.ceil(val * 10 - 0.5) / 10:.1f}"
                int_rounded = f"{round_half_up(val)}"
                abs_int_rounded = f"{round_half_up(abs_val)}"
                
                variants = {
                    exact_str, exact_2dec, f"{val:.1f}", one_dec, int_rounded,
                    f"{int_val}", f"{int_val:,}", f"{val/1000:.1f}" if val >= 1000 else "",
                    f"{abs_val:.2f}", abs_int_rounded, f"{int(abs_val)}"
                }
                return [v for v in variants if v]

            def is_token_match(w1, w2, case_sensitive=False):
                if not isinstance(w1, str) or not w1:
                    return False
                if not case_sensitive:
                    w1_check = w1.lower().replace('-', '').replace('.', '')
                    w2_check = str(w2).lower().replace('-', '').replace('.', '')
                else:
                    w1_check = w1.replace('-', '').replace('.', '')
                    w2_check = str(w2).replace('-', '').replace('.', '')
                if w1_check == w2_check:
                    return True
                if w1_check.endswith('s') and w1_check[:-1] == w2_check:
                    return True
                if w2_check.endswith('s') and w2_check[:-1] == w1_check:
                    return True
                return False

            def highlight_word_on_pages(pdf_obj, search_term, target_pages=None):
                term_str = str(search_term).strip()
                if not term_str or len(term_str) < 2:
                    return
                for p_idx, page in enumerate(pdf_obj.pages, start=1):
                    if target_pages and p_idx not in target_pages:
                        continue
                    matches = page.search(term_str)
                    if not matches:
                        words = page.extract_words()
                        for w in words:
                            if term_str.lower() in w['text'].lower():
                                highlights_by_page[p_idx].append((w['x0'], w['top'], w['x1'], w['bottom']))
                    else:
                        for m in matches:
                            highlights_by_page[p_idx].append((m['x0'], m['top'], m['x1'], m['bottom']))

            # ─── GATHER EXCEL DATA ───
            wb = openpyxl.load_workbook(excel_file, data_only=True)
            excel_items = []
            excel_words_set = set()
            excel_table_titles = []
            excel_text_values = []
            
            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                rows = list(ws.iter_rows(values_only=False))
                if not rows:
                    continue

                sheet_title_candidates = []
                for r in rows[:3]:
                    row_str = " ".join([str(c.value) for c in r if c.value is not None])
                    clean_title = clean_visual_text(row_str)
                    if clean_title and len(clean_title) > 3 and not clean_title.isdigit():
                        sheet_title_candidates.append(clean_title)

                if sheet_title_candidates:
                    full_sheet_title = " ".join(sheet_title_candidates[:2])
                    excel_table_titles.append({
                        "Sheet_Name": sheet_name,
                        "Title": full_sheet_title,
                        "Raw_Title": sheet_title_candidates[0]
                    })

                for row_idx, row in enumerate(rows, start=1):
                    for cell in row:
                        if cell.value is None:
                            continue
                        raw_str = str(cell.value).strip()
                        if "###" in raw_str:
                            continue

                        if isinstance(cell.value, (float, int)):
                            excel_items.append({
                                "Type": "Number",
                                "Value": cell.value,
                                "Location": f"Sheet: {sheet_name} | Row: {row_idx}"
                            })
                        else:
                            norm_val = clean_visual_text(cell.value)
                            if norm_val and len(norm_val) > 1:
                                excel_items.append({
                                    "Type": "Text",
                                    "Clean_Val": norm_val,
                                    "Display_Val": raw_str,
                                    "Location": f"Sheet: {sheet_name} | Row: {row_idx}"
                                })
                                excel_text_values.append(norm_val)
                                for w in norm_val.split():
                                    if len(w) > 1 and not w.replace('-', '').replace('.', '').isdigit():
                                        excel_words_set.add(w)

            excel_words_set.update([
                "total", "operatives", "the", "ensure", "onwards", "relates", "outstanding", 
                "coscardbs", "scardbs", "dbs", "and", "allied", "years", "pertain", "coscard", "scards",
                "liabilities", "assets", "bond", "bonds", "excludes", "refinancing", "includes",
                "paid", "up", "capital", "paidup", "score", "state", "sdg", "sdgs", "mospi", "maternal", "mortality", "ratio"
            ])

            skip_tokens = {
                "page", "table", "continued", "source", "notes", "see", "also", "services",
                "nonnet", "nonrbi", "nonrbis", "nonm3", "ment", "tion", "ing", "ance", "nonment",
                "theensure", "scardbsand", "andallied", "scardbs", "coscardbs", "excludesrefinancing",
                "sd", "gs"
            }

            # ─── EXTRACT PDF TEXT & MAP HIGHLIGHTS ───
            pdf_global_raw = ""
            pdf_words_exact_info = []
            
            with pdfplumber.open(io.BytesIO(pdf_bytes_data)) as pdf:
                for page_num, page in enumerate(pdf.pages, start=1):
                    page_obj = page.dedupe_chars() if hasattr(page, 'dedupe_chars') else page
                    raw_text = page_obj.extract_text()
                    
                    if raw_text:
                        if "#" in raw_text:
                            hashtag_findings.append({
                                "Notification": "# character detected in PDF",
                                "Location Details": f"Page {page_num}"
                            })
                        
                        dehyphenated_page_text = dehyphenate_text(raw_text)
                        clean_pdf_page = clean_visual_text(dehyphenated_page_text)
                        
                        raw_pdf_words = clean_pdf_page.split()
                        expanded_pdf_words = []
                        for w in raw_pdf_words:
                            if re.search(r'\([A-Z]+s?\)', w) or w.isupper() or len(w) <= 3:
                                expanded_pdf_words.append(w)
                            else:
                                s1 = re.sub(r'([a-z])([A-Z])', r'\1 \2', w)
                                s2 = re.sub(r'([A-Z]+)([A-Z][a-z])', r'\1 \2', s1)
                                expanded_pdf_words.extend(s2.split())

                        pdf_words = [w for w in expanded_pdf_words if len(w) > 1 and not w.replace('.', '').replace('-', '').isdigit()]
                        for w in set(pdf_words):
                            pdf_words_exact_info.append((w, page_num))
                            if not any(is_token_match(w, ew, case_sensitive=False) for ew in excel_words_set) and w.lower() not in skip_tokens:
                                unrecognized_pdf_words.append({
                                    "Page Number": page_num,
                                    "Unrecognized Word in PDF": w,
                                    "Details": "Word present in PDF output but missing from Excel source."
                                })

                        pdf_global_raw += " " + dehyphenated_page_text

            pdf_global_text = clean_visual_text(pdf_global_raw)
            pdf_spaceless_pool = re.sub(r'\s+', '', pdf_global_text)
            pdf_words_exact_set = set([item[0] for item in pdf_words_exact_info])

            # ─── VALIDATION & ERROR MAPPING ───
            with pdfplumber.open(io.BytesIO(pdf_bytes_data)) as pdf_pages_obj:
                for t_info in excel_table_titles:
                    sheet_name = t_info["Sheet_Name"]
                    title_text = t_info["Title"]
                    
                    norm_excel_title = re.sub(r'[^a-zA-Z0-9]', '', title_text).lower()
                    norm_pdf_pool = re.sub(r'[^a-zA-Z0-9]', '', pdf_global_text).lower()
                    
                    table_num_match = re.search(r'table\s*\d+', title_text.lower())
                    table_num_str = table_num_match.group(0).replace(" ", "") if table_num_match else ""

                    if (table_num_str and table_num_str in norm_pdf_pool) or (norm_excel_title in norm_pdf_pool) or (title_text.lower() in pdf_global_text.lower()):
                        continue

                    title_words = [w for w in title_text.split() if len(w) > 2 and not w.isdigit()]
                    matched_title_words = [w for w in title_words if any(is_token_match(w, pw, case_sensitive=False) for pw in pdf_words_exact_set)]
                    
                    match_ratio = (len(matched_title_words) / len(title_words)) if title_words else 1.0
                    if match_ratio < 0.30:
                        table_name_mismatches.append({
                            "Excel Sheet Name": sheet_name,
                            "Excel Table Title": t_info["Raw_Title"],
                            "Status": "Table Name / Header Not Found in PDF",
                            "Details": "The table title in Excel does not match any table title printed in the PDF document."
                        })

                for item in excel_items:
                    if item["Type"] == "Text":
                        clean_val = item["Clean_Val"]
                        spaceless_clean_val = clean_val.replace(" ", "")
                        
                        if clean_val in pdf_global_text or spaceless_clean_val in pdf_spaceless_pool:
                            continue

                        if clean_val.lower() in pdf_global_text.lower() or spaceless_clean_val.lower() in pdf_spaceless_pool.lower():
                            words = clean_val.split()
                            for w in words:
                                if w and not w.isdigit() and len(w) > 2:
                                    if w not in pdf_words_exact_set and w.lower() in [pw.lower() for pw in pdf_words_exact_set]:
                                        match_pages = [p_num for pw, p_num in pdf_words_exact_info if pw.lower() == w.lower()]
                                        page_str = f"Page {match_pages[0]}" if match_pages else "N/A"
                                        case_mismatches.append({
                                            "Excel Location": item["Location"],
                                            "Excel Text (Expected Case)": w,
                                            "PDF Text (Printed Case)": w.lower() if w.isupper() else w.upper(),
                                            "PDF Page": page_str,
                                            "Issue": "Case Sensitivity Discrepancy"
                                        })
                                        highlight_word_on_pages(pdf_pages_obj, w, target_pages=match_pages)
                            continue

                        words = clean_val.split()
                        missing_words = [w for w in words if not any(is_token_match(w, pw, case_sensitive=False) for pw in pdf_words_exact_set)]

                        if missing_words:
                            missing_excel_items.append({
                                "Excel Location": item["Location"],
                                "Excel Content / Sentence": item["Display_Val"],
                                "Issue": "Missing Word(s) in PDF / Error in Sentence",
                                "Missing Detail": ", ".join(missing_words)
                            })

                    elif item["Type"] == "Number":
                        val = float(item["Value"])
                        variants = generate_precision_variants(val)

                        if any(v in pdf_global_text or v in pdf_spaceless_pool or v.replace(',', '') in pdf_spaceless_pool for v in variants):
                            continue

                        expected_rounded = round_half_up(val)
                        missing_excel_items.append({
                            "Excel Location": item["Location"],
                            "Excel Content / Sentence": f"Numeric value: {val}",
                            "Issue": "Number Not Found in PDF / Value Error",
                            "Missing Detail": f"Exact: {val:.2f} | Expected Rounded (>=0.5): {expected_rounded}"
                        })
                        highlight_word_on_pages(pdf_pages_obj, f"{int(val):,}" if val.is_integer() else f"{val}")

            # Deduplicate bounding boxes
            for p_num in highlights_by_page:
                unique_boxes = []
                seen = set()
                for b in highlights_by_page[p_num]:
                    key = (round(b[0], 1), round(b[1], 1), round(b[2], 1), round(b[3], 1))
                    if key not in seen:
                        seen.add(key)
                        unique_boxes.append(b)
                highlights_by_page[p_num] = unique_boxes

            annotated_pdf_bytes = annotate_original_pdf(pdf_bytes_data, highlights_by_page)
            base64_pdf = base64.b64encode(annotated_pdf_bytes).decode('utf-8')

            # ─── ENTERPRISE METRICS DASHBOARD ───
            total_discrepancies = len(case_mismatches) + len(table_name_mismatches) + len(missing_excel_items) + len(unrecognized_pdf_words)
            
            c1, c2, c3, c4 = st.columns(4)
            with c1:
                st.metric("Total Excel Cells Audited", f"{len(excel_items):,}")
            with c2:
                st.metric("Publication Pages", f"{total_pages}")
            with c3:
                st.metric("Total Discrepancies Flagged", f"{total_discrepancies}", delta="-0" if total_discrepancies == 0 else f"{total_discrepancies} errors", delta_color="inverse")
            with c4:
                accuracy_score = max(0.0, 100.0 - (total_discrepancies * 0.5))
                st.metric("Audit Integrity Score", f"{accuracy_score:.1f}%")

            st.markdown("---")

            def render_discrepancy_table(df):
                if df is None or df.empty:
                    return
                cols = df.columns.tolist()
                header_th = "".join([
                    f'<th style="background-color: #E2E8F0 !important; color: #172B4D !important; font-weight: 700 !important; border: 1px solid #CBD5E1 !important; padding: 10px 14px !important; text-align: left; font-size: 0.875rem !important; white-space: nowrap !important; letter-spacing: 0.01em !important;">{html.escape(str(c))}</th>'
                    for c in cols
                ])
                rows_tr = []
                for row_idx, row in enumerate(df.itertuples(index=False)):
                    row_bg = "#FFFFFF" if row_idx % 2 == 0 else "#F8FAFC"
                    tds = []
                    for val in row:
                        v_str = "" if val is None or (isinstance(val, float) and pd.isna(val)) else str(val)
                        safe_val = html.escape(v_str).replace('\n', '<br>')
                        tds.append(
                            f'<td style="background-color: inherit; color: #1E293B !important; border: 1px solid #CBD5E1 !important; padding: 9px 14px !important; font-size: 0.875rem !important; vertical-align: top !important; line-height: 1.45 !important;">{safe_val}</td>'
                        )
                    rows_tr.append(f'<tr class="audit-row" style="background-color: {row_bg};">{"".join(tds)}</tr>')
                rows_html = "".join(rows_tr)
                table_html = f'<div class="audit-table-wrapper"><table class="audit-table"><thead><tr>{header_th}</tr></thead><tbody>{rows_html}</tbody></table></div>'
                st.markdown(table_html, unsafe_allow_html=True)

            # ─── DISCREPANCY LOGS DISPLAYED FIRST ───
            st.markdown("### 📋 Detailed Audit Discrepancy Logs")
            st.markdown("Review all identified discrepancies categorized by issue type before publishing.")
            
            if case_mismatches:
                st.error("🚨 Case Sensitivity Mismatches")
                render_discrepancy_table(pd.DataFrame(case_mismatches).drop_duplicates())

            if table_name_mismatches:
                st.error("🚨 Mismatched or Incorrect Table / Sheet Names")
                render_discrepancy_table(pd.DataFrame(table_name_mismatches).drop_duplicates())

            if missing_excel_items:
                st.error("🚨 Missing Words / Numbers / Sentence Errors in PDF")
                render_discrepancy_table(pd.DataFrame(missing_excel_items).drop_duplicates())

            if unrecognized_pdf_words:
                st.warning("⚠️ Unrecognized / Extra Words in PDF")
                render_discrepancy_table(pd.DataFrame(unrecognized_pdf_words).drop_duplicates(subset=["Page Number", "Unrecognized Word in PDF"]))

            if total_discrepancies == 0:
                st.balloons()
                st.success("✅ PERFECT AUDIT! All items verified successfully with zero discrepancies across all records.")

            st.markdown("---")

            # ─── ROBUST INTERACTIVE SPREADSHEET ENGINE ───
            st.markdown("### 🖥️ Interactive Reference Workspace")
            st.markdown("Launch a dedicated full-screen side-by-side comparison view with the Natively Highlighted PDF on the left and a **Fully Interactive Excel Editor** on the right.")

            sheets_data = {}
            sheet_tabs_html = ""
            
            for idx, s_name in enumerate(wb.sheetnames):
                ws = wb[s_name]
                max_r = max(ws.max_row, 50)
                max_c = max(ws.max_column, 15)
                
                sheet_grid = []
                for r_idx in range(1, max_r + 1):
                    row_data = []
                    for c_idx in range(1, max_c + 1):
                        val = ws.cell(r_idx, c_idx).value
                        row_data.append("" if val is None else str(val))
                    sheet_grid.append(row_data)
                
                sheets_data[s_name] = sheet_grid

            sheets_json_str = json.dumps(sheets_data)
            first_sheet_name = wb.sheetnames[0]

            split_screen_html = f"""
            <!DOCTYPE html>
            <html>
            <head>
                <meta charset="utf-8" />
                <title>RBI Handbook 50/50 Fully Interactive Comparison Workspace</title>
                <!-- Handsontable CSS & JS -->
                <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/handsontable@14.6.1/dist/handsontable.full.min.css" />
                <script src="https://cdn.jsdelivr.net/npm/handsontable@14.6.1/dist/handsontable.full.min.js"></script>
                <style>
                    * {{ box-sizing: border-box; margin: 0; padding: 0; }}
                    body {{ display: flex; height: 100vh; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; background: #0F172A; overflow: hidden; }}
                    .pane {{ width: 50%; height: 100vh; display: flex; flex-direction: column; background: #ffffff; }}
                    .pdf-pane {{ border-right: 2px solid #D9E2EF; }}
                    
                    /* PDF Header */
                    .pane-header-pdf {{ background: #0F294A; color: #ffffff; padding: 10px 18px; font-size: 13.5px; font-weight: 600; display: flex; align-items: center; border-bottom: 1px solid #1E3A8A; }}
                    .pane-content-pdf {{ flex: 1; overflow: hidden; background: #0b1120; }}
                    iframe {{ width: 100%; height: 100%; border: none; }}

                    /* Excel Desktop Theme */
                    .excel-container {{ display: flex; flex-direction: column; height: 100vh; background: #F5F7FB; }}
                    .excel-ribbon-top {{ background: #107C41; color: white; padding: 9px 18px; font-size: 13.5px; font-weight: 600; display: flex; justify-content: space-between; align-items: center; }}
                    .excel-ribbon-tabs {{ background: #E8EEF8; border-bottom: 1px solid #D9E2EF; display: flex; padding: 0 12px; font-size: 12px; color: #1E293B; }}
                    .excel-ribbon-tab {{ padding: 7px 14px; cursor: pointer; border-bottom: 2px solid transparent; font-weight: 600; user-select: none; color: #1E293B; }}
                    .excel-ribbon-tab.active {{ border-bottom: 2px solid #107C41; color: #107C41; background: #FFFFFF; }}
                    
                    .excel-formula-bar {{ background: #FFFFFF; border-bottom: 1px solid #D9E2EF; padding: 5px 12px; display: flex; align-items: center; font-size: 12.5px; color: #172B4D; }}
                    .excel-fx {{ font-weight: bold; color: #475569; margin-right: 10px; font-style: italic; user-select: none; }}
                    .excel-formula-input {{ flex: 1; border: 1px solid #D9E2EF; border-radius: 4px; padding: 4px 8px; font-size: 12px; color: #172B4D; background: #FFFFFF; outline: none; }}
                    .excel-formula-input:focus {{ border-color: #107C41; background: #FFFFFF; }}

                    .excel-grid-container {{ flex: 1 1 auto; min-height: 0; min-width: 0; height: 100%; width: 100%; overflow: auto; background: #FFFFFF; }}
                    #excel-grid .handsontable {{ width: 100%; }}
                    #excel-grid .wtHolder {{ overflow: auto !important; }}
                    .grid-error {{ padding: 18px; color: #991B1B; background: #FEE2E2; font: 14px Arial, sans-serif; }}
                    .excel-footer-tabs {{ background: #F5F7FB; border-top: 1px solid #D9E2EF; min-height: 38px; height: 38px; display: flex; align-items: center; padding: 0 10px; gap: 4px; overflow-x: auto; overflow-y: hidden; white-space: nowrap; flex: 0 0 38px; }}
                    .excel-sheet-tab {{ background: #E8EEF8; padding: 8px 16px; font-size: 12px; border: 1px solid #D9E2EF; border-bottom: none; border-radius: 4px 4px 0 0; color: #1E293B; cursor: pointer; user-select: none; flex: 0 0 auto; white-space: nowrap; font-weight: 500; transition: all 0.15s ease; }}
                    .excel-sheet-tab:hover {{ background: #D9E2EF; color: #172B4D; }}
                    .excel-sheet-tab.active-tab {{ background: #FFFFFF; font-weight: bold; color: #107C41; border-top: 3px solid #107C41; }}
                </style>
            </head>
            <body>
                <!-- Left Pane: PDF Viewer -->
                <div class="pane pdf-pane">
                    <div class="pane-header-pdf">
                        <span>📄 Natively Highlighted Publication PDF (Left 50%)</span>
                    </div>
                    <div class="pane-content-pdf">
                        <iframe id="pdf-frame"></iframe>
                    </div>
                </div>

                <!-- Right Pane: Interactive Excel Suite -->
                <div class="pane">
                    <div class="excel-container">
                        <div class="excel-ribbon-top">
                            <span>🟢 Excel Editor - {excel_file.name}</span>
                            <span style="font-size: 12px; opacity: 0.9;">Fully Interactive</span>
                        </div>
                        <div class="excel-ribbon-tabs">
                            <div class="excel-ribbon-tab active">Home</div>
                            <div class="excel-ribbon-tab">Insert</div>
                            <div class="excel-ribbon-tab">Page Layout</div>
                            <div class="excel-ribbon-tab">Formulas</div>
                            <div class="excel-ribbon-tab">Data</div>
                            <div class="excel-ribbon-tab">Review</div>
                            <div class="excel-ribbon-tab">View</div>
                        </div>
                        <div class="excel-formula-bar">
                            <span class="excel-fx">fx</span>
                            <input type="text" id="formula-input" class="excel-formula-input" value="Active Sheet: {first_sheet_name}" readonly>
                        </div>
                        <div id="excel-grid" class="excel-grid-container"><div style="padding:16px;color:#475569;font:13px Arial,sans-serif;">Loading worksheet…</div></div>
                        <div id="excel-footer-tabs" class="excel-footer-tabs" aria-label="Excel worksheets"></div>
                    </div>
                </div>

                <script>
                    const sheetsData = {sheets_json_str};
                    let hotInstance = null;
                    let currentSheet = "{first_sheet_name}";

                    function showGridError(message) {{
                        const container = document.getElementById('excel-grid');
                        container.innerHTML = '<div class="grid-error"><strong>Excel grid could not load.</strong><br>' + message + '<br><br>Check your internet connection and allow cdn.jsdelivr.net, then refresh this page.</div>';
                    }}

                    function initHandsontable(data) {{
                        const container = document.getElementById('excel-grid');
                        if (typeof Handsontable === 'undefined') {{
                            showGridError('Handsontable library did not load from the CDN.');
                            return;
                        }}
                        if (!Array.isArray(data) || data.length === 0) {{
                            showGridError('This worksheet has no cell data to display.');
                            return;
                        }}
                        if (hotInstance) {{
                            hotInstance.destroy();
                            hotInstance = null;
                        }}
                        container.innerHTML = '';
                        try {{
                            hotInstance = new Handsontable(container, {{
                                data: data,
                                width: '100%',
                                height: '100%',
                                stretchH: 'none',
                                rowHeaders: true,
                                colHeaders: true,
                                contextMenu: true,
                                manualColumnResize: true,
                                manualRowResize: true,
                                selectionMode: 'multiple',
                                outsideClickDeselects: false,
                                licenseKey: 'non-commercial-and-evaluation',
                                afterSelection: function(row, col) {{
                                    if (row < 0 || col < 0 || !hotInstance) return;
                                    const cellVal = hotInstance.getDataAtCell(row, col);
                                    document.getElementById('formula-input').value = cellVal == null ? '' : String(cellVal);
                                }}
                            }});
                            hotInstance.render();
                            hotInstance.refreshDimensions();
                        }} catch (err) {{
                            showGridError('Initialization error: ' + (err && err.message ? err.message : String(err)));
                        }}
                    }}

                    function renderSheetTabs() {{
                        const footer = document.getElementById('excel-footer-tabs');
                        footer.innerHTML = '';
                        Object.keys(sheetsData).forEach(function(sheetName, index) {{
                            const tab = document.createElement('button');
                            tab.type = 'button';
                            tab.className = 'excel-sheet-tab' + (sheetName === currentSheet ? ' active-tab' : '');
                            tab.textContent = sheetName;
                            tab.dataset.sheet = sheetName;
                            tab.title = 'Open worksheet ' + sheetName;
                            tab.addEventListener('click', function() {{ switchSheet(sheetName); }});
                            footer.appendChild(tab);
                        }});
                    }}

                    function switchSheet(sheetName) {{
                        if (!Object.prototype.hasOwnProperty.call(sheetsData, sheetName)) {{
                            showGridError('Worksheet not found: ' + sheetName);
                            return;
                        }}
                        currentSheet = sheetName;
                        document.querySelectorAll('.excel-sheet-tab').forEach(tab => {{
                            tab.classList.toggle('active-tab', tab.dataset.sheet === sheetName);
                        }});
                        document.getElementById('formula-input').value = 'Active Sheet: ' + sheetName;
                        initHandsontable(sheetsData[sheetName]);
                    }}

                    window.addEventListener('load', function() {{
                        renderSheetTabs();
                        if (typeof Handsontable === 'undefined') {{
                            showGridError('The Handsontable script was blocked or could not be downloaded.');
                            return;
                        }}
                        switchSheet(currentSheet);
                    }});

                    const pdfData = "{base64_pdf}";
                    const byteCharacters = atob(pdfData);
                    const byteNumbers = new Array(byteCharacters.length);
                    for (let i = 0; i < byteCharacters.length; i++) {{
                        byteNumbers[i] = byteCharacters.charCodeAt(i);
                    }}
                    const byteArray = new Uint8Array(byteNumbers);
                    const blob = new Blob([byteArray], {{ type: 'application/pdf' }});
                    const blobUrl = URL.createObjectURL(blob);
                    document.getElementById('pdf-frame').src = blobUrl;
                </script>
            </body>
            </html>
            """
            
            # ─── LAUNCH WORKSPACE ───
            # Use components.html() because st.markdown() does not reliably execute JS.
            html_bytes = split_screen_html.encode("utf-8")
            html_base64 = base64.b64encode(html_bytes).decode("ascii")

            btn_html = f"""
            <!doctype html>
            <html>
            <head>
              <meta charset="utf-8" />
              <style>
                body {{
                  margin: 0;
                  padding: 8px 0;
                  font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
                  background: transparent;
                }}
                #open-workspace {{
                  padding: 12px 24px;
                  background-color: #2563EB;
                  color: #FFFFFF;
                  border: 1px solid #2563EB;
                  border-radius: 6px;
                  font-weight: 600;
                  font-size: 15px;
                  cursor: pointer;
                  box-shadow: 0 1px 3px rgba(37, 99, 235, 0.25);
                  transition: all 0.15s ease;
                  display: inline-flex;
                  align-items: center;
                  gap: 8px;
                }}
                #open-workspace:hover {{
                  background-color: #1D4ED8;
                  border-color: #1D4ED8;
                  color: #FFFFFF;
                  box-shadow: 0 2px 6px rgba(29, 78, 216, 0.35);
                }}
                #open-workspace:focus {{
                  outline: 2px solid #2563EB;
                  outline-offset: 2px;
                }}
                #launch-status {{
                  margin-left: 12px;
                  color: #172B4D;
                  font-size: 13px;
                  font-weight: 500;
                }}
              </style>
            </head>
            <body>
              <button id="open-workspace">
                🖥️ Open 50/50 Side-by-Side Workspace in New Tab
              </button>
              <span id="launch-status"></span>
              <script>
                const workspaceHTML = atob("{html_base64}");
                document.getElementById("open-workspace").addEventListener("click", function () {{
                  const newTab = window.open("about:blank", "_blank");
                  if (!newTab) {{
                    document.getElementById("launch-status").textContent =
                      "Popup blocked. Use the Download Workspace HTML button below.";
                    return;
                  }}
                  newTab.document.open();
                  newTab.document.write(workspaceHTML);
                  newTab.document.close();
                  document.getElementById("launch-status").textContent = "Workspace opened.";
                }});
              </script>
            </body>
            </html>
            """

            components.html(btn_html, height=65, scrolling=False)

            # Reliable fallback if the browser blocks popups from embedded components.
            st.download_button(
                label="⬇️ Download Interactive 50/50 Workspace (HTML)",
                data=split_screen_html.encode("utf-8"),
                file_name="RBI_Interactive_Comparison_Workspace.html",
                mime="text/html",
                help="If the new-tab button is blocked, download this file and open it in Chrome or Edge.",
            )

            st.caption(
                "The right pane displays all populated Excel rows and columns in a clickable, "
                "editable grid. Click a cell to view its value in the formula bar. This is an "
                "interactive browser grid, not a full Excel clone: macros, formula recalculation, "
                "and all original cell formatting are not reproduced."
            )

        except Exception as e:
            st.error(f"Enterprise audit pipeline execution exception: {e}")
else:
    st.info("👈 Please upload your multi-sheet Excel file and publication PDF in the sidebar to launch the enterprise workspace.")
