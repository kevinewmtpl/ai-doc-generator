import os
import json
import base64
import hmac
import re
from copy import deepcopy
from urllib.parse import quote
from io import BytesIO
from datetime import date, datetime, timedelta

import requests

import streamlit as st
from openai import OpenAI
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn

# PowerPoint support is optional so the existing system continues to work
# even before python-pptx is added to requirements.txt.
try:
    from pptx import Presentation
    from pptx.util import Pt as PPTXPt
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    from pptx.dml.color import RGBColor
    PPTX_AVAILABLE = True
except Exception:
    Presentation = None
    PPTXPt = None
    MSO_SHAPE_TYPE = None
    RGBColor = None
    PPTX_AVAILABLE = False

# =====================
# PAGE CONFIG
# =====================
st.set_page_config(
    page_title="EWMT AI Document System",
    page_icon="🏗️",
    layout="wide"
)

# =====================
# SESSION PAGE CONTROL
# =====================
PAGES = [
    "🏠 Dashboard",
    "📄 Method Statement",
    "📑 Method Statement PRO",
    "🏗️ Lifting Plan",
    "⚠️ Risk Assessment Pro",
    "🧰 Lifting Gear Register",
    "👷 Worker Training Certificate",
    "⏰ Expiry Alerts",
    "⚙️ Settings"
]

if "page" not in st.session_state:
    st.session_state.page = "🏠 Dashboard"

# =====================
# OPENAI CLIENT
# =====================
client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

# =====================
# VECTOR STORES
# =====================
MS_VECTOR_STORE_ID = "vs_69ecc533a1208191a8595c674753e99e"
RA_VECTOR_STORE_ID = "vs_69ecd191648481919d1d1d57f21264af"
LP_VECTOR_STORE_ID = "vs_69ecdc44d59081919fb10574510b7454"

# =====================
# PATHS + ASSETS
# =====================
BASE_DIR = os.path.dirname(__file__)
ASSET_DIR = os.path.join(BASE_DIR, "assets")

MS_TEMPLATE = os.path.join(BASE_DIR, "Templates", "Method of statement Template.docx")
RA_TEMPLATE = os.path.join(BASE_DIR, "Templates", "RA Template.docx")
LP_TEMPLATE = os.path.join(BASE_DIR, "Templates", "Lifting Plan Template.docx")

# MOS PRO uses your PowerPoint as the master document.
# The code accepts either filename so you can upload your current file without renaming it.
MOS_PRO_TEMPLATE_CANDIDATES = [
    os.path.join(BASE_DIR, "Templates", "MOS New.pptx"),
    os.path.join(BASE_DIR, "Templates", "MOS New(1).pptx"),
    os.path.join(BASE_DIR, "Templates", "Method Statement PRO.pptx"),
]


def image_to_base64(path):
    try:
        with open(path, "rb") as img_file:
            return base64.b64encode(img_file.read()).decode()
    except Exception:
        return ""


def asset_image(filename, fallback="banner.jpg"):
    path = os.path.join(ASSET_DIR, filename)

    if os.path.exists(path):
        return image_to_base64(path)

    fallback_path = os.path.join(ASSET_DIR, fallback)

    if os.path.exists(fallback_path):
        return image_to_base64(fallback_path)

    return ""


HEADER_IMAGE = asset_image("banner.jpg")
METHOD_IMAGE = asset_image("method_statement.jpg")
LIFTING_IMAGE = asset_image("lifting_plan.jpg")
RISK_IMAGE = asset_image("risk_assessment.jpg", "method_statement.jpg")
GEAR_IMAGE = asset_image("gear_register.jpg", "lifting_plan.jpg")
TRAINING_IMAGE = asset_image("training_certificate.jpg", "method_statement.jpg")
EXPIRY_IMAGE = asset_image("expiry_alert.jpg", "training_certificate.jpg")

# =====================
# PROFESSIONAL UI STYLE
# =====================
st.markdown(f"""
<style>
.stApp {{
    background: linear-gradient(180deg, #f4f7fb 0%, #eef2f7 100%);
}}

.block-container {{
    padding-top: 1.1rem;
    padding-bottom: 2rem;
    max-width: 1520px;
}}

[data-testid="stSidebar"] {{
    background: linear-gradient(180deg, #071126 0%, #0f172a 52%, #1e293b 100%);
    border-right: 1px solid rgba(255,255,255,0.08);
}}

[data-testid="stSidebar"] * {{
    color: white;
}}

.ewmt-header {{
    position: relative;
    overflow: hidden;
    min-height: 155px;
    background:
        linear-gradient(90deg, rgba(15,23,42,0.96), rgba(30,58,138,0.90), rgba(15,23,42,0.76)),
        url("data:image/jpg;base64,{HEADER_IMAGE}");
    background-size: cover;
    background-position: center;
    padding: 30px 36px;
    border-radius: 24px;
    color: white;
    margin-bottom: 24px;
    box-shadow: 0px 16px 38px rgba(15,23,42,0.26);
    border: 1px solid rgba(255,255,255,0.14);
}}

.ewmt-header:after {{
    content: "";
    position: absolute;
    left: 36px;
    right: 36px;
    bottom: 0;
    height: 4px;
    background: linear-gradient(90deg, #f59e0b, rgba(245,158,11,0.15));
    border-radius: 99px;
}}

.ewmt-badge {{
    display: inline-block;
    background: rgba(255,255,255,0.14);
    border: 1px solid rgba(255,255,255,0.22);
    padding: 7px 12px;
    border-radius: 999px;
    font-size: 13px;
    color: #dbeafe;
    margin-bottom: 10px;
    backdrop-filter: blur(6px);
}}

.ewmt-title {{
    font-size: 36px;
    line-height: 1.15;
    font-weight: 900;
    margin-bottom: 8px;
    letter-spacing: -0.6px;
}}

.ewmt-subtitle {{
    font-size: 17px;
    color: #dbeafe;
    max-width: 850px;
}}

.section-title {{
    font-size: 25px;
    font-weight: 850;
    color: #0f172a;
    margin-top: 18px;
    margin-bottom: 8px;
}}

.section-caption {{
    color: #64748b;
    margin-bottom: 18px;
}}

.metric-card {{
    background: rgba(255,255,255,0.92);
    border: 1px solid #e2e8f0;
    border-radius: 18px;
    padding: 18px 20px;
    box-shadow: 0px 8px 24px rgba(15,23,42,0.07);
    position: relative;
    overflow: hidden;
    min-height: 105px;
}}

.metric-card:before {{
    content: "";
    position: absolute;
    left: 0;
    top: 0;
    bottom: 0;
    width: 5px;
    background: #f59e0b;
}}

.metric-label {{
    color: #64748b;
    font-size: 14px;
    font-weight: 700;
}}

.metric-value {{
    color: #0f172a;
    font-size: 32px;
    font-weight: 900;
    margin-top: 6px;
}}

.metric-small {{
    color: #94a3b8;
    font-size: 12px;
    margin-top: 2px;
}}

.dashboard-card {{
    background: white;
    border-radius: 22px;
    border: 1px solid #e2e8f0;
    box-shadow: 0px 10px 28px rgba(15,23,42,0.08);
    overflow: hidden;
    margin-bottom: 12px;
    min-height: 330px;
    transition: all 0.18s ease;
}}

.dashboard-card:hover {{
    transform: translateY(-4px);
    box-shadow: 0px 18px 38px rgba(15,23,42,0.14);
    border-color: rgba(30,58,138,0.28);
}}

.dashboard-img {{
    height: 145px;
    background-size: cover;
    background-position: center;
    position: relative;
}}

.dashboard-img:after {{
    content: "";
    position: absolute;
    inset: 0;
    background: linear-gradient(180deg, rgba(15,23,42,0.05), rgba(15,23,42,0.74));
}}

.dashboard-pill {{
    position: absolute;
    left: 18px;
    bottom: 14px;
    z-index: 2;
    color: white;
    background: rgba(15,23,42,0.70);
    border: 1px solid rgba(255,255,255,0.20);
    padding: 6px 10px;
    border-radius: 999px;
    font-size: 12px;
    font-weight: 800;
}}

.dashboard-content {{
    padding: 20px 22px 10px 22px;
}}

.dashboard-content h3 {{
    margin: 0 0 10px 0;
    color: #0f172a;
    font-size: 23px;
    font-weight: 900;
}}

.dashboard-content p {{
    color: #475569;
    font-size: 15px;
    line-height: 1.55;
    min-height: 52px;
    margin-bottom: 0;
}}

.card-accent {{
    height: 4px;
    width: 58px;
    background: #f59e0b;
    border-radius: 999px;
    margin-bottom: 14px;
}}

.stButton > button {{
    background: linear-gradient(90deg, #1e3a8a, #2563eb);
    color: white;
    border-radius: 12px;
    padding: 0.70rem 1.25rem;
    font-weight: 800;
    border: none;
    width: 100%;
    box-shadow: 0px 8px 18px rgba(30,58,138,0.20);
}}

.stButton > button:hover {{
    background: linear-gradient(90deg, #0f172a, #1e3a8a);
    color: white;
}}

.stDownloadButton > button {{
    background: linear-gradient(90deg, #047857, #059669);
    color: white;
    border-radius: 12px;
    padding: 0.70rem 1.25rem;
    font-weight: 800;
    border: none;
}}

.stDownloadButton > button:hover {{
    background: #065f46;
    color: white;
}}

div[data-testid="stExpander"] {{
    border-radius: 15px;
    border: 1px solid #e2e8f0;
    box-shadow: 0px 4px 14px rgba(15,23,42,0.04);
    background: white;
}}

.footer-note {{
    margin-top: 20px;
    padding: 18px 22px;
    background: #0f172a;
    color: #cbd5e1;
    border-radius: 18px;
    border-left: 5px solid #f59e0b;
}}
</style>
""", unsafe_allow_html=True)

# =====================
# HEADER
# =====================
st.markdown("""
<div class="ewmt-header">
    <div class="ewmt-badge">EWMT INTERNAL AI SYSTEM</div>
    <div class="ewmt-title">Eric Wong Machinery Transportation Pte Ltd</div>
    <div class="ewmt-subtitle">
        Heavy Machinery Moving • Lifting • Transportation • AI Document Control System
    </div>
</div>
""", unsafe_allow_html=True)

# =====================
# COMMON FUNCTIONS
# =====================
def go_to_page(page_name):
    st.session_state.page = page_name
    st.rerun()


def replace_all(doc, data):
    def replace_in_paragraph(paragraph, replacements):
        if not paragraph.runs:
            return

        full_text = "".join(run.text for run in paragraph.runs)
        original_text = full_text

        for k, v in replacements.items():
            full_text = full_text.replace(k, str(v))

        if full_text != original_text:
            for run in paragraph.runs:
                run.text = ""
            paragraph.runs[0].text = full_text

    def replace_in_table(table, replacements):
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    replace_in_paragraph(paragraph, replacements)

                for nested_table in cell.tables:
                    replace_in_table(nested_table, replacements)

    for paragraph in doc.paragraphs:
        replace_in_paragraph(paragraph, data)

    for table in doc.tables:
        replace_in_table(table, data)

    for section in doc.sections:
        for paragraph in section.header.paragraphs:
            replace_in_paragraph(paragraph, data)

        for table in section.header.tables:
            replace_in_table(table, data)

        for paragraph in section.footer.paragraphs:
            replace_in_paragraph(paragraph, data)

        for table in section.footer.tables:
            replace_in_table(table, data)


def tick(value):
    return "☑" if value else "☐"


def safe_text(value):
    return "" if value is None else str(value)


def format_date_ddmmyyyy(value):
    """Return dates consistently as DD/MM/YYYY across the EWMT app."""
    if value is None or value == "":
        return ""

    if isinstance(value, (date, datetime)):
        return value.strftime("%d/%m/%Y")

    text = str(value).strip()
    if not text:
        return ""

    # Accept common existing / legacy date formats, but always output DD/MM/YYYY.
    for fmt in (
        "%d/%m/%Y",
        "%Y-%m-%d",
        "%d-%m-%Y",
        "%d.%m.%Y",
        "%Y/%m/%d",
        "%Y.%m.%d",
    ):
        try:
            return datetime.strptime(text, fmt).strftime("%d/%m/%Y")
        except ValueError:
            pass

    # If it is not a recognised date, preserve the user's text rather than failing.
    return text


def clean_ms_text(text):
    banned_phrases = [
        "The above equipment selection",
        "These safety controls",
        "The above sequence",
        "This sequence",
        "This method statement",
        "consistent with Eric Wong Machinery Transportation Pte Ltd",
        "company’s established method statement style",
        "company's established method statement style",
        "standard precautions repeatedly stated",
        "reflect the standard precautions",
        "follows the company",
        "previous method statements",
        "prior method statements",
    ]

    lines = str(text).splitlines()
    cleaned = []

    for line in lines:
        if not any(phrase.lower() in line.lower() for phrase in banned_phrases):
            cleaned.append(line)

    return "\n".join(cleaned).strip()


def set_ra_cell_text(cell, text):
    cell.text = ""
    p = cell.paragraphs[0]

    for i, line in enumerate(str(text).split("\n")):
        if i > 0:
            p.add_run().add_break()

        run = p.add_run(line)
        run.font.name = "Times New Roman"
        run._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
        run.font.size = Pt(10)


def format_risk_assessment(doc):
    for para in doc.paragraphs:
        for run in para.runs:
            run.font.name = "Times New Roman"
            run._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
            run.font.size = Pt(10)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        run.font.name = "Times New Roman"
                        run._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
                        run.font.size = Pt(10)


def find_ra_table(doc):
    for table in doc.tables:
        full = " ".join(c.text for row in table.rows for c in row.cells)
        if "Hazard Identification" in full and "Risk Evaluation" in full and "Risk Control" in full:
            return table
    return None


def find_ra_column_header_row(table):
    for i, row in enumerate(table.rows):
        texts = [cell.text.strip() for cell in row.cells]
        if "Ref" in texts and "Work Activity" in texts and "Hazard" in texts:
            return i
    return None


def clear_rows_after_column_header(table):
    header_index = find_ra_column_header_row(table)

    if header_index is None:
        raise Exception("Cannot find RA column header row")

    while len(table.rows) > header_index + 1:
        row = table.rows[header_index + 1]
        row._element.getparent().remove(row._element)


def add_ra_row(table, values):
    row = table.add_row().cells
    for i, v in enumerate(values):
        if i < len(row):
            set_ra_cell_text(row[i], v)


def merge_same_work_activity_cells(table):
    header_index = find_ra_column_header_row(table)
    if header_index is None:
        return

    start_row = header_index + 1
    end_row = len(table.rows) - 1
    current_start = start_row

    while current_start <= end_row:
        activity = table.rows[current_start].cells[1].text.strip()
        current_end = current_start

        while (
            current_end + 1 <= end_row
            and table.rows[current_end + 1].cells[1].text.strip() == ""
        ):
            current_end += 1

        if activity and current_end > current_start:
            table.rows[current_start].cells[0].merge(table.rows[current_end].cells[0])
            table.rows[current_start].cells[1].merge(table.rows[current_end].cells[1])

        current_start = current_end + 1


def find_inventory_table(doc):
    for table in doc.tables:
        full = " ".join(c.text for row in table.rows for c in row.cells)
        if "Ref No." in full and "Location" in full and "Process" in full and "S/No." in full and "Work Activity" in full:
            return table
    return None


def fill_inventory_table(doc, activities_text, location, process):
    table = find_inventory_table(doc)

    if table is None:
        return

    activity_list = [
        a.strip()
        for a in activities_text.split("\n")
        if a.strip()
    ]

    start_row = None
    for i, row in enumerate(table.rows):
        row_text = " ".join(cell.text.strip() for cell in row.cells)
        if "S/No." in row_text and "Work Activity" in row_text:
            start_row = i + 1
            break

    if start_row is None:
        return

    for idx, activity in enumerate(activity_list, start=1):
        row_index = start_row + idx - 1

        if row_index >= len(table.rows):
            table.add_row()

        row = table.rows[row_index].cells

        if len(row) >= 6:
            set_ra_cell_text(row[0], "")
            set_ra_cell_text(row[1], location if idx == 1 else "")
            set_ra_cell_text(row[2], process if idx == 1 else "")
            set_ra_cell_text(row[3], str(idx))
            set_ra_cell_text(row[4], activity)
            set_ra_cell_text(row[5], "")


def certificate_browser(folder_name, title, info_text, search_label, search_placeholder, download_label):
    st.markdown(f"## {title}")
    st.info(info_text)

    cert_folder = os.path.join(BASE_DIR, folder_name)

    if not os.path.exists(cert_folder):
        st.error(f"Folder not found: {folder_name}")
        st.code(folder_name)
        st.info(f"Create this folder in your GitHub project and upload files inside.")
        return

    files = [
        f for f in os.listdir(cert_folder)
        if f.lower().endswith((".pdf", ".png", ".jpg", ".jpeg", ".docx"))
    ]

    files = sorted(files)

    if not files:
        st.warning(f"No files found inside {folder_name} folder.")
        return

    st.success(f"Found {len(files)} file(s).")

    search = st.text_input(
        search_label,
        "",
        placeholder=search_placeholder
    )

    filtered_files = files

    if search:
        search_words = search.lower().split()
        filtered_files = [
            f for f in files
            if all(word in f.lower() for word in search_words)
        ]

    if not filtered_files:
        st.warning("No matching file found.")
        return

    st.success(f"Found {len(filtered_files)} matching file(s).")

    selected_file = st.selectbox("Choose file", filtered_files)
    file_path = os.path.join(cert_folder, selected_file)

    st.write("Selected file:")
    st.code(selected_file)

    with open(file_path, "rb") as f:
        file_bytes = f.read()

    mime_type = "application/octet-stream"

    if selected_file.lower().endswith(".pdf"):
        mime_type = "application/pdf"
    elif selected_file.lower().endswith(".docx"):
        mime_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    elif selected_file.lower().endswith(".png"):
        mime_type = "image/png"
    elif selected_file.lower().endswith((".jpg", ".jpeg")):
        mime_type = "image/jpeg"

    st.download_button(
        download_label,
        file_bytes,
        file_name=selected_file,
        mime=mime_type
    )

    if selected_file.lower().endswith((".png", ".jpg", ".jpeg")):
        st.image(file_path, caption=selected_file)

    if selected_file.lower().endswith(".pdf"):
        st.markdown("### PDF Preview")
        base64_pdf = base64.b64encode(file_bytes).decode("utf-8")

        st.markdown(
            f"""
            <a href="data:application/pdf;base64,{base64_pdf}"
               target="_blank"
               style="
                   display:inline-block;
                   background:#1e3a8a;
                   color:white;
                   padding:12px 20px;
                   border-radius:10px;
                   text-decoration:none;
                   font-weight:700;
               ">
               Open PDF Preview in New Tab
            </a>
            """,
            unsafe_allow_html=True
        )

        st.info("Chrome blocks embedded PDF preview in Streamlit. Click the button above to preview before downloading.")

    if selected_file.lower().endswith(".docx"):
        st.info("Word document preview is not supported inside Streamlit. Please download the file to view.")

# =====================
# GITHUB DOCUMENT MANAGEMENT
# =====================
def get_secret(name, default=None):
    """Safely read a Streamlit secret without exposing it in the UI."""
    try:
        return st.secrets[name]
    except Exception:
        return default


def github_settings():
    return {
        "token": get_secret("GITHUB_TOKEN", ""),
        "repo": get_secret("GITHUB_REPO", ""),
        "branch": get_secret("GITHUB_BRANCH", "main"),
    }


def github_headers():
    cfg = github_settings()
    return {
        "Authorization": f"Bearer {cfg['token']}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
        "User-Agent": "EWMT-Streamlit-Document-Manager",
    }


def github_api_url(path=""):
    cfg = github_settings()
    repo = cfg["repo"].strip().strip("/")
    safe_path = quote(path.strip("/"), safe="/")
    base = f"https://api.github.com/repos/{repo}/contents"
    return f"{base}/{safe_path}" if safe_path else base


def validate_github_config():
    cfg = github_settings()
    missing = []

    if not get_secret("ADMIN_PASSWORD", ""):
        missing.append("ADMIN_PASSWORD")
    if not cfg["token"]:
        missing.append("GITHUB_TOKEN")
    if not cfg["repo"]:
        missing.append("GITHUB_REPO")

    return missing


def github_get_file(path):
    """Return GitHub file metadata or None when the file does not exist."""
    cfg = github_settings()
    response = requests.get(
        github_api_url(path),
        headers=github_headers(),
        params={"ref": cfg["branch"]},
        timeout=30,
    )

    if response.status_code == 404:
        return None

    response.raise_for_status()
    return response.json()


def github_list_folder(folder_path):
    cfg = github_settings()
    response = requests.get(
        github_api_url(folder_path),
        headers=github_headers(),
        params={"ref": cfg["branch"]},
        timeout=30,
    )

    if response.status_code == 404:
        return []

    response.raise_for_status()
    data = response.json()

    if isinstance(data, dict):
        data = [data]

    return sorted(
        [item for item in data if item.get("type") == "file"],
        key=lambda item: item.get("name", "").lower(),
    )


def github_upload_file(folder_path, filename, file_bytes, replace_existing=False):
    cfg = github_settings()
    filename = os.path.basename(filename).strip()

    if not filename:
        raise ValueError("Invalid file name.")

    target_path = f"{folder_path.strip('/')}/{filename}"
    existing = github_get_file(target_path)

    if existing and not replace_existing:
        raise FileExistsError(
            "A file with the same name already exists. Tick 'Replace existing file' to overwrite it."
        )

    payload = {
        "message": f"EWMT document upload: {target_path}",
        "content": base64.b64encode(file_bytes).decode("utf-8"),
        "branch": cfg["branch"],
    }

    if existing:
        payload["sha"] = existing["sha"]

    response = requests.put(
        github_api_url(target_path),
        headers=github_headers(),
        json=payload,
        timeout=60,
    )

    if response.status_code not in (200, 201):
        try:
            detail = response.json().get("message", response.text)
        except Exception:
            detail = response.text
        raise RuntimeError(f"GitHub upload failed: {detail}")

    return response.json()


def github_delete_file(folder_path, filename, sha):
    cfg = github_settings()
    filename = os.path.basename(filename).strip()
    target_path = f"{folder_path.strip('/')}/{filename}"

    payload = {
        "message": f"EWMT document delete: {target_path}",
        "sha": sha,
        "branch": cfg["branch"],
    }

    response = requests.delete(
        github_api_url(target_path),
        headers=github_headers(),
        json=payload,
        timeout=60,
    )

    if response.status_code != 200:
        try:
            detail = response.json().get("message", response.text)
        except Exception:
            detail = response.text
        raise RuntimeError(f"GitHub delete failed: {detail}")

    return response.json()


def admin_is_logged_in():
    return bool(st.session_state.get("admin_authenticated", False))


def admin_login_form():
    st.markdown("### 🔐 Administrator Access")
    st.caption("Only authorised administrators can upload, replace or delete company documents.")

    with st.form("admin_login_form", clear_on_submit=False):
        password = st.text_input("Administrator Password", type="password")
        login = st.form_submit_button("🔓 Login as Administrator", use_container_width=True)

    if login:
        saved_password = str(get_secret("ADMIN_PASSWORD", ""))

        if not saved_password:
            st.error("ADMIN_PASSWORD is not configured in Streamlit Secrets.")
            return False

        if hmac.compare_digest(str(password), saved_password):
            st.session_state.admin_authenticated = True
            st.success("Administrator access granted.")
            st.rerun()
        else:
            st.error("Incorrect administrator password.")

    return False


def render_admin_document_manager():
    st.markdown("### 📁 Admin Document Manager")
    st.caption(
        "Upload new certificates, replace renewed certificates and delete obsolete files. "
        "Changes are committed directly to your GitHub repository."
    )

    missing = validate_github_config()
    if missing:
        st.error("Missing Streamlit Secrets: " + ", ".join(missing))
        st.info(
            "Add ADMIN_PASSWORD, GITHUB_TOKEN, GITHUB_REPO and optionally GITHUB_BRANCH "
            "in Streamlit Cloud → App settings → Secrets."
        )
        return

    cfg = github_settings()

    top_left, top_right = st.columns([3, 1])
    with top_left:
        st.info(f"Repository: {cfg['repo']}  •  Branch: {cfg['branch']}")
    with top_right:
        if st.button("🔒 Logout", key="admin_logout", use_container_width=True):
            st.session_state.admin_authenticated = False
            st.rerun()

    folder_map = {
        "🧰 Lifting Gear Certificates": "Lifting Gears Certificate",
        "👷 Worker Training Certificates": "Workers Certificate",
    }

    selected_type = st.selectbox(
        "Document Category",
        list(folder_map.keys()),
        key="admin_document_category",
    )
    selected_folder = folder_map[selected_type]

    st.markdown(f"#### Current Files — `{selected_folder}`")

    try:
        github_files = github_list_folder(selected_folder)
    except Exception as e:
        st.error("Unable to read documents from GitHub.")
        st.exception(e)
        github_files = []

    if github_files:
        search_admin = st.text_input(
            "Search current files",
            placeholder="Type worker name, sling, shackle, expiry date...",
            key=f"admin_search_{selected_folder}",
        )

        shown_files = github_files
        if search_admin:
            words = search_admin.lower().split()
            shown_files = [
                item for item in github_files
                if all(word in item.get("name", "").lower() for word in words)
            ]

        st.caption(f"{len(shown_files)} file(s) shown • {len(github_files)} total")

        if not shown_files:
            st.warning("No matching file found.")
        else:
            for item in shown_files:
                filename = item.get("name", "")
                size_bytes = int(item.get("size", 0) or 0)
                size_kb = size_bytes / 1024

                with st.container(border=True):
                    c1, c2 = st.columns([5, 1])
                    with c1:
                        st.markdown(f"**{filename}**")
                        st.caption(f"{size_kb:,.1f} KB")
                    with c2:
                        if item.get("html_url"):
                            st.link_button(
                                "View",
                                item["html_url"],
                                use_container_width=True,
                            )

                    delete_confirm = st.checkbox(
                        f"I confirm I want to permanently delete {filename}",
                        key=f"delete_confirm_{selected_folder}_{item.get('sha', filename)}",
                    )

                    if st.button(
                        "🗑 Delete Permanently",
                        key=f"delete_{selected_folder}_{item.get('sha', filename)}",
                        disabled=not delete_confirm,
                        use_container_width=True,
                    ):
                        try:
                            with st.spinner(f"Deleting {filename}..."):
                                github_delete_file(selected_folder, filename, item["sha"])
                            st.success(f"Deleted: {filename}")
                            st.cache_data.clear()
                            st.rerun()
                        except Exception as e:
                            st.error("Delete failed.")
                            st.exception(e)
    else:
        st.warning("No documents found in this GitHub folder, or the folder does not exist yet.")

    st.markdown("---")
    st.markdown("#### ⬆ Upload / Replace Document")

    allowed_types = ["pdf", "png", "jpg", "jpeg", "docx"]
    uploaded_file = st.file_uploader(
        "Choose document",
        type=allowed_types,
        key=f"admin_upload_{selected_folder}",
    )

    if uploaded_file is not None:
        default_name = os.path.basename(uploaded_file.name)
        upload_name = st.text_input(
            "File name in GitHub",
            value=default_name,
            key=f"admin_filename_{selected_folder}",
        )

        replace_existing = st.checkbox(
            "Replace existing file if the same filename already exists",
            value=False,
            key=f"admin_replace_{selected_folder}",
        )

        st.info(
            "Tip: for lifting gear expiry alerts, include the expiry date in the filename, "
            "for example: 10 Ton Shackle Expiry 31-08-2027.pdf"
        )

        if st.button(
            "⬆ Upload Document",
            key=f"admin_upload_button_{selected_folder}",
            use_container_width=True,
        ):
            clean_name = os.path.basename(upload_name).strip()
            extension = clean_name.rsplit(".", 1)[-1].lower() if "." in clean_name else ""

            if not clean_name:
                st.error("Please enter a valid file name.")
            elif extension not in allowed_types:
                st.error("Allowed file types: PDF, PNG, JPG, JPEG and DOCX.")
            else:
                try:
                    with st.spinner(f"Uploading {clean_name}..."):
                        github_upload_file(
                            selected_folder,
                            clean_name,
                            uploaded_file.getvalue(),
                            replace_existing=replace_existing,
                        )
                    st.success(f"Uploaded successfully: {clean_name}")
                    st.info(
                        "GitHub has been updated. Streamlit Cloud normally redeploys automatically "
                        "after the repository changes."
                    )
                    st.cache_data.clear()
                except FileExistsError as e:
                    st.warning(str(e))
                except Exception as e:
                    st.error("Upload failed.")
                    st.exception(e)

# =====================
# METHOD STATEMENT PRO - POWERPOINT MASTER HELPERS
# =====================
def mos_pro_find_master_template():
    """Return the first available MOS PRO PowerPoint master template."""
    for path in MOS_PRO_TEMPLATE_CANDIDATES:
        if os.path.exists(path):
            return path
    return None


def mos_pro_replace_text_in_paragraph(paragraph, replacements):
    """Replace text while keeping the first run's original PowerPoint formatting."""
    if not getattr(paragraph, "runs", None):
        return 0

    original = "".join(run.text for run in paragraph.runs)
    updated = original

    for pattern, replacement in replacements:
        updated = re.sub(pattern, replacement, updated, flags=re.IGNORECASE)

    if updated == original:
        return 0

    first_run = paragraph.runs[0]
    for run in paragraph.runs:
        run.text = ""
    first_run.text = updated
    return 1


def mos_pro_replace_project_information(prs, data):
    """
    Update only the repeated project-information table in the master PPTX.
    Standard procedure wording, drawings and existing slide contents are untouched.
    """
    replacements = [
        (r"(Customer\s*/\s*Tenant\s*Company\s*:\s*)[^\r\n]*", rf"\g<1>{safe_text(data.get('customer'))}"),
        (r"(Site\s*Location\s*:\s*)[^\r\n]*", rf"\g<1>{safe_text(data.get('location'))}"),
        (r"(Process\s*:\s*)[^\r\n]*", rf"\g<1>{safe_text(data.get('process'))}"),
        (r"(Prepared\s*by\s*:\s*)[^\r\n]*", rf"\g<1>{safe_text(data.get('prepared_by'))}"),
        (r"(Approved\s*By\s*:\s*)[^\r\n]*", rf"\g<1>{safe_text(data.get('approved_by'))}"),
        (r"(Last\s*Review\s*Date\s*:\s*)[^\r\n]*", rf"\g<1>{safe_text(data.get('last_review'))}"),
        (r"(Next\s*Review\s*Date\s*:\s*)[^\r\n]*", rf"\g<1>{safe_text(data.get('next_review'))}"),
    ]

    changed = 0

    def process_shape(shape):
        nonlocal changed

        # Recurse into groups, including the EWMT header group.
        if MSO_SHAPE_TYPE is not None and shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            for child in shape.shapes:
                process_shape(child)

        if getattr(shape, "has_text_frame", False):
            for paragraph in shape.text_frame.paragraphs:
                changed += mos_pro_replace_text_in_paragraph(paragraph, replacements)

        if getattr(shape, "has_table", False):
            for row in shape.table.rows:
                for cell in row.cells:
                    for paragraph in cell.text_frame.paragraphs:
                        changed += mos_pro_replace_text_in_paragraph(paragraph, replacements)

    for slide in prs.slides:
        for shape in slide.shapes:
            process_shape(shape)

    return changed


def mos_pro_find_work_method_shape(prs):
    """Locate the existing 'Work Method Statement for Lifting Operation' text box."""
    for slide_index, slide in enumerate(prs.slides):
        for shape in slide.shapes:
            if getattr(shape, "has_text_frame", False):
                txt = shape.text or ""
                if "Work Method Statement for Lifting Operation" in txt:
                    return slide_index, shape
    return None, None


def mos_pro_find_work_method_shape_on_slide(slide):
    """Locate the work-method textbox on a specific slide."""
    for shape in slide.shapes:
        if getattr(shape, "has_text_frame", False):
            txt = shape.text or ""
            if "Work Method Statement for Lifting Operation" in txt:
                return shape
    return None


def mos_pro_duplicate_slide_after(prs, source_index):
    """
    Duplicate one complete MOS slide and place the copy immediately after it.

    This preserves the EWMT header, footer, logo, project-information table
    and page design. Image and hyperlink relationships are remapped so the
    duplicated slide remains valid.
    """
    source_slide = prs.slides[source_index]
    new_slide = prs.slides.add_slide(source_slide.slide_layout)

    # Remove any placeholders automatically created by the selected layout.
    for shape in list(new_slide.shapes):
        element = shape.element
        element.getparent().remove(element)

    rel_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    relationship_map = {}

    # Create equivalent relationships on the new slide, excluding slide layout.
    for old_rid, rel in source_slide.part.rels.items():
        if rel.reltype.endswith("/slideLayout"):
            continue

        relationship_map[old_rid] = new_slide.part.rels._add_relationship(
            rel.reltype,
            rel._target,
            rel.is_external,
        )

    # Deep-copy every shape and remap relationship IDs used by images/hyperlinks.
    for shape in source_slide.shapes:
        new_element = deepcopy(shape.element)

        for node in new_element.iter():
            for attr_name, attr_value in list(node.attrib.items()):
                if (
                    attr_name.startswith("{" + rel_ns + "}")
                    and attr_value in relationship_map
                ):
                    node.set(attr_name, relationship_map[attr_value])

        new_slide.shapes._spTree.insert_element_before(
            new_element,
            "p:extLst"
        )

    # New slides are appended to the end by python-pptx.
    # Move this one so it sits directly after the original work-method page.
    slide_id_list = prs.slides._sldIdLst
    new_slide_id = slide_id_list[-1]
    slide_id_list.remove(new_slide_id)
    slide_id_list.insert(source_index + 1, new_slide_id)

    return new_slide


def mos_pro_get_slide_number_text(slide):
    """Read the existing MOS page number from its slide-number placeholder."""
    for shape in slide.shapes:
        if "Slide Number Placeholder" in getattr(shape, "name", ""):
            txt = (shape.text or "").strip()
            if txt:
                return txt

    # Fallback: look for a small textbox containing only a page number.
    for shape in slide.shapes:
        if getattr(shape, "has_text_frame", False):
            txt = (shape.text or "").strip()
            if re.fullmatch(r"\d+[A-Za-z]?", txt):
                return txt

    return ""


def mos_pro_set_slide_number_text(slide, value):
    """
    Set only the duplicated page's footer number.

    The continuation page uses e.g. 15A / 26A so the original master page
    numbering after it does not need to be changed.
    """
    for shape in slide.shapes:
        if "Slide Number Placeholder" in getattr(shape, "name", ""):
            if getattr(shape, "has_text_frame", False):
                shape.text_frame.clear()
                p = shape.text_frame.paragraphs[0]
                p.text = str(value)
                return True

    return False


def mos_pro_set_run_font(run, size=10.0, bold=False, color=None):
    """Apply clean MOS PRO body formatting."""
    run.font.name = "Arial"
    run.font.size = PPTXPt(size)
    run.font.bold = bold

    if color is not None and RGBColor is not None:
        run.font.color.rgb = RGBColor(*color)


def mos_pro_add_paragraph(
    tf,
    text,
    size=10.0,
    bold=False,
    color=None,
    space_before=0,
    space_after=4,
):
    """Add one consistently formatted paragraph to the work-method textbox."""
    p = tf.add_paragraph()
    p.text = str(text)
    p.level = 0
    p.space_before = PPTXPt(space_before)
    p.space_after = PPTXPt(space_after)

    if p.runs:
        mos_pro_set_run_font(
            p.runs[0],
            size=size,
            bold=bold,
            color=color,
        )

    return p


def mos_pro_write_work_method_page(
    shape,
    page_title,
    sections,
    start_number=1,
    stop_work=None,
):
    """
    Write ONE clean work-method page.

    The page contains clear section headings and short numbered steps instead
    of compressing a long 10-18 step paragraph into one tiny textbox.
    """
    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = True

    title = tf.paragraphs[0]
    title.text = page_title
    title.space_after = PPTXPt(7)

    if title.runs:
        mos_pro_set_run_font(
            title.runs[0],
            size=13.0,
            bold=True,
        )

    step_no = start_number

    for section_title, steps in sections:
        clean_steps = [
            str(step).strip()
            for step in (steps or [])
            if str(step).strip()
        ]

        if not clean_steps:
            continue

        mos_pro_add_paragraph(
            tf,
            section_title,
            size=10.8,
            bold=True,
            color=(31, 78, 121),
            space_before=6,
            space_after=3,
        )

        for step in clean_steps:
            # The app controls numbering so accidental AI numbering is removed.
            clean_step = re.sub(
                r"^\s*\d+[\.\)]\s*",
                "",
                step,
            ).strip()

            mos_pro_add_paragraph(
                tf,
                f"{step_no}. {clean_step}",
                size=9.8,
                bold=False,
                space_before=0,
                space_after=4,
            )
            step_no += 1

    if stop_work:
        mos_pro_add_paragraph(
            tf,
            "STOP WORK",
            size=10.8,
            bold=True,
            color=(192, 0, 0),
            space_before=8,
            space_after=2,
        )

        mos_pro_add_paragraph(
            tf,
            stop_work,
            size=9.8,
            bold=True,
            color=(192, 0, 0),
            space_before=0,
            space_after=0,
        )

    return step_no


def mos_pro_write_work_method_two_pages(prs, slide_index, generated):
    """
    Convert the original crowded Work Method page into TWO clean pages.

    PAGE 1
    - A. Preparation
    - B. Equipment Set-Up / Rigging

    PAGE 2
    - C. Hoisting / Machinery Movement
    - D. Final Positioning / Completion
    - STOP WORK

    The second page is a duplicate of the master page, so EWMT branding,
    header/footer and project-information formatting remain unchanged.
    """
    first_slide = prs.slides[slide_index]
    first_shape = mos_pro_find_work_method_shape_on_slide(first_slide)

    if first_shape is None:
        raise RuntimeError(
            "Could not locate the existing Work Method Statement textbox."
        )

    source_page_number = mos_pro_get_slide_number_text(first_slide)

    second_slide = mos_pro_duplicate_slide_after(
        prs,
        slide_index,
    )

    second_shape = mos_pro_find_work_method_shape_on_slide(second_slide)

    if second_shape is None:
        raise RuntimeError(
            "Could not locate the Work Method textbox on the continuation page."
        )

    # Keep the original document numbering stable.
    # Example: page 15 becomes 15 + 15A, while the next master page stays page 16.
    if source_page_number:
        mos_pro_set_slide_number_text(
            second_slide,
            f"{source_page_number}A",
        )

    page1_sections = [
        (
            "A. PREPARATION",
            generated.get("preparation", []),
        ),
        (
            "B. EQUIPMENT SET-UP / RIGGING",
            generated.get("equipment_setup_rigging", []),
        ),
    ]

    page2_sections = [
        (
            "C. HOISTING / MACHINERY MOVEMENT",
            generated.get("lifting_movement", []),
        ),
        (
            "D. FINAL POSITIONING / COMPLETION",
            generated.get("completion", []),
        ),
    ]

    next_step_number = mos_pro_write_work_method_page(
        first_shape,
        "Work Method Statement for Lifting Operation",
        page1_sections,
        start_number=1,
    )

    mos_pro_write_work_method_page(
        second_shape,
        "Work Method Statement for Lifting Operation — Continued",
        page2_sections,
        start_number=next_step_number,
        stop_work=generated.get("stop_work", ""),
    )

    first_page_label = source_page_number or str(slide_index + 1)
    second_page_label = (
        f"{source_page_number}A"
        if source_page_number
        else f"{slide_index + 1}A"
    )

    return first_page_label, second_page_label


def mos_pro_generate_work_method(
    description,
    machine,
    equipment,
    site_notes,
    operation_type,
):
    """
    Generate a concise two-page project work method.

    The existing Method Statement vector store is still searched for relevant
    historical methodology, but old customer/project-specific details must not
    be copied into the new job.
    """
    prompt = f"""
Prepare PROJECT-SPECIFIC work-method steps for the EWMT professional Method Statement PowerPoint.

Before generating the method, use the EWMT Method Statement vector store to study relevant previous machinery-moving and lifting cases.

CURRENT JOB
Operation type: {operation_type}
Description of work: {description}
Machine / load: {machine}
Equipment actually intended: {equipment}
Site / access notes: {site_notes}

IMPORTANT RULES
- Use previous vector-store cases only for methodology, practical work sequence and relevant controls.
- NEVER copy old customer names, old site names, dates, prices, crane registration numbers, worker names or machine details from previous jobs.
- Use ONLY equipment stated for this current job.
- Do not invent cranes, forklifts, slings, shackles, spreader beams, hydraulic jacks, machine skates or other equipment.
- Keep each step SHORT, practical and specific.
- Each step should normally be about 12 to 22 words.
- Do not write paragraph-length steps.
- Do not repeat the general WSH wording already contained in the master MOS.
- Total work sequence should normally contain approximately 10 to 14 steps.
- Include a site briefing / toolbox meeting before commencement.
- Include barricading / exclusion-zone control when relevant.
- If crane / lorry-loader lifting is actually stated, include relevant setup, load/radius/SWL verification, rigging inspection, trial lift and controlled lifting.
- If floor shifting is involved, use only the floor-moving equipment actually stated by the user.
- Include final positioning, de-rigging and housekeeping.
- STOP WORK must be one short statement only.

Return JSON only with these exact fields:

preparation
- Array of 3 to 4 short steps.

equipment_setup_rigging
- Array of 2 to 4 short steps.
- Return an empty array when this section is not applicable.

lifting_movement
- Array of 3 to 4 short steps.

completion
- Array of 2 to 3 short steps.

stop_work
- One concise STOP WORK statement.
"""

    response = client.responses.create(
        model="gpt-5.4",
        input=prompt,
        tools=[{
            "type": "file_search",
            "vector_store_ids": [MS_VECTOR_STORE_ID],
        }],
        text={
            "format": {
                "type": "json_schema",
                "name": "mos_pro_clean_work_method_schema",
                "schema": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "preparation": {
                            "type": "array",
                            "items": {"type": "string"},
                        },
                        "equipment_setup_rigging": {
                            "type": "array",
                            "items": {"type": "string"},
                        },
                        "lifting_movement": {
                            "type": "array",
                            "items": {"type": "string"},
                        },
                        "completion": {
                            "type": "array",
                            "items": {"type": "string"},
                        },
                        "stop_work": {
                            "type": "string",
                        },
                    },
                    "required": [
                        "preparation",
                        "equipment_setup_rigging",
                        "lifting_movement",
                        "completion",
                        "stop_work",
                    ],
                },
            },
        },
    )

    result = json.loads(response.output_text)

    # Final defensive cleanup. The structured JSON means each item is already
    # separate, but remove unwanted "company format" commentary if it appears.
    for key in (
        "preparation",
        "equipment_setup_rigging",
        "lifting_movement",
        "completion",
    ):
        result[key] = [
            clean_ms_text(step)
            for step in result.get(key, [])
            if clean_ms_text(step)
        ]

    result["stop_work"] = clean_ms_text(
        result.get("stop_work", "")
    )

    return result


def mos_pro_build_powerpoint(project_data, replace_work_method=False, work_method_data=None):
    if not PPTX_AVAILABLE:
        raise RuntimeError("python-pptx is not installed. Add python-pptx to requirements.txt and redeploy Streamlit.")

    master_path = mos_pro_find_master_template()
    if not master_path:
        raise FileNotFoundError(
            "MOS PRO master template not found. Upload 'MOS New.pptx' (or 'MOS New(1).pptx') into the Templates folder."
        )

    prs = Presentation(master_path)
    changed_fields = mos_pro_replace_project_information(prs, project_data)

    replaced_work_method = False
    work_method_slide = None

    if replace_work_method:
        work_method_data = work_method_data or {}

        generated = mos_pro_generate_work_method(
            work_method_data.get("description", ""),
            work_method_data.get("machine", ""),
            work_method_data.get("equipment", ""),
            work_method_data.get("site_notes", ""),
            work_method_data.get("operation_type", ""),
        )

        slide_index, shape = mos_pro_find_work_method_shape(prs)

        if shape is None:
            raise RuntimeError(
                "Could not locate the existing Work Method Statement page in the master PowerPoint."
            )

        first_page, continuation_page = mos_pro_write_work_method_two_pages(
            prs,
            slide_index,
            generated,
        )

        replaced_work_method = True
        work_method_slide = f"{first_page} / {continuation_page}"

    output = BytesIO()
    prs.save(output)
    output.seek(0)

    return output, {
        "master": os.path.basename(master_path),
        "slides": len(prs.slides),
        "changed_fields": changed_fields,
        "work_method_replaced": replaced_work_method,
        "work_method_slide": work_method_slide,
    }


# =====================
# DASHBOARD COUNT FUNCTIONS
# =====================
def count_files_in_folder(folder_name, allowed_ext=(".pdf", ".png", ".jpg", ".jpeg", ".docx")):
    folder_path = os.path.join(BASE_DIR, folder_name)

    if not os.path.exists(folder_path):
        return 0

    return len([
        f for f in os.listdir(folder_path)
        if f.lower().endswith(allowed_ext)
    ])


def get_lifting_gear_expiry_counts(alert_days=30):
    import re

    cert_folder = os.path.join(BASE_DIR, "Lifting Gears Certificate")

    counts = {
        "expired": 0,
        "expiring_soon": 0,
        "valid": 0,
        "unknown": 0
    }

    if not os.path.exists(cert_folder):
        return counts

    files = [
        f for f in os.listdir(cert_folder)
        if f.lower().endswith((".pdf", ".png", ".jpg", ".jpeg"))
    ]

    today = date.today()

    patterns = [
        r"(\d{4})[-_\\.](\d{1,2})[-_\\.](\d{1,2})",
        r"(\d{1,2})[-_\\.](\d{1,2})[-_\\.](\d{4})",
    ]

    for f in files:
        found_date = None

        for pattern in patterns:
            match = re.search(pattern, f)

            if match:
                try:
                    parts = match.groups()

                    if len(parts[0]) == 4:
                        found_date = date(int(parts[0]), int(parts[1]), int(parts[2]))
                    else:
                        found_date = date(int(parts[2]), int(parts[1]), int(parts[0]))

                    break
                except Exception:
                    found_date = None

        if not found_date:
            counts["unknown"] += 1
        else:
            days_left = (found_date - today).days

            if days_left < 0:
                counts["expired"] += 1
            elif days_left <= alert_days:
                counts["expiring_soon"] += 1
            else:
                counts["valid"] += 1

    return counts

# =====================
# SIDEBAR NAVIGATION
# =====================
with st.sidebar:
    st.markdown("## EWMT System")
    st.markdown("AI Document Control")
    st.markdown("---")

    if st.button("🏠 Dashboard", key="side_dashboard"):
        go_to_page("🏠 Dashboard")

    if st.button("📄 Method Statement", key="side_method_statement"):
        go_to_page("📄 Method Statement")

    if st.button("📑 Method Statement PRO", key="side_method_statement_pro"):
        go_to_page("📑 Method Statement PRO")

    if st.button("🏗️ Lifting Plan", key="side_lifting_plan"):
        go_to_page("🏗️ Lifting Plan")

    if st.button("⚠️ Risk Assessment Pro", key="side_risk_assessment"):
        go_to_page("⚠️ Risk Assessment Pro")

    if st.button("🧰 Lifting Gear Register", key="side_lifting_gear"):
        go_to_page("🧰 Lifting Gear Register")

    if st.button("👷 Worker Training Certificate", key="side_worker_training"):
        go_to_page("👷 Worker Training Certificate")

    if st.button("⏰ Expiry Alerts", key="side_expiry_alerts"):
        go_to_page("⏰ Expiry Alerts")

    if st.button("⚙️ Settings", key="side_settings"):
        go_to_page("⚙️ Settings")

    page = st.session_state.page

    st.markdown("---")
    st.caption("Internal system for document preparation and lifting operation records.")
    
# ======================================================
# DASHBOARD
# ======================================================
if page == "🏠 Dashboard":
    lifting_gear_count = count_files_in_folder("Lifting Gears Certificate")
    worker_cert_count = count_files_in_folder("Workers Certificate")
    expiry_counts = get_lifting_gear_expiry_counts(alert_days=30)

    expired_count = expiry_counts["expired"]
    expiring_soon_count = expiry_counts["expiring_soon"]
    valid_count = expiry_counts["valid"]

    st.markdown("""
    <div class="section-title">EWMT AI Document Control Dashboard</div>
    <div class="section-caption">
        Professional document generation, lifting operation records and certificate control system.
    </div>
    """, unsafe_allow_html=True)

    m1, m2, m3, m4 = st.columns(4)

    with m1:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">Document Modules</div>
            <div class="metric-value">4</div>
            <div class="metric-small">Method Statement / MOS PRO / RA / Lifting Plan</div>
        </div>
        """, unsafe_allow_html=True)

    with m2:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">Gear Records</div>
            <div class="metric-value">{lifting_gear_count}</div>
            <div class="metric-small">Files in Lifting Gears Certificate folder</div>
        </div>
        """, unsafe_allow_html=True)

    with m3:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">Expiring Soon</div>
            <div class="metric-value">{expiring_soon_count}</div>
            <div class="metric-small">Within next 30 days</div>
        </div>
        """, unsafe_allow_html=True)

    with m4:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">Worker Certificates</div>
            <div class="metric-value">{worker_cert_count}</div>
            <div class="metric-small">Files in Workers Certificate folder</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown('<div class="section-title">Document Modules</div>', unsafe_allow_html=True)

    def dashboard_card(title, desc, image_b64, tag):
        st.markdown(f"""
        <div class="dashboard-card">
            <div class="dashboard-img" style='background-image:url("data:image/jpg;base64,{image_b64}")'>
                <div class="dashboard-pill">{tag}</div>
            </div>
            <div class="dashboard-content">
                <div class="card-accent"></div>
                <h3>{title}</h3>
                <p>{desc}</p>
            </div>
        </div>
        """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns(3)

    with col1:
        dashboard_card(
            "📄 Method Statement",
            "Create professional method statements for machinery moving, factory shifting, transport and lifting works.",
            METHOD_IMAGE,
            "WORK METHOD"
        )

        if st.button("Open Method Statement", key="open_ms"):
            go_to_page("📄 Method Statement")

    with col2:
        dashboard_card(
            "🏗️ Lifting Plan",
            "Generate lifting plan and permit-to-work documents based on load, crane, radius and site conditions.",
            LIFTING_IMAGE,
            "LIFTING OPERATION"
        )

        if st.button("Open Lifting Plan", key="open_lp"):
            go_to_page("🏗️ Lifting Plan")

    with col3:
        dashboard_card(
            "⚠️ Risk Assessment Pro",
            "Create structured 5x5 risk assessments using actual activity, hazard, controls and residual risk.",
            RISK_IMAGE,
            "SAFETY CONTROL"
        )

        if st.button("Open Risk Assessment", key="open_ra"):
            go_to_page("⚠️ Risk Assessment Pro")

    pro_col1, pro_col2, pro_col3 = st.columns(3)

    with pro_col1:
        dashboard_card(
            "📑 Method Statement PRO",
            "Develop the professional PowerPoint MOS using your existing EWMT master file while keeping the current Word MOS untouched.",
            METHOD_IMAGE,
            "POWERPOINT MASTER"
        )

        if st.button("Open Method Statement PRO", key="open_ms_pro"):
            go_to_page("📑 Method Statement PRO")

    st.markdown('<div class="section-title">Certificate / Records Modules</div>', unsafe_allow_html=True)

    col4, col5, col6 = st.columns(3)

    with col4:
        dashboard_card(
            "🧰 Lifting Gear Register",
            "Manage shackles, slings, wire ropes, lifting certificates, SWL records and expiry dates.",
            GEAR_IMAGE,
            "GEAR RECORDS"
        )

        if st.button("Open Lifting Gear Register", key="open_lg"):
            go_to_page("🧰 Lifting Gear Register")

    with col5:
        dashboard_card(
            "👷 Worker Training Certificate",
            "Search, preview and download worker training certificates uploaded into your GitHub folders.",
            TRAINING_IMAGE,
            "WORKER RECORDS"
        )

        if st.button("Open Worker Training Certificate", key="open_worker_cert"):
            go_to_page("👷 Worker Training Certificate")

    with col6:
        dashboard_card(
            "⏰ Expiry Alerts",
            "Check expired and expiring lifting gear certificates using expiry dates in your file names.",
            EXPIRY_IMAGE,
            "EXPIRY MONITORING"
        )

        if st.button("Open Expiry Alerts", key="open_expiry"):
            go_to_page("⏰ Expiry Alerts")

    st.markdown('<div class="section-title">System</div>', unsafe_allow_html=True)

    col7, col8, col9 = st.columns(3)

    with col7:
        dashboard_card(
            "⚙️ Settings",
            "Manage template placeholders, prepared-by names and future default company details.",
            HEADER_IMAGE,
            "SYSTEM CONFIG"
        )

        if st.button("Open Settings", key="open_settings"):
            go_to_page("⚙️ Settings")

    st.markdown("""
    <div class="footer-note">
        <b>EWMT Internal System</b><br>
        Dashboard counts are now calculated from your GitHub folders and lifting gear expiry filenames.
    </div>
    """, unsafe_allow_html=True)
    
# ======================================================
# METHOD STATEMENT
# ======================================================
if page == "📄 Method Statement":
    st.markdown("## 📄 Method Statement")
    st.caption("Fill in the work details and generate a Word method statement.")

    with st.expander("Project Details", expanded=True):
        ms_company = st.text_input("Company", "Eric Wong Machinery Transportation Pte Ltd", key="ms_company")
        ms_project_name = st.text_input("Project Name", key="ms_project_name")
        ms_date_input = st.date_input("Date", value=date.today(), format="DD/MM/YYYY", key="ms_date_input")
        ms_description = st.text_area("Description of Work", key="ms_description")
        ms_machine = st.text_input("Machine Model, Dimension and Weight", key="ms_machine")
        ms_operation_time = st.text_input("Operation Date & Time", key="ms_operation_time")
        ms_location = st.text_input("Location of Operation", key="ms_location")

    with st.expander("Standard Site Information", expanded=True):
        ms_obstacles = st.text_area(
            "Obstacles",
            value="Clear obstruction in way of working area and route to machine position.\nBarricade operation area to prevent persons who are not involved from entering unintentionally.",
            key="ms_obstacles"
        )

        ms_environment = st.text_area(
            "Environment",
            value="No operation will be carried out during heavy rain, thunderstorms and lightning weather.\nAll debris will be cleared and disposed.",
            key="ms_environment"
        )

        ms_lifting_crew = st.text_area(
            "Lifting Crew",
            value="MOM certified lifting supervisor, rigger, signalman and lorry loader operator will be involved in this operation.",
            key="ms_lifting_crew"
        )

        ms_prepared_by = st.text_input("Prepared By", value="Kevin Wong / Zailani", key="ms_prepared_by")

    generate_ms = st.button("📄 Generate Method Statement", key="generate_ms")

    if generate_ms:
        try:
            with st.spinner("Generating Method Statement..."):
                prompt = f"""
Create a professional Method Statement for machinery moving and lifting work in Singapore.

Company: {ms_company}
Project: {ms_project_name}
Location: {ms_location}
Description: {ms_description}
Machine: {ms_machine}
Obstacles / Site Access: {ms_obstacles}
Environment: {ms_environment}
Lifting Crew: {ms_lifting_crew}

Rules:
- Use formal contractor wording.
- Return plain text content for each field.
- Do not return dictionary-looking text.
- job_scope must be numbered steps.
- Do not include explanation, justification, summary, conclusion, or reference to company format.
- Do not write sentences starting with "The above", "These safety controls", "This sequence", or "This method statement".
- Do not mention that the content is consistent with previous company documents.
- Return only the actual content to be inserted into the Word document.
- equipment must only list equipment and materials.
- safety_aspect must only list safety precautions.
- job_scope must only list work steps.

Return these fields:
equipment
safety_aspect
job_scope
"""

                response = client.responses.create(
                    model="gpt-5.4",
                    input=prompt,
                    tools=[{
                        "type": "file_search",
                        "vector_store_ids": [MS_VECTOR_STORE_ID]
                    }],
                    text={
                        "format": {
                            "type": "json_schema",
                            "name": "ms_schema",
                            "schema": {
                                "type": "object",
                                "additionalProperties": False,
                                "properties": {
                                    "equipment": {"type": "string"},
                                    "safety_aspect": {"type": "string"},
                                    "job_scope": {"type": "string"}
                                },
                                "required": ["equipment", "safety_aspect", "job_scope"]
                            }
                        }
                    }
                )

                data = json.loads(response.output_text)

                data["equipment"] = clean_ms_text(data["equipment"])
                data["safety_aspect"] = clean_ms_text(data["safety_aspect"])
                data["job_scope"] = clean_ms_text(data["job_scope"])

                doc = Document(MS_TEMPLATE)

                replace_all(doc, {
                    "{{company}}": safe_text(ms_company),
                    "{{project_name}}": safe_text(ms_project_name),
                    "{{date}}": format_date_ddmmyyyy(ms_date_input),
                    "{{location}}": safe_text(ms_location),
                    "{{description_of_work}}": safe_text(ms_description),
                    "{{machine_spec}}": safe_text(ms_machine),
                    "{{equipment}}": safe_text(data["equipment"]),
                    "{{safety_aspect}}": safe_text(data["safety_aspect"]),
                    "{{job_scope}}": safe_text(data["job_scope"]),
                    "{{risk_assessment_note}}": "Refer as attached",
                    "{{operation_date}}": format_date_ddmmyyyy(ms_date_input),
                    "{{operation_time}}": safe_text(ms_operation_time),
                    "{{obstacles}}": safe_text(ms_obstacles),
                    "{{environment}}": safe_text(ms_environment),
                    "{{lifting_crew}}": safe_text(ms_lifting_crew),
                    "{{prepared_by}}": safe_text(ms_prepared_by),
                })

                buffer = BytesIO()
                doc.save(buffer)
                buffer.seek(0)

                st.download_button(
                    "Download Method Statement",
                    buffer,
                    "Method_Statement.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

        except Exception as e:
            st.error("Method Statement generation failed")
            st.exception(e)



# ======================================================
# METHOD STATEMENT PRO - POWERPOINT MASTER
# ======================================================
if page == "📑 Method Statement PRO":
    st.markdown("## 📑 Method Statement PRO")
    st.caption("Experimental professional MOS built from your existing EWMT PowerPoint master. The normal Word Method Statement remains untouched and continues to work separately.")

    master_path = mos_pro_find_master_template()

    if not PPTX_AVAILABLE:
        st.error("MOS PRO requires `python-pptx`. Add `python-pptx` to requirements.txt and redeploy the app.")
    elif not master_path:
        st.error("MOS PRO master PowerPoint not found.")
        st.code("Templates/MOS New.pptx")
        st.info("Upload your supplied MOS PowerPoint into the Templates folder. The code also accepts `MOS New(1).pptx`.")
    else:
        try:
            preview_prs = Presentation(master_path)
            st.success(f"Master loaded: {os.path.basename(master_path)} • {len(preview_prs.slides)} pages")
        except Exception as exc:
            st.error(f"Master template could not be opened: {exc}")

    st.info(
        "SAFE DEVELOPMENT MODE: The PowerPoint master is copied first. Standard MOS wording, flowcharts, pictures, lifting pages and emergency procedure remain in place. "
        "V1 only updates the repeated project-information block. The optional Work Method replacement is OFF by default."
    )

    with st.expander("1. Project Information — updates throughout the PowerPoint", expanded=True):
        c1, c2 = st.columns(2)
        with c1:
            pro_customer = st.text_input("Customer / Tenant Company", key="mospro_customer")
            pro_location = st.text_input("Site Location", key="mospro_location")
            pro_process = st.text_input(
                "Process",
                value="Lifting and Moving of Machinery",
                key="mospro_process"
            )
            pro_prepared = st.text_input("Prepared By", value="Kevin Wong", key="mospro_prepared")
        with c2:
            pro_approved = st.text_input("Approved By", value="Eric Wong (Director)", key="mospro_approved")
            pro_last_review = st.date_input("Last Review Date", value=date.today(), format="DD/MM/YYYY", key="mospro_last_review")
            pro_next_review = st.date_input("Next Review Date", value=date.today() + timedelta(days=365), format="DD/MM/YYYY", key="mospro_next_review")

        st.caption(
            "These values replace the old customer/site/process/review information in the repeated header table on the master pages. "
            "The MOS standard wording itself is not rewritten."
        )

    with st.expander("2. Optional Project-Specific Work Method — EXPERIMENTAL", expanded=False):
        replace_pro_work_method = st.checkbox(
            "Replace the existing Work Method Statement page with AI-generated project steps",
            value=False,
            key="mospro_replace_work_method"
        )

        st.warning(
            "Leave this OFF while you are first testing the master-template system. "
            "When ON, only the existing 'Work Method Statement for Lifting Operation' textbox is replaced."
        )

        pro_operation_type = st.selectbox(
            "Operation Type",
            [
                "Ground Floor - Unload / Shift / Position",
                "Ground Floor - Crane / Lorry Crane Hoisting",
                "Upper Floor / Roof Hoisting",
                "Indoor Machinery Shifting",
                "Loading / Unloading Only",
                "Custom / Other",
            ],
            key="mospro_operation_type"
        )
        pro_description = st.text_area("Description of Work", height=100, key="mospro_description")
        pro_machine = st.text_input("Machine / Load — model, dimensions and weight", key="mospro_machine")
        pro_equipment = st.text_area(
            "Equipment Actually Intended for this Job",
            height=90,
            placeholder="Example: 50T mobile crane, 10T forklift, hydraulic jacks, machine skates, 4 x 5T webbing slings",
            key="mospro_equipment"
        )
        pro_site_notes = st.text_area(
            "Special Site / Access Notes",
            height=90,
            placeholder="Only enter job-specific information. Standard EWMT safety wording already exists in the master MOS.",
            key="mospro_site_notes"
        )

    with st.expander("3. Master File Protection / Current V1 Scope", expanded=False):
        st.markdown(
            """
**What V1 changes**
- Customer / Tenant Company
- Site Location
- Process
- Prepared By
- Approved By
- Last Review Date
- Next Review Date
- Optional: existing project Work Method page only

**What V1 does NOT change**
- EWMT company header / logo
- Workplace Safety & Health standard pages
- Objectives / Scope / Responsibilities
- Routine / Non-Routine lift pages and flowcharts
- Competent Person / Supervisor / Operator / Rigger / Signalman standards
- Crane selection / ground / outriggers / pre-start standard wording
- Existing drawings, certificates, lifting-plan pages and pictures
- Emergency procedure page

This keeps the PRO version safe to develop slowly while the current Word Method Statement continues operating independently.
            """
        )

    generate_mos_pro = st.button(
        "📑 Generate MOS PRO PowerPoint",
        key="generate_mos_pro",
        type="primary",
        use_container_width=True
    )

    if generate_mos_pro:
        required = {
            "Customer / Tenant Company": pro_customer,
            "Site Location": pro_location,
            "Process": pro_process,
            "Prepared By": pro_prepared,
            "Approved By": pro_approved,
        }
        missing = [label for label, value in required.items() if not str(value).strip()]

        if missing:
            st.error("Please complete: " + ", ".join(missing))
        elif replace_pro_work_method and not pro_description.strip():
            st.error("Please enter Description of Work before replacing the Work Method page.")
        else:
            try:
                with st.spinner("Copying EWMT master PowerPoint and applying project information..."):
                    project_data = {
                        "customer": pro_customer,
                        "location": pro_location,
                        "process": pro_process,
                        "prepared_by": pro_prepared,
                        "approved_by": pro_approved,
                        "last_review": format_date_ddmmyyyy(pro_last_review),
                        "next_review": format_date_ddmmyyyy(pro_next_review),
                    }
                    work_method_data = {
                        "operation_type": pro_operation_type,
                        "description": pro_description,
                        "machine": pro_machine,
                        "equipment": pro_equipment,
                        "site_notes": pro_site_notes,
                    }

                    pro_buffer, pro_info = mos_pro_build_powerpoint(
                        project_data,
                        replace_work_method=replace_pro_work_method,
                        work_method_data=work_method_data,
                    )

                st.success(
                    f"MOS PRO generated from `{pro_info['master']}`. "
                    f"{pro_info['slides']} pages preserved; {pro_info['changed_fields']} project-information text blocks updated."
                )

                if pro_info.get("work_method_replaced"):
                    st.info(f"Experimental Work Method replacement applied on page {pro_info.get('work_method_slide')}.")
                else:
                    st.info("Standard master contents were preserved. Work Method replacement was not enabled.")

                safe_customer = re.sub(r"[^A-Za-z0-9_-]+", "_", pro_customer.strip())[:45] or "Project"
                st.download_button(
                    "⬇️ Download MOS PRO (.pptx)",
                    pro_buffer,
                    file_name=f"EWMT_MOS_PRO_{safe_customer}.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True,
                )

            except Exception as exc:
                st.error("MOS PRO generation failed")
                st.exception(exc)



# ======================================================
# LIFTING PLAN
# ======================================================
if page == "🏗️ Lifting Plan":
    st.markdown("## 🏗️ Lifting Plan / Permit to Work")
    st.caption("Fill in all lifting details. Checkboxes selected here will be inserted into your Word template.")

    with st.expander("1. General", expanded=True):
        lp_company = st.text_input("Company", "Eric Wong Machinery Transportation Pte Ltd", key="lp_company")
        lp_project_name = st.text_input("Project", key="lp_project_name")
        lp_location = st.text_input("Location of Lifting Operation", key="lp_location")
        lp_date_input = st.date_input("Date", value=date.today(), format="DD/MM/YYYY", key="lp_date_input")
        lp_operation_time = st.text_input("Operation Time", key="lp_operation_time")
        lp_validity = st.text_input("Validity Period of Lifting Operation", "1 Day", key="lp_validity")

    with st.expander("2. Details of Loads to be Hoist", expanded=True):
        lp_description = st.text_area("Description of Load", key="lp_description")
        lp_machine = st.text_input("Machine Name / Spec", key="lp_machine")
        lp_machine_dimension = st.text_input("Overall Dimension of Load", key="lp_machine_dimension")
        lp_machine_weight = st.text_input("Weight of Load", key="lp_machine_weight")

        weight_known = st.checkbox("Known Weight", value=True, key="weight_known")
        weight_estimated = st.checkbox("Estimated Weight", value=False, key="weight_estimated")

        cg_obvious = st.checkbox("Centre of Gravity - Obvious", value=True, key="cg_obvious")
        cg_estimated = st.checkbox("Centre of Gravity - Estimated", value=False, key="cg_estimated")
        cg_drawing = st.checkbox("Centre of Gravity - Determined by Drawing", value=False, key="cg_drawing")

    with st.expander("3. Details of Lifting Equipment", expanded=True):
        mobile_crane = st.checkbox("Mobile Crane", value=False, key="mobile_crane")
        lorry_loader = st.checkbox("Lorry Loader", value=True, key="lorry_loader")

        crane_name = st.text_input("LM / LE Registration No.", key="crane_name")
        crane_renew = st.text_input("Date of Last Certification (DD/MM/YYYY)", placeholder="DD/MM/YYYY", key="crane_renew")
        crane_expiry = st.text_input("Expiry Date of Certificate (DD/MM/YYYY)", placeholder="DD/MM/YYYY", key="crane_expiry")
        crane_swl = st.text_input("Max Safe Working Load", key="crane_swl")
        boom_length = st.text_input("Max Boom / Jib Length", key="boom_length")
        crane_radius = st.text_input("Intended Load Radius", key="crane_radius")
        crane_swl_radius = st.text_input("SWL at This Radius", key="crane_swl_radius")

        lifting_gear_manual = st.text_area(
            "Type of Lifting Gears",
            height=120,
            value="Wire rope slings / webbing slings, Shackles, Timber mats / steel plates, Tag lines",
            key="lifting_gear_manual"
        )

        lg_weight = st.text_input("Combined Weight of Lifting Gears", key="lg_weight")
        total_swl_lg = st.text_input("SWL of Lifting Gear", key="total_swl_lg")

        lg_cert_yes = st.checkbox("Certification of Lifting Gears - Yes", value=True, key="lg_cert_yes")
        lg_cert_no = st.checkbox("Certification of Lifting Gears - No", value=False, key="lg_cert_no")

        lg_expiry = st.text_input("Expiry Date of Lifting Gear Certificate (DD/MM/YYYY)", placeholder="DD/MM/YYYY", key="lg_expiry")

    with st.expander("4. Means of Communication", expanded=True):
        operator_can_see_yes = st.checkbox("Operator Can See Loading / Unloading Position - Yes", value=True, key="operator_can_see_yes")
        operator_can_see_no = st.checkbox("Operator Can See Loading / Unloading Position - No", value=False, key="operator_can_see_no")

        comm_standard = st.checkbox("Standard Hand Signals", value=True, key="comm_standard")
        comm_radio = st.checkbox("Radio", value=False, key="comm_radio")
        comm_others = st.checkbox("Others", value=False, key="comm_others")
        comm_others_text = st.text_input("Others Communication Details", key="comm_others_text")

    with st.expander("5. Personnel Involved in Lifting Operation", expanded=True):
        site_supervisor = st.text_input("Site Supervisor", "Ibrahim / Zahari / Zaharin / Wong Yen Siong", key="site_supervisor")
        lifting_supervisor = st.text_input("Lifting Supervisor", "Ibrahim / Zahari / Zaharin / Wong Yen Siong", key="lifting_supervisor")
        equipment_operator = st.text_input("Lifting Equipment Operator", "Lim Poh Soon / Norhalim / Lim Poh Thian / Ngaimin / Azmi", key="equipment_operator")
        rigger_1 = st.text_input("Rigger / Signalman 1", "Rizal / Hanifah / Aziz / Jamari / Ahmad", key="rigger_1")
        rigger_2 = st.text_input("Rigger / Signalman 2", "Rahman / Malik / Sarawanan / Sing Kwok Liang", key="rigger_2")

    with st.expander("6. Physical and Environmental Considerations", expanded=True):
        ground_safe_yes = st.checkbox("Ground Made Safe - Yes", value=True, key="ground_safe_yes")
        ground_safe_no = st.checkbox("Ground Made Safe - No", value=False, key="ground_safe_no")

        outriggers_yes = st.checkbox("Outriggers Evenly Extended - Yes", value=True, key="outriggers_yes")
        outriggers_no = st.checkbox("Outriggers Evenly Extended - No", value=False, key="outriggers_no")

        overhead_obstacle_yes = st.checkbox("Overhead Obstacles - Yes", value=False, key="overhead_obstacle_yes")
        overhead_obstacle_no = st.checkbox("Overhead Obstacles - No", value=True, key="overhead_obstacle_no")

        obstruction_yes = st.checkbox("Structure / Equipment / Materials Obstruction - Yes", value=False, key="obstruction_yes")
        obstruction_no = st.checkbox("Structure / Equipment / Materials Obstruction - No", value=True, key="obstruction_no")

        lighting_yes = st.checkbox("Lighting Adequate - Yes", value=True, key="lighting_yes")
        lighting_no = st.checkbox("Lighting Adequate - No", value=False, key="lighting_no")

        barricade_yes = st.checkbox("Zone Barricaded / Demarcated - Yes", value=True, key="barricade_yes")
        barricade_no = st.checkbox("Zone Barricaded / Demarcated - No", value=False, key="barricade_no")

        other_precautions = st.text_area("Other Precautions", key="other_precautions")

    with st.expander("7. Tasks", expanded=True):
        task_sequence = st.text_area(
            "Sequence of Lifting Operations",
            height=280,
            value="""1. Deploy lorry loader at designated unloading area
2. Set up crane with outriggers fully extended and resting on timber mats as base plate
3. Rigger to insert sling to crane hook
4. Secure sling to rigging point of load
5. Lorry loader to hoist down load from lorry chassis to the ground
6. Using forklift to unload and fork down machine from lorry chassis to the ground
7. Transport machine to door entrance
8. Using pallet truck to shift and position machine into factory premise
9. Position at the designated location
10. Once job complete, carry out proper housekeeping
11. All debris will be cleared and disposed
12. Job complete""",
            key="task_sequence"
        )

        person_in_charge = st.text_input(
            "Person in Charge for Each Step",
            "Zahari / Ibrahim / Wong Yen Siong",
            key="person_in_charge"
        )

    with st.expander("8. Approval of Lifting Plan", expanded=True):
        applied_by = st.text_input("Applied By", "Zailani", key="applied_by")
        applied_designation = st.text_input("Applied By Designation", "Supervisor", key="applied_designation")
        prepared_by = st.text_input("Prepared By", "Zahari", key="prepared_by_lp")
        prepared_designation = st.text_input("Prepared By Designation", "Lifting Supervisor", key="prepared_designation")
        assessed_by = st.text_input("Assessed By", "Kevin Wong", key="assessed_by")
        assessed_designation = st.text_input("Assessed By Designation", "Project Manager", key="assessed_designation")
        approved_by = st.text_input("Approved By", "Eric Wong", key="approved_by")
        approved_designation = st.text_input("Approved By Designation", "Managing Director", key="approved_designation")

    generate_lp = st.button("🏗️ Generate Lifting Plan", key="generate_lp")

    if generate_lp:
        try:
            with st.spinner("Generating Lifting Plan..."):

                response = client.responses.create(
                    model="gpt-5.4",
                    input=f"""
Improve this lifting task sequence into formal lifting plan wording.

Task sequence:
{task_sequence}

Return JSON only:
{{
 "lifting_method": "",
 "safety_controls": ""
}}
""",
                    tools=[{
                        "type": "file_search",
                        "vector_store_ids": [LP_VECTOR_STORE_ID]
                    }],
                    text={
                        "format": {
                            "type": "json_schema",
                            "name": "lp_schema",
                            "schema": {
                                "type": "object",
                                "additionalProperties": False,
                                "properties": {
                                    "lifting_method": {"type": "string"},
                                    "safety_controls": {"type": "string"}
                                },
                                "required": ["lifting_method", "safety_controls"]
                            }
                        }
                    }
                )

                data = json.loads(response.output_text)
                doc = Document(LP_TEMPLATE)

                replacements = {
                    "{{company}}": safe_text(lp_company),
                    "{{project_name}}": safe_text(lp_project_name),
                    "{{location}}": safe_text(lp_location),
                    "{{date}}": format_date_ddmmyyyy(lp_date_input),
                    "{{operation_date}}": format_date_ddmmyyyy(lp_date_input),
                    "{{operation_time}}": safe_text(lp_operation_time),
                    "{{validity_period}}": safe_text(lp_validity),

                    "{{description_of_work}}": safe_text(lp_description),
                    "{{machine_spec}}": safe_text(lp_machine),
                    "{{machine_name}}": safe_text(lp_machine),
                    "{{machine_dimension}}": safe_text(lp_machine_dimension),
                    "{{machine_weight}}": safe_text(lp_machine_weight),

                    "{{kw}}": tick(weight_known),
                    "{{ew}}": tick(weight_estimated),
                    "{{obv}}": tick(cg_obvious),
                    "{{Est}}": tick(cg_estimated),
                    "{{est}}": tick(cg_estimated),
                    "{{ddw}}": tick(cg_drawing),

                    "{{mob_cr}}": tick(mobile_crane),
                    "{{lor_cr}}": tick(lorry_loader),

                    "{{Crane_lm}}": safe_text(crane_name),
                    "{{crane_lm}}": safe_text(crane_name),
                    "{{crane_name}}": safe_text(crane_name),
                    "{{crane_renew}}": format_date_ddmmyyyy(crane_renew),
                    "{{crane_expiry}}": format_date_ddmmyyyy(crane_expiry),
                    "{{crane_swl}}": safe_text(crane_swl),
                    "{{boom_length}}": safe_text(boom_length),

                    "{{crane_radius}}": safe_text(crane_radius),
                    "{{ crane_radius }}": safe_text(crane_radius),
                    "{{crane_radius }}": safe_text(crane_radius),
                    "{{ crane_radius}}": safe_text(crane_radius),
                    "{{crane_swl_radius}}": safe_text(crane_swl_radius),

                    "{{lifting_gear}}": safe_text(lifting_gear_manual),
                    "{{lg_weight}}": safe_text(lg_weight),
                    "{{lifting_gear_wt}}": safe_text(lg_weight),
                    "{{total_swl_lg}}": safe_text(total_swl_lg),

                    "{{c_lg_y}}": tick(lg_cert_yes),
                    "{{c_lg_n}}": tick(lg_cert_no),
                    "{{lg_expiry}}": format_date_ddmmyyyy(lg_expiry),

                    "{{coms_y}}": tick(operator_can_see_yes),
                    "{{coms_n}}": tick(operator_can_see_no),
                    "{{coms}}": tick(operator_can_see_yes),

                    "{{shs}}": tick(comm_standard),
                    "{{rad}}": tick(comm_radio),
                    "{{comm_standard}}": tick(comm_standard),
                    "{{comm_radio}}": tick(comm_radio),
                    "{{comm_others}}": tick(comm_others),
                    "{{comm_others_text}}": safe_text(comm_others_text),

                    "{{site_supervisor}}": safe_text(site_supervisor),
                    "{{lifting_supervisor}}": safe_text(lifting_supervisor),
                    "{{equipment_operator}}": safe_text(equipment_operator),
                    "{{rigger_1}}": safe_text(rigger_1),
                    "{{rigger_2}}": safe_text(rigger_2),

                    "{{gc_y}}": tick(ground_safe_yes),
                    "{{gc_n}}": tick(ground_safe_no),
                    "{{go_y}}": tick(outriggers_yes),
                    "{{go_n}}": tick(outriggers_no),
                    "{{ob_y}}": tick(overhead_obstacle_yes),
                    "{{ob_n}}": tick(overhead_obstacle_no),
                    "{{st_y}}": tick(obstruction_yes),
                    "{{st_n}}": tick(obstruction_no),
                    "{{li_y}}": tick(lighting_yes),
                    "{{li_n}}": tick(lighting_no),
                    "{{de_y}}": tick(barricade_yes),
                    "{{de_n}}": tick(barricade_no),
                    "{{other_precautions}}": safe_text(other_precautions),

                    "{{task_sequence}}": safe_text(task_sequence),
                    "{{tasks}}": safe_text(task_sequence),
                    "{{lifting_method}}": safe_text(data.get("lifting_method", "")),
                    "{{safety_controls}}": safe_text(data.get("safety_controls", "")),
                    "{{person_in_charge}}": safe_text(person_in_charge),
                    "{{task_pic}}": safe_text(person_in_charge),

                    "{{applied_by}}": safe_text(applied_by),
                    "{{applied_designation}}": safe_text(applied_designation),
                    "{{prepared_by}}": safe_text(prepared_by),
                    "{{prepared_designation}}": safe_text(prepared_designation),
                    "{{assessed_by}}": safe_text(assessed_by),
                    "{{assessed_designation}}": safe_text(assessed_designation),
                    "{{approved_by}}": safe_text(approved_by),
                    "{{approved_designation}}": safe_text(approved_designation),

                    "{{known_weight_checked}}": tick(weight_known),
                    "{{estimated_weight_checked}}": tick(weight_estimated),
                    "{{center_gravity_obvious}}": tick(cg_obvious),
                    "{{center_gravity_estimated}}": tick(cg_estimated),
                    "{{center_gravity_drawing}}": tick(cg_drawing),
                    "{{mobile_crane_checked}}": tick(mobile_crane),
                    "{{lorry_loader_checked}}": tick(lorry_loader),
                    "{{lg_cert_yes}}": tick(lg_cert_yes),
                    "{{lg_cert_no}}": tick(lg_cert_no),
                    "{{operator_can_see_yes}}": tick(operator_can_see_yes),
                    "{{operator_can_see_no}}": tick(operator_can_see_no),
                    "{{ground_safe_yes}}": tick(ground_safe_yes),
                    "{{ground_safe_no}}": tick(ground_safe_no),
                    "{{outriggers_yes}}": tick(outriggers_yes),
                    "{{outriggers_no}}": tick(outriggers_no),
                    "{{obstacles_yes}}": tick(overhead_obstacle_yes),
                    "{{obstacles_no}}": tick(overhead_obstacle_no),
                    "{{obstruction_yes}}": tick(obstruction_yes),
                    "{{obstruction_no}}": tick(obstruction_no),
                    "{{lighting_yes}}": tick(lighting_yes),
                    "{{lighting_no}}": tick(lighting_no),
                    "{{barricade_yes}}": tick(barricade_yes),
                    "{{barricade_no}}": tick(barricade_no),
                }

                replace_all(doc, replacements)

                buffer = BytesIO()
                doc.save(buffer)
                buffer.seek(0)

                st.download_button(
                    "Download Lifting Plan",
                    buffer,
                    "Lifting_Plan.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

        except Exception as e:
            st.error("Lifting Plan generation failed")
            st.exception(e)


# ======================================================
# RISK ASSESSMENT PRO
# ======================================================
if page == "⚠️ Risk Assessment Pro":
    st.markdown("## ⚠️ Risk Assessment Pro")
    st.caption("Generate a professional 5x5 Risk Assessment based on actual work activities.")

    with st.expander("Project Details", expanded=True):
        ra_company = st.text_input("Company", "Eric Wong Machinery Transportation Pte Ltd", key="ra_company")
        ra_project_name = st.text_input("Project Name", key="ra_project_name")
        ra_location = st.text_input("Location", key="ra_location")
        ra_machine = st.text_input("Machine Spec", key="ra_machine")
        ra_description = st.text_area("Description of Work", key="ra_description")
        ra_date_input = st.date_input("Date", value=date.today(), format="DD/MM/YYYY", key="ra_date_input")
        ra_due_date_input = st.date_input("Due Date", value=date.today(), format="DD/MM/YYYY", key="ra_due_date_input")

    with st.expander("Risk Assessment Details", expanded=True):
        ra_process = st.text_input("RA Process", "Machinery Moving / Lifting Operation", key="ra_process")

        activities = st.text_area(
            "Work Activities (1 per line)",
            height=220,
            value="""Transport of lifting machinery into or out of site premises
Setting up of crane on site
Lifting operation
Signalling of load""",
            key="activities"
        )

    generate_ra_pro = st.button("⚠️ Generate Risk Assessment Pro", key="generate_ra_pro")

    if generate_ra_pro:
        try:
            with st.spinner("Generating Risk Assessment Pro..."):

                prompt = f"""
Create professional Singapore style 5x5 Risk Assessment.

Company: {ra_company}
Project: {ra_project_name}
Location: {ra_location}
Process: {ra_process}
Machine: {ra_machine}
Description: {ra_description}
Date: {format_date_ddmmyyyy(ra_date_input)}
Due Date: {format_date_ddmmyyyy(ra_due_date_input)}

Activities:
{activities}

Important:
- Every generated row must directly match the work activities provided.
- Do not invent unrelated activities.
- For each activity, create 1 to 3 relevant hazards.
- If one activity has more than one hazard, only the first row should contain ref and work_activity.
- Subsequent hazard rows for the same activity must use empty string for ref and work_activity.
- Use machinery moving / lifting / forklift / jacking / roller / crating style controls.
- Use wording style similar to Eric Wong Machinery Transportation Pte Ltd RA examples.
- Return JSON only.

Schema:
{{
 "rows":[
   {{
    "ref":"1",
    "work_activity":"",
    "hazard":"",
    "possible_injury":"",
    "existing_controls":"",
    "s":"4",
    "l":"2",
    "rpn":"8",
    "additional_controls":"",
    "rs":"4",
    "rl":"1",
    "rrpn":"4",
    "person":"Supervisor on site",
    "due_date":"{format_date_ddmmyyyy(ra_due_date_input)}",
    "remark":""
   }}
 ]
}}
"""

                response = client.responses.create(
                    model="gpt-5.4",
                    input=prompt,
                    tools=[{
                        "type": "file_search",
                        "vector_store_ids": [RA_VECTOR_STORE_ID]
                    }]
                )

                data = json.loads(response.output_text)

                doc = Document(RA_TEMPLATE)

                replace_all(doc, {
                    "{{company}}": ra_company,
                    "{{location}}": ra_location,
                    "{{process}}": ra_process,
                    "{{date}}": format_date_ddmmyyyy(ra_date_input),
                    "{{due_date}}": format_date_ddmmyyyy(ra_due_date_input)
                })

                fill_inventory_table(doc, activities, ra_location, ra_process)

                table = find_ra_table(doc)

                if table:
                    clear_rows_after_column_header(table)

                    for r in data["rows"]:
                        add_ra_row(table, [
                            r["ref"],
                            r["work_activity"],
                            r["hazard"],
                            r["possible_injury"],
                            r["existing_controls"],
                            r["s"],
                            r["l"],
                            r["rpn"],
                            r["additional_controls"],
                            r["rs"],
                            r["rl"],
                            r["rrpn"],
                            r["person"],
                            format_date_ddmmyyyy(ra_due_date_input),
                            r["remark"]
                        ])

                    merge_same_work_activity_cells(table)

                format_risk_assessment(doc)

                buffer = BytesIO()
                doc.save(buffer)
                buffer.seek(0)

                st.download_button(
                    "Download Risk Assessment Pro",
                    buffer,
                    "Risk_Assessment_Pro.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

        except Exception as e:
            st.error("Risk Assessment Pro generation failed")
            st.exception(e)


# ======================================================
# LIFTING GEAR REGISTER
# ======================================================
if page == "🧰 Lifting Gear Register":
    certificate_browser(
        folder_name="Lifting Gears Certificate",
        title="🧰 Lifting Gear Register",
        info_text="Certificates loaded from GitHub folder: Lifting Gears Certificate",
        search_label="Search by SWL / keyword",
        search_placeholder="Example: 3 Ton, 10 Ton, shackle, round sling",
        download_label="Download Selected Certificate"
    )


# ======================================================
# WORKER TRAINING CERTIFICATE
# ======================================================
if page == "👷 Worker Training Certificate":
    certificate_browser(
        folder_name="Workers Certificate",
        title="👷 Worker Training Certificate",
        info_text="Certificates loaded from GitHub folder: Workers Certificate",
        search_label="Search by worker name / course / keyword",
        search_placeholder="Example: Ibrahim, forklift, boom lift, lifting supervisor, rigger",
        download_label="Download Selected Worker Certificate"
    )


# ======================================================
# EXPIRY ALERTS
# ======================================================
if page == "⏰ Expiry Alerts":
    st.markdown("## ⏰ Expiry Alerts")
    st.caption("Show expired and expiring lifting gear certificates from your GitHub folder.")

    import re

    CERT_FOLDER = os.path.join(BASE_DIR, "Lifting Gears Certificate")

    if not os.path.exists(CERT_FOLDER):
        st.error("Folder not found: Lifting Gears Certificate")
        st.code("Lifting Gears Certificate")
    else:
        files = [
            f for f in os.listdir(CERT_FOLDER)
            if f.lower().endswith((".pdf", ".png", ".jpg", ".jpeg"))
        ]

        if not files:
            st.warning("No certificate files found.")
        else:
            today = date.today()
            alert_days = st.number_input(
                "Show certificates expiring within how many days?",
                min_value=1,
                max_value=365,
                value=30
            )

            records = []

            for f in files:
                found_date = None

                patterns = [
                    r"(\d{4})[-_\\.](\d{1,2})[-_\\.](\d{1,2})",
                    r"(\d{1,2})[-_\\.](\d{1,2})[-_\\.](\d{4})",
                ]

                for pattern in patterns:
                    match = re.search(pattern, f)
                    if match:
                        try:
                            parts = match.groups()

                            if len(parts[0]) == 4:
                                found_date = date(int(parts[0]), int(parts[1]), int(parts[2]))
                            else:
                                found_date = date(int(parts[2]), int(parts[1]), int(parts[0]))

                            break
                        except Exception:
                            found_date = None

                if found_date:
                    days_left = (found_date - today).days

                    if days_left < 0:
                        status = "Expired"
                    elif days_left <= alert_days:
                        status = "Expiring Soon"
                    else:
                        status = "Valid"

                    records.append({
                        "Certificate File": f,
                        "Expiry Date": format_date_ddmmyyyy(found_date),
                        "Days Left": days_left,
                        "Status": status
                    })
                else:
                    records.append({
                        "Certificate File": f,
                        "Expiry Date": "No date found in filename",
                        "Days Left": "",
                        "Status": "Unknown"
                    })

            expired = [r for r in records if r["Status"] == "Expired"]
            expiring = [r for r in records if r["Status"] == "Expiring Soon"]
            valid = [r for r in records if r["Status"] == "Valid"]

            c1, c2, c3 = st.columns(3)
            c1.metric("Expired", len(expired))
            c2.metric("Expiring Soon", len(expiring))
            c3.metric("Valid", len(valid))

            st.markdown("### Certificate Expiry List")
            st.dataframe(records, use_container_width=True)

            st.info("For this to work, put expiry date inside the certificate filename, example: 3 Ton Shackle Expiry 30-06-2026.pdf")


# ======================================================
# SETTINGS / ADMIN DOCUMENT MANAGER
# ======================================================
if page == "⚙️ Settings":
    st.markdown("## ⚙️ Settings / Administrator")
    st.caption("Secure document control and template reference for the EWMT internal system.")

    if not admin_is_logged_in():
        admin_login_form()
    else:
        render_admin_document_manager()

    st.markdown("---")
    with st.expander("📑 Method Statement PRO Master Template"):
        st.code("Templates/MOS New.pptx")
        st.write("MOS PRO uses the PowerPoint itself as the master. No placeholders are required in V1; the app finds and updates the repeated Customer / Site / Process / Prepared / Approved / Review fields while preserving the rest of the presentation.")
        st.caption("Accepted alternate filenames: MOS New(1).pptx or Method Statement PRO.pptx")

    with st.expander("📄 Method Statement Placeholder Guide"):
        st.code("""
Use these placeholders in your Method Statement Word template:

{{date}}
{{description_of_work}}
{{machine_spec}}
{{operation_date}}
{{operation_time}}
{{location}}
{{equipment}}
{{obstacles}}
{{environment}}
{{lifting_crew}}
{{safety_aspect}}
{{job_scope}}
{{prepared_by}}
""")

    with st.expander("🏗️ Lifting Plan Placeholder Guide"):
        st.code("""
Use these placeholders in your Lifting Plan Word template:

General:
{{project_name}}
{{location}}
{{operation_date}}
{{operation_time}}
{{validity_period}}

Load:
{{machine_name}}
{{machine_dimension}}
{{machine_weight}}
{{kw}} Known weight
{{ew}} Estimated weight
{{obv}} Obvious
{{Est}} Estimated
{{ddw}} Determined by drawing

Lifting Equipment:
{{mob_cr}} Mobile crane
{{lor_cr}} Lorry loader
{{Crane_lm}}
{{crane_lm}}
{{crane_renew}}
{{crane_expiry}}
{{crane_swl}}
{{crane_radius}}
{{crane_swl_radius}}
{{lifting_gear}}
{{lifting_gear_wt}}
{{total_swl_lg}}
{{c_lg_y}} Yes
{{c_lg_n}} No
{{lg_expiry}}

Communication:
{{coms_y}} Yes
{{coms_n}} No
{{shs}} Standard hand signals
{{rad}} Radio

Physical / Environmental:
{{gc_y}} Yes    {{gc_n}} No
{{go_y}} Yes    {{go_n}} No
{{ob_y}} Yes    {{ob_n}} No
{{st_y}} Yes    {{st_n}} No
{{li_y}} Yes    {{li_n}} No
{{de_y}} Yes    {{de_n}} No

Approval:
{{applied_by}}
{{applied_designation}}
{{prepared_by}}
{{prepared_designation}}
{{assessed_by}}
{{assessed_designation}}
{{approved_by}}
{{approved_designation}}
""")
