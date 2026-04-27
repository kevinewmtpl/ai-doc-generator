import os
import json
import base64
from io import BytesIO
from datetime import date

import streamlit as st
import streamlit.components.v1 as components
from openai import OpenAI
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn

# =====================================================
# PAGE CONFIG
# =====================================================
st.set_page_config(
    page_title="EWMT AI Document System",
    page_icon="🏗️",
    layout="wide"
)

# =====================================================
# SESSION STATE
# =====================================================
PAGES = [
    "🏠 Dashboard",
    "📄 Method Statement",
    "🏗️ Lifting Plan",
    "⚠️ Risk Assessment Pro",
    "🧰 Lifting Gear Register",
    "⏰ Expiry Alerts",
    "⚙️ Settings"
]

if "page" not in st.session_state:
    st.session_state.page = "🏠 Dashboard"

# =====================================================
# STYLE
# =====================================================
st.markdown("""
<style>
.block-container {
    max-width: 1500px;
    padding-top: 1rem;
}
[data-testid="stSidebar"] {
    background: linear-gradient(180deg,#0f172a,#1e293b);
}
[data-testid="stSidebar"] * {
    color:white;
}
.ewmt-header{
    background:linear-gradient(90deg,#0f172a,#1e3a8a);
    padding:24px;
    border-radius:18px;
    color:white;
    margin-bottom:20px;
}
.ewmt-title{
    font-size:34px;
    font-weight:800;
}
.ewmt-sub{
    color:#dbeafe;
    font-size:16px;
}
.dashboard-card{
    background:white;
    border:1px solid #e5e7eb;
    border-radius:16px;
    padding:20px;
    min-height:145px;
    box-shadow:0 4px 14px rgba(0,0,0,0.08);
    margin-bottom:12px;
}
.stButton > button{
    width:100%;
    background:#1e3a8a;
    color:white;
    border:none;
    border-radius:10px;
    font-weight:700;
}
.stDownloadButton > button{
    background:#047857;
    color:white;
    border:none;
    border-radius:10px;
    font-weight:700;
}
</style>
""", unsafe_allow_html=True)

# =====================================================
# HEADER
# =====================================================
st.markdown("""
<div class="ewmt-header">
<div class="ewmt-title">Eric Wong Machinery Transportation Pte Ltd</div>
<div class="ewmt-sub">
Heavy Machinery Moving • Lifting • Transportation • AI Document Control System
</div>
</div>
""", unsafe_allow_html=True)

# =====================================================
# OPENAI
# =====================================================
client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

MS_VECTOR_STORE_ID = "vs_69ecc533a1208191a8595c674753e99e"
RA_VECTOR_STORE_ID = "vs_69ecd191648481919d1d1d57f21264af"
LP_VECTOR_STORE_ID = "vs_69ecdc44d59081919fb10574510b7454"

# =====================================================
# PATHS
# =====================================================
BASE_DIR = os.path.dirname(__file__)

MS_TEMPLATE = os.path.join(BASE_DIR, "Templates", "Method of statement Template.docx")
RA_TEMPLATE = os.path.join(BASE_DIR, "Templates", "RA Template.docx")
LP_TEMPLATE = os.path.join(BASE_DIR, "Templates", "Lifting Plan Template.docx")

# =====================================================
# FUNCTIONS
# =====================================================
def go_to_page(name):
    st.session_state.page = name
    st.rerun()

def replace_all(doc, data):
    for p in doc.paragraphs:
        for k, v in data.items():
            if k in p.text:
                p.text = p.text.replace(k, str(v))

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for k, v in data.items():
                        if k in p.text:
                            p.text = p.text.replace(k, str(v))

def set_font(doc, size=10):
    for p in doc.paragraphs:
        for run in p.runs:
            run.font.name = "Times New Roman"
            run._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
            run.font.size = Pt(size)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for run in p.runs:
                        run.font.name = "Times New Roman"
                        run._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
                        run.font.size = Pt(size)

# =====================================================
# SIDEBAR
# =====================================================
with st.sidebar:
    st.markdown("## EWMT System")
    st.markdown("AI Document Control")
    st.markdown("---")

    selected = st.radio(
        "Navigation",
        PAGES,
        index=PAGES.index(st.session_state.page)
    )

    st.session_state.page = selected
    page = selected

# =====================================================
# DASHBOARD
# =====================================================
if page == "🏠 Dashboard":

    st.markdown("## EWMT Dashboard")

    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown("""
        <div class="dashboard-card">
        <h3>📄 Method Statement</h3>
        Create professional method statements.
        </div>
        """, unsafe_allow_html=True)
        if st.button("Open Method Statement"):
            go_to_page("📄 Method Statement")

    with col2:
        st.markdown("""
        <div class="dashboard-card">
        <h3>🏗️ Lifting Plan</h3>
        Create lifting plan documents.
        </div>
        """, unsafe_allow_html=True)
        if st.button("Open Lifting Plan"):
            go_to_page("🏗️ Lifting Plan")

    with col3:
        st.markdown("""
        <div class="dashboard-card">
        <h3>⚠️ Risk Assessment</h3>
        Create 5x5 RA documents.
        </div>
        """, unsafe_allow_html=True)
        if st.button("Open Risk Assessment"):
            go_to_page("⚠️ Risk Assessment Pro")

# =====================================================
# METHOD STATEMENT
# =====================================================
if page == "📄 Method Statement":

    st.markdown("## 📄 Method Statement")

    company = st.text_input("Company", "Eric Wong Machinery Transportation Pte Ltd")
    project = st.text_input("Project Name")
    work_date = st.date_input("Date", value=date.today())
    desc = st.text_area("Description")
    machine = st.text_input("Machine")
    location = st.text_input("Location")

    if st.button("Generate Method Statement"):

        doc = Document(MS_TEMPLATE)

        replace_all(doc, {
            "{{company}}": company,
            "{{date}}": str(work_date),
            "{{description_of_work}}": desc,
            "{{machine_spec}}": machine,
            "{{location}}": location,
            "{{prepared_by}}": "Kevin Wong / Zailani"
        })

        set_font(doc, 12)

        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        st.download_button(
            "Download Method Statement",
            buffer,
            file_name="Method_Statement.docx"
        )

# =====================================================
# LIFTING PLAN
# =====================================================
if page == "🏗️ Lifting Plan":

    st.markdown("## 🏗️ Lifting Plan")

    company = st.text_input("Company", "Eric Wong Machinery Transportation Pte Ltd", key="lp1")
    project = st.text_input("Project", key="lp2")
    location = st.text_input("Location", key="lp3")
    load = st.text_input("Load Description")
    weight = st.text_input("Weight")
    crane = st.text_input("Crane Name")

    if st.button("Generate Lifting Plan"):

        doc = Document(LP_TEMPLATE)

        replace_all(doc, {
            "{{company}}": company,
            "{{project_name}}": project,
            "{{location}}": location,
            "{{description_of_work}}": load,
            "{{machine_weight}}": weight,
            "{{crane_name}}": crane,
            "{{prepared_by}}": "Kevin Wong / Zailani"
        })

        set_font(doc, 10)

        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        st.download_button(
            "Download Lifting Plan",
            buffer,
            file_name="Lifting_Plan.docx"
        )

# =====================================================
# RISK ASSESSMENT
# =====================================================
if page == "⚠️ Risk Assessment Pro":

    st.markdown("## ⚠️ Risk Assessment")

    company = st.text_input("Company", "Eric Wong Machinery Transportation Pte Ltd", key="ra1")
    location = st.text_input("Location", key="ra2")
    process = st.text_input("Process", "Machinery Moving")
    hazards = st.text_area("Activities / Hazards")

    if st.button("Generate Risk Assessment"):

        doc = Document(RA_TEMPLATE)

        replace_all(doc, {
            "{{company}}": company,
            "{{location}}": location,
            "{{process}}": process,
            "{{date}}": str(date.today())
        })

        set_font(doc, 10)

        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        st.download_button(
            "Download Risk Assessment",
            buffer,
            file_name="Risk_Assessment.docx"
        )

# =====================================================
# LIFTING GEAR REGISTER
# =====================================================
if page == "🧰 Lifting Gear Register":

    st.markdown("## 🧰 Lifting Gear Register")

    cert_folder = os.path.join(BASE_DIR, "Lifting Gears Certificate")

    if not os.path.exists(cert_folder):
        st.error("Folder not found")
    else:
        files = sorted([
            f for f in os.listdir(cert_folder)
            if f.lower().endswith((".pdf",".png",".jpg",".jpeg"))
        ])

        st.success(f"Found {len(files)} files")

        keyword = st.text_input(
            "Search by keyword",
            placeholder="3 Ton, shackle, sling"
        )

        filtered = files

        if keyword:
            words = keyword.lower().split()
            filtered = [
                f for f in files
                if all(w in f.lower() for w in words)
            ]

        if filtered:

            selected = st.selectbox("Choose File", filtered)

            file_path = os.path.join(cert_folder, selected)

            with open(file_path, "rb") as f:
                file_bytes = f.read()

            st.download_button(
                "Download Certificate",
                file_bytes,
                file_name=selected
            )

            if selected.lower().endswith((".png",".jpg",".jpeg")):
                st.image(file_path)

            if selected.lower().endswith(".pdf"):

                st.markdown("### PDF Preview")

                base64_pdf = base64.b64encode(file_bytes).decode("utf-8")

                pdf_html = f"""
                <iframe
                    src="data:application/pdf;base64,{base64_pdf}"
                    width="100%"
                    height="900px"
                    style="border:1px solid #ccc;border-radius:10px;">
                </iframe>
                """

                components.html(pdf_html, height=920)

# =====================================================
# EXPIRY
# =====================================================
if page == "⏰ Expiry Alerts":
    st.markdown("## ⏰ Expiry Alerts")
    st.info("Coming Soon")

# =====================================================
# SETTINGS
# =====================================================
if page == "⚙️ Settings":
    st.markdown("## ⚙️ Settings")
    st.info("Coming Soon")
