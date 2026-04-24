import os
import json
from io import BytesIO
from datetime import date

import streamlit as st
from openai import OpenAI
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn

# ======================================================
# PAGE CONFIG
# ======================================================
st.set_page_config(
    page_title="EWMT Document Generator",
    page_icon="🏗️",
    layout="wide"
)

# ======================================================
# STYLE
# ======================================================
st.markdown("""
<style>
.block-container{
    padding-top:1rem;
    max-width:1450px;
}

.main-title{
    background:linear-gradient(90deg,#0f172a,#1e3a8a);
    color:white;
    padding:22px;
    border-radius:16px;
    margin-bottom:20px;
}

.main-title h1{
    margin:0;
    font-size:34px;
}

.main-title p{
    margin:0;
    color:#dbeafe;
}

.stButton > button{
    background:#1e3a8a;
    color:white;
    font-weight:700;
    border-radius:10px;
}

.stButton > button:hover{
    background:#0f172a;
    color:white;
}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="main-title">
<h1>Eric Wong Machinery Transportation Pte Ltd</h1>
<p>Heavy Machinery Moving • Lifting • Transportation • AI Document Generator</p>
</div>
""", unsafe_allow_html=True)

# ======================================================
# OPENAI
# ======================================================
client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

# ======================================================
# VECTOR STORES
# ======================================================
MS_VECTOR_STORE_ID = "vs_69e3287971a48191b6b4da9f7a9679eb"
RA_VECTOR_STORE_ID = "vs_69e34a1321a88191a9d80166ae6316c7"
LP_VECTOR_STORE_ID = "vs_69e865a133888191b5c39d3e47f9d578"

# ======================================================
# FILE PATHS
# ======================================================
BASE_DIR = os.path.dirname(__file__)

MS_TEMPLATE = os.path.join(BASE_DIR, "Templates", "Method of statement Template.docx")
RA_TEMPLATE = os.path.join(BASE_DIR, "Templates", "RA Template.docx")
LP_TEMPLATE = os.path.join(BASE_DIR, "Templates", "Lifting Plan Template.docx")

# ======================================================
# COMMON FUNCTIONS
# ======================================================
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


# ======================================================
# FULL RA FONT FIX
# ======================================================
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
        if "Hazard Identification" in full:
            return table
    return None


def find_ra_column_header_row(table):
    for i, row in enumerate(table.rows):
        txt = [c.text.strip() for c in row.cells]
        if "Ref" in txt and "Work Activity" in txt:
            return i
    return None


def clear_rows_after_column_header(table):
    idx = find_ra_column_header_row(table)

    while len(table.rows) > idx + 1:
        row = table.rows[idx + 1]
        row._element.getparent().remove(row._element)


def add_ra_row(table, values):
    row = table.add_row().cells

    for i, v in enumerate(values):
        if i < len(row):
            set_ra_cell_text(row[i], v)


# ======================================================
# UI TABS
# ======================================================
tab1, tab2, tab3 = st.tabs([
    "Method Statement",
    "Lifting Plan",
    "Risk Assessment Pro"
])

# ======================================================
# TAB 3 ONLY (YOUR ISSUE)
# ======================================================
with tab3:

    st.header("Risk Assessment Pro")

    company = st.text_input("Company", "Eric Wong Machinery Transportation Pte Ltd")
    project = st.text_input("Project Name")
    location = st.text_input("Location")
    process = st.text_input("Process", "Machinery Moving / Lifting Operation")
    date_input = st.date_input("Date", value=date.today())

    activities = st.text_area(
        "Activities (1 per line)",
        height=200,
        value="""Transport lifting machinery into premises
Setting up crane
Lifting operation
Signalling load"""
    )

    generate = st.button("Generate Risk Assessment")

    if generate:

        prompt = f"""
Create Singapore 5x5 Risk Assessment.

Company:{company}
Project:{project}
Location:{location}
Process:{process}

Activities:
{activities}

Return JSON:
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
   "person":"Supervisor",
   "due_date":"{date_input}",
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
            "{{company}}": company,
            "{{location}}": location,
            "{{process}}": process,
            "{{date}}": str(date_input)
        })

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
                    r["due_date"],
                    r["remark"]
                ])

        # ==================================================
        # THIS FIXES ALL FONT ISSUE
        # ==================================================
        format_risk_assessment(doc)

        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        st.download_button(
            "Download Risk Assessment",
            buffer,
            "Risk_Assessment_Pro.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
