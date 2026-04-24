import os
import json
from io import BytesIO
from datetime import date

import streamlit as st
from openai import OpenAI
from docx import Document

client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

# =====================
# VECTOR STORES
# =====================
MS_VECTOR_STORE_ID = "vs_69e3287971a48191b6b4da9f7a9679eb"
RA_VECTOR_STORE_ID = "vs_69e34a1321a88191a9d80166ae6316c7"
LP_VECTOR_STORE_ID = "vs_69e865a133888191b5c39d3e47f9d578"

# =====================
# PATHS
# =====================
BASE_DIR = os.path.dirname(__file__)

MS_TEMPLATE = os.path.join(BASE_DIR, "Templates", "Method of statement Template.docx")
RA_TEMPLATE = os.path.join(BASE_DIR, "Templates", "RA Template.docx")
LP_TEMPLATE = os.path.join(BASE_DIR, "Templates", "Lifting Plan Template.docx")

# =====================
# UI
# =====================
st.title("Document Generator")

company = st.text_input("Company", "Eric Wong Machinery Transportation Pte Ltd")
project_name = st.text_input("Project Name")
location = st.text_input("Location")
description = st.text_area("Description of Work")
machine = st.text_input("Machine Spec")
date_input = st.date_input("Date", value=date.today())

st.subheader("Lifting Plan Details")

operation_time = st.text_input("Operation Time")
machine_dimension = st.text_input("Machine Dimension")
machine_weight = st.text_input("Machine Weight")

crane_name = st.text_input("Crane Name / Model")
crane_renew = st.text_input("Crane Cert Date")
crane_expiry = st.text_input("Crane Cert Expiry")

crane_swl = st.text_input("Crane SWL")
crane_radius = st.text_input("Crane Radius")
crane_swl_radius = st.text_input("SWL at Radius")

total_swl_lg = st.text_input("Total SWL of Lifting Gear")
lg_expiry = st.text_input("Lifting Gear Expiry")

st.subheader("Personnel")

site_supervisor = st.text_input("Site Supervisor")
lifting_supervisor = st.text_input("Lifting Supervisor")
equipment_operator = st.text_input("Equipment Operator")
rigger_1 = st.text_input("Rigger 1")
rigger_2 = st.text_input("Rigger 2")

st.subheader("Conditions")

ground_safe = st.checkbox("Ground Safe", value=True)
outriggers = st.checkbox("Outriggers Extended", value=True)
no_overhead_obstacles = st.checkbox("No Overhead Obstacles", value=True)
lighting = st.checkbox("Lighting Adequate", value=True)
barricade = st.checkbox("Area Barricaded", value=True)

task_sequence = st.text_area(
    "Lifting Steps",
    height=200,
    value="""1. Deploy crane / lorry loader at designated unloading area
2. Set up outriggers fully extended and rest on timber mats / steel plates
3. Carry out rigging and hook-on
4. Conduct trial lift
5. Hoist load slowly and steadily
6. Shift load to designated position
7. Lower load in a controlled manner
8. Remove lifting gear and carry out housekeeping"""
)

col1, col2, col3 = st.columns(3)
generate_ms = col1.button("Generate Method Statement")
generate_ra = col2.button("Generate Risk Assessment")
generate_lp = col3.button("Generate Lifting Plan")

# =====================
# REPLACE FUNCTION
# =====================
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

# =====================
# KEEP YOUR EXISTING METHOD STATEMENT / RA / LP CODE HERE
# (no changes)
# =====================


# ======================================================
# ADD THIS NEW SECTION ONLY (Risk Assessment Pro)
# ======================================================

st.markdown("---")
st.subheader("Risk Assessment Pro")

ra_process = st.text_input(
    "RA Process",
    "Machinery Moving / Lifting Operation"
)

activities = st.text_area(
    "Work Activities (1 per line)",
    height=200,
    value="""Transport of lifting machinery into or out of site premises
Setting up of crane on site
Lifting operation
Signalling of load"""
)

generate_ra_pro = st.button("Generate Risk Assessment Pro")

# -------------------------
# helper functions
# -------------------------
def set_cell_text(cell, text):
    cell.text = str(text)

def find_ra_table(doc):
    for table in doc.tables:
        full = " ".join(
            c.text for row in table.rows for c in row.cells
        )
        if "Work Activity" in full and "Hazard" in full:
            return table
    return None

def clear_rows_after_header(table):
    while len(table.rows) > 1:
        row = table.rows[1]
        row._element.getparent().remove(row._element)

def add_row(table, values):
    row = table.add_row().cells
    for i, v in enumerate(values):
        if i < len(row):
            set_cell_text(row[i], v)

# -------------------------
# generate RA PRO
# -------------------------
if generate_ra_pro:
    try:
        with st.spinner("Generating Risk Assessment Pro..."):

            prompt = f"""
Create professional Singapore style 5x5 Risk Assessment.

Company: {company}
Project: {project_name}
Location: {location}
Process: {ra_process}
Machine: {machine}
Description: {description}
Date: {date_input}

Activities:
{activities}

Return JSON only.

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
                "{{process}}": ra_process,
                "{{date}}": str(date_input)
            })

            table = find_ra_table(doc)

            if table:
                clear_rows_after_header(table)

                for r in data["rows"]:
                    add_row(table, [
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

            buffer = BytesIO()
            doc.save(buffer)
            buffer.seek(0)

            st.download_button(
                "Download Risk Assessment Pro",
                buffer,
                "Risk_Assessment_Pro.docx"
            )

    except Exception as e:
        st.error("Risk Assessment Pro generation failed")
        st.exception(e)
