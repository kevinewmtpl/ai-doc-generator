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
# METHOD STATEMENT
# =====================
if generate_ms:
    try:
        with st.spinner("Generating Method Statement..."):
            prompt = f"""
Create a professional Method Statement for machinery moving and lifting work in Singapore.

Company: {company}
Project: {project_name}
Location: {location}
Description: {description}
Machine: {machine}

Rules:
- Use formal contractor wording.
- Return plain text content for each field.
- Do not return dictionary-looking text.
- job_scope must be numbered steps.

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

            doc = Document(MS_TEMPLATE)

            replace_all(doc, {
                "{{company}}": company,
                "{{date}}": str(date_input),
                "{{location}}": location,
                "{{description_of_work}}": description,
                "{{machine_spec}}": machine,
                "{{equipment}}": data["equipment"],
                "{{safety_aspect}}": data["safety_aspect"],
                "{{job_scope}}": data["job_scope"],
                "{{risk_assessment_note}}": "A copy of Risk Assessment will be attached",
                "{{operation_date}}": str(date_input),
                "{{operation_time}}": operation_time if operation_time else "To be confirmed",
                "{{obstacles}}": "To be confirmed",
                "{{environment}}": "To be confirmed",
                "{{lifting_crew}}": "To be confirmed",
                "{{prepared_by}}": "Kevin Wong / Zailani",
            })

            buffer = BytesIO()
            doc.save(buffer)
            buffer.seek(0)

            st.download_button(
                "Download Method Statement",
                buffer,
                "Method_Statement.docx"
            )

    except Exception as e:
        st.error("Method Statement generation failed")
        st.exception(e)

# =====================
# RISK ASSESSMENT
# =====================
if generate_ra:
    try:
        with st.spinner("Generating Risk Assessment..."):
            prompt = f"""
Create a Risk Assessment for machinery moving work.

Company: {company}
Location: {location}
Process: Machinery Moving

Return these fields:
hazards
controls
"""

            response = client.responses.create(
                model="gpt-5.4",
                input=prompt,
                tools=[{
                    "type": "file_search",
                    "vector_store_ids": [RA_VECTOR_STORE_ID]
                }],
                text={
                    "format": {
                        "type": "json_schema",
                        "name": "ra_schema",
                        "schema": {
                            "type": "object",
                            "additionalProperties": False,
                            "properties": {
                                "hazards": {"type": "string"},
                                "controls": {"type": "string"}
                            },
                            "required": ["hazards", "controls"]
                        }
                    }
                }
            )

            data = json.loads(response.output_text)

            doc = Document(RA_TEMPLATE)

            replace_all(doc, {
                "{{company}}": company,
                "{{location}}": location,
                "{{process}}": "Machinery Moving",
                "{{date}}": str(date_input),
                "{{hazards}}": data["hazards"],
                "{{controls}}": data["controls"]
            })

            buffer = BytesIO()
            doc.save(buffer)
            buffer.seek(0)

            st.download_button(
                "Download Risk Assessment",
                buffer,
                "Risk_Assessment.docx"
            )

    except Exception as e:
        st.error("Risk Assessment generation failed")
        st.exception(e)

# =====================
# LIFTING PLAN
# =====================
if generate_lp:
    try:
        with st.spinner("Generating Lifting Plan..."):
            prompt = f"""
Improve and professionalize this lifting method for a lifting plan in Singapore.

Company: {company}
Project: {project_name}
Location: {location}
Description: {description}
Machine: {machine}
Machine dimension: {machine_dimension}
Machine weight: {machine_weight}
Crane name: {crane_name}
Crane SWL: {crane_swl}
Crane radius: {crane_radius}
SWL at radius: {crane_swl_radius}

Sequence of lifting operations:
{task_sequence}

Generate:
- lifting_gear
- lifting_method
- safety_controls

Rules:
- Use formal lifting-plan wording
- Return plain text only
- No dictionary-looking text
"""

            response = client.responses.create(
                model="gpt-5.4",
                input=prompt,
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
                                "lifting_gear": {"type": "string"},
                                "lifting_method": {"type": "string"},
                                "safety_controls": {"type": "string"}
                            },
                            "required": ["lifting_gear", "lifting_method", "safety_controls"]
                        }
                    }
                }
            )

            data = json.loads(response.output_text)

            doc = Document(LP_TEMPLATE)

            replace_all(doc, {
                "{{company}}": company,
                "{{project_name}}": project_name,
                "{{location}}": location,
                "{{date}}": str(date_input),
                "{{operation_date}}": str(date_input),
                "{{operation_time}}": operation_time,
                "{{description_of_work}}": description,
                "{{machine_spec}}": machine,
                "{{machine_name}}": machine,
                "{{machine_dimension}}": machine_dimension,
                "{{machine_weight}}": machine_weight,
                "{{crane_name}}": crane_name,
                "{{crane_renew}}": crane_renew,
                "{{crane_expiry}}": crane_expiry,
                "{{crane_swl}}": crane_swl,
                "{{crane_radius}}": crane_radius,
                "{{crane_swl_radius}}": crane_swl_radius,
                "{{total_swl_lg}}": total_swl_lg,
                "{{lg_expiry}}": lg_expiry,
                "{{lifting_gear}}": data["lifting_gear"],
                "{{lifting_method}}": data["lifting_method"],
                "{{safety_controls}}": data["safety_controls"],
                "{{site_supervisor}}": site_supervisor,
                "{{lifting_supervisor}}": lifting_supervisor,
                "{{equipment_operator}}": equipment_operator,
                "{{rigger_1}}": rigger_1,
                "{{rigger_2}}": rigger_2,
                "{{ground_safe}}": "Yes" if ground_safe else "No",
                "{{outriggers}}": "Yes" if outriggers else "No",
                "{{obstacles}}": "No" if no_overhead_obstacles else "Yes",
                "{{lighting}}": "Yes" if lighting else "No",
                "{{barricade}}": "Yes" if barricade else "No",
                "{{prepared_by}}": "Kevin Wong / Zailani"
            })

            buffer = BytesIO()
            doc.save(buffer)
            buffer.seek(0)

            st.download_button(
                "Download Lifting Plan",
                buffer,
                "Lifting_Plan.docx"
            )

    except Exception as e:
        st.error("Lifting Plan generation failed")
        st.exception(e)

# =====================
# RISK ASSESSMENT PRO ADD-ON
# =====================

st.markdown("---")
st.subheader("Risk Assessment Pro")

st.info("Paste the full RA Pro block I sent previously here next, below this line.")
