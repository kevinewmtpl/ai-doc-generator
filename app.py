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
                "{{operation_time}}": "To be confirmed",
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
Create a professional Lifting Plan for machinery moving and lifting work in Singapore.

Company: {company}
Project: {project_name}
Location: {location}
Description: {description}
Machine: {machine}

Rules:
- Use formal lifting-plan wording.
- Return plain text content for each field.
- Do not return dictionary-looking text.

Return these fields:
lifting_gear
crew
lifting_method
safety_controls
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
                                "crew": {"type": "string"},
                                "lifting_method": {"type": "string"},
                                "safety_controls": {"type": "string"}
                            },
                            "required": ["lifting_gear", "crew", "lifting_method", "safety_controls"]
                        }
                    }
                }
            )

            data = json.loads(response.output_text)

            doc = Document(LP_TEMPLATE)

            replace_all(doc, {
                "{{company}}": company,
                "{{date}}": str(date_input),
                "{{location}}": location,
                "{{description_of_work}}": description,
                "{{machine_spec}}": machine,
                "{{lifting_gear}}": data["lifting_gear"],
                "{{crew}}": data["crew"],
                "{{lifting_method}}": data["lifting_method"],
                "{{safety_controls}}": data["safety_controls"],
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
