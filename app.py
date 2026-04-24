import os
import json
from io import BytesIO
from datetime import date

import streamlit as st
from openai import OpenAI
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn

# =====================
# PAGE CONFIG + STYLE
# =====================
st.set_page_config(
    page_title="EWMT Document Generator",
    page_icon="🏗️",
    layout="wide"
)

st.markdown("""
<style>
.block-container {
    padding-top: 1.5rem;
    padding-bottom: 2rem;
    max-width: 1400px;
}

.ewmt-header {
    background: linear-gradient(90deg, #0f172a, #1e3a8a);
    padding: 24px 30px;
    border-radius: 18px;
    color: white;
    margin-bottom: 25px;
    box-shadow: 0px 4px 14px rgba(0,0,0,0.18);
}

.ewmt-title {
    font-size: 34px;
    font-weight: 800;
    margin-bottom: 4px;
}

.ewmt-subtitle {
    font-size: 17px;
    color: #dbeafe;
}

.stButton > button {
    background-color: #1e3a8a;
    color: white;
    border-radius: 10px;
    padding: 0.6rem 1.2rem;
    font-weight: 700;
    border: none;
}

.stButton > button:hover {
    background-color: #0f172a;
    color: white;
}

[data-testid="stTabs"] button {
    font-size: 16px;
    font-weight: 700;
}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="ewmt-header">
    <div class="ewmt-title">Eric Wong Machinery Transportation Pte Ltd</div>
    <div class="ewmt-subtitle">
        Heavy Machinery Moving • Lifting • Transportation • AI Document Generator
    </div>
</div>
""", unsafe_allow_html=True)

st.markdown("## AI Document Generator Dashboard")

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
# COMMON FUNCTIONS
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


def format_method_statement(doc):
    headings = [
        "METHOD OF STATEMENT",
        "Description of work",
        "Machine Spec",
        "Risk Assessment",
        "Date / Time of Operation",
        "Location of Operation",
        "Equipment Use",
        "Obstacles",
        "Environment",
        "Lifting Crew",
        "Safety Aspect",
        "Job Scope",
    ]

    for para in doc.paragraphs:
        txt = para.text.strip()

        for run in para.runs:
            run.font.name = "Times New Roman"
            run._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")

            if txt in headings:
                run.font.size = Pt(16)
                run.bold = True
                run.underline = True
            else:
                run.font.size = Pt(12)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    txt = para.text.strip()

                    for run in para.runs:
                        run.font.name = "Times New Roman"
                        run._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")

                        if txt in headings:
                            run.font.size = Pt(16)
                            run.bold = True
                            run.underline = True
                        else:
                            run.font.size = Pt(12)


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


# =====================
# TABS
# =====================
tab_ms, tab_lp, tab_ra = st.tabs([
    "📄 Method Statement",
    "🏗️ Lifting Plan",
    "⚠️ Risk Assessment Pro"
])

# ======================================================
# TAB 1 - METHOD STATEMENT
# ======================================================
with tab_ms:
    st.header("Method Statement")

    with st.expander("Project Details", expanded=True):
        ms_company = st.text_input("Company", "Eric Wong Machinery Transportation Pte Ltd", key="ms_company")
        ms_project_name = st.text_input("Project Name", key="ms_project_name")
        ms_location = st.text_input("Location of Operation", key="ms_location")
        ms_description = st.text_area("Description of Work", key="ms_description")
        ms_machine = st.text_input("Machine Spec", key="ms_machine")
        ms_date_input = st.date_input("Date", value=date.today(), key="ms_date_input")
        ms_operation_time = st.text_input("Operation Date & Time", key="ms_operation_time")

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
                    "{{company}}": ms_company,
                    "{{date}}": str(ms_date_input),
                    "{{location}}": ms_location,
                    "{{description_of_work}}": ms_description,
                    "{{machine_spec}}": ms_machine,
                    "{{equipment}}": data["equipment"],
                    "{{safety_aspect}}": data["safety_aspect"],
                    "{{job_scope}}": data["job_scope"],
                    "{{risk_assessment_note}}": "A copy of Risk Assessment will be attached",
                    "{{operation_date}}": str(ms_date_input),
                    "{{operation_time}}": ms_operation_time if ms_operation_time else "To be confirmed",
                    "{{obstacles}}": "To be confirmed",
                    "{{environment}}": "To be confirmed",
                    "{{lifting_crew}}": "To be confirmed",
                    "{{prepared_by}}": "Kevin Wong / Zailani",
                })

                format_method_statement(doc)

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
# TAB 2 - LIFTING PLAN
# ======================================================
with tab_lp:
    st.header("Lifting Plan")

    with st.expander("Project Details", expanded=True):
        lp_company = st.text_input("Company", "Eric Wong Machinery Transportation Pte Ltd", key="lp_company")
        lp_project_name = st.text_input("Project Name", key="lp_project_name")
        lp_location = st.text_input("Location of Lifting Operation", key="lp_location")
        lp_description = st.text_area("Description of Load / Work", key="lp_description")
        lp_machine = st.text_input("Machine Name / Spec", key="lp_machine")
        lp_date_input = st.date_input("Date", value=date.today(), key="lp_date_input")
        lp_operation_time = st.text_input("Operation Time", key="lp_operation_time")

    with st.expander("Load Details", expanded=True):
        lp_machine_dimension = st.text_input("Machine Dimension", key="lp_machine_dimension")
        lp_machine_weight = st.text_input("Machine Weight", key="lp_machine_weight")

    with st.expander("Lifting Equipment Details", expanded=True):
        lifting_equipment_type = st.selectbox(
            "Type of Lifting Equipment",
            ["Mobile crane", "Lorry loader"],
            key="lifting_equipment_type"
        )

        lifting_gear_manual = st.text_area(
            "Type of Lifting Gears / Equipment Details",
            height=160,
            value="""Mobile crane of suitable capacity
Wire rope slings / webbing slings
Shackles
Spreader beam if required
Timber mats / steel plates
Tag lines""",
            key="lifting_gear_manual"
        )

        crane_name = st.text_input("Crane Name / Model", key="crane_name")
        crane_renew = st.text_input("Crane Cert Date", key="crane_renew")
        crane_expiry = st.text_input("Crane Cert Expiry", key="crane_expiry")
        crane_swl = st.text_input("Crane SWL", key="crane_swl")
        crane_radius = st.text_input("Crane Radius", key="crane_radius")
        crane_swl_radius = st.text_input("SWL at Radius", key="crane_swl_radius")
        total_swl_lg = st.text_input("Total SWL of Lifting Gear", key="total_swl_lg")
        lg_expiry = st.text_input("Lifting Gear Expiry", key="lg_expiry")

    with st.expander("Personnel", expanded=True):
        site_supervisor = st.text_input("Site Supervisor", key="site_supervisor")
        lifting_supervisor = st.text_input("Lifting Supervisor", key="lifting_supervisor")
        equipment_operator = st.text_input("Equipment Operator", key="equipment_operator")
        rigger_1 = st.text_input("Rigger 1", key="rigger_1")
        rigger_2 = st.text_input("Rigger 2", key="rigger_2")

    with st.expander("Site Conditions", expanded=True):
        ground_safe = st.checkbox("Ground Safe", value=True, key="ground_safe")
        outriggers = st.checkbox("Outriggers Extended", value=True, key="outriggers")
        no_overhead_obstacles = st.checkbox("No Overhead Obstacles", value=True, key="no_overhead_obstacles")
        lighting = st.checkbox("Lighting Adequate", value=True, key="lighting")
        barricade = st.checkbox("Area Barricaded", value=True, key="barricade")

    with st.expander("Lifting Steps", expanded=True):
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
8. Remove lifting gear and carry out housekeeping""",
            key="task_sequence"
        )

    generate_lp = st.button("🏗️ Generate Lifting Plan", key="generate_lp")

    if generate_lp:
        try:
            with st.spinner("Generating Lifting Plan..."):
                prompt = f"""
Improve and professionalize this lifting method for a lifting plan in Singapore.

Company: {lp_company}
Project: {lp_project_name}
Location: {lp_location}
Description: {lp_description}
Machine: {lp_machine}
Machine dimension: {lp_machine_dimension}
Machine weight: {lp_machine_weight}
Crane name: {crane_name}
Crane SWL: {crane_swl}
Crane radius: {crane_radius}
SWL at radius: {crane_swl_radius}
Type of lifting equipment selected: {lifting_equipment_type}
Type of lifting gears / equipment details:
{lifting_gear_manual}

Sequence of lifting operations:
{task_sequence}

Generate:
- lifting_gear
- lifting_method
- safety_controls

Rules:
- Use formal lifting-plan wording.
- Return plain text only.
- No dictionary-looking text.
- Do not include explanation, justification, summary, or reference to company format.
- Return only the actual content to be inserted into the Word document.
- lifting_gear may be generated but will be overridden by user's manual input.
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
                    "{{company}}": lp_company,
                    "{{project_name}}": lp_project_name,
                    "{{location}}": lp_location,
                    "{{date}}": str(lp_date_input),
                    "{{operation_date}}": str(lp_date_input),
                    "{{operation_time}}": lp_operation_time,
                    "{{description_of_work}}": lp_description,
                    "{{machine_spec}}": lp_machine,
                    "{{machine_name}}": lp_machine,
                    "{{machine_dimension}}": lp_machine_dimension,
                    "{{machine_weight}}": lp_machine_weight,
                    "{{crane_name}}": crane_name,
                    "{{crane_renew}}": crane_renew,
                    "{{crane_expiry}}": crane_expiry,
                    "{{crane_swl}}": crane_swl,
                    "{{crane_radius}}": crane_radius,
                    "{{crane_swl_radius}}": crane_swl_radius,
                    "{{total_swl_lg}}": total_swl_lg,
                    "{{lg_expiry}}": lg_expiry,
                    "{{lifting_gear}}": lifting_gear_manual,
                    "{{lifting_method}}": data["lifting_method"],
                    "{{safety_controls}}": data["safety_controls"],
                    "{{mobile_crane_checked}}": "☒" if lifting_equipment_type == "Mobile crane" else "☐",
                    "{{lorry_loader_checked}}": "☒" if lifting_equipment_type == "Lorry loader" else "☐",
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
                    "Lifting_Plan.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

        except Exception as e:
            st.error("Lifting Plan generation failed")
            st.exception(e)


# ======================================================
# TAB 3 - RISK ASSESSMENT PRO
# ======================================================
with tab_ra:
    st.header("Risk Assessment Pro")

    with st.expander("Project Details", expanded=True):
        ra_company = st.text_input("Company", "Eric Wong Machinery Transportation Pte Ltd", key="ra_company")
        ra_project_name = st.text_input("Project Name", key="ra_project_name")
        ra_location = st.text_input("Location", key="ra_location")
        ra_machine = st.text_input("Machine Spec", key="ra_machine")
        ra_description = st.text_area("Description of Work", key="ra_description")
        ra_date_input = st.date_input("Date", value=date.today(), key="ra_date_input")

    with st.expander("Risk Assessment Details", expanded=True):
        ra_process = st.text_input(
            "RA Process",
            "Machinery Moving / Lifting Operation",
            key="ra_process"
        )

        activities = st.text_area(
            "Work Activities (1 per line)",
            height=200,
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
Date: {ra_date_input}

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
    "due_date":"{ra_date_input}",
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
                    "{{date}}": str(ra_date_input)
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
                            r["due_date"],
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
