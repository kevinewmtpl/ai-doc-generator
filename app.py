import os
import json
import base64
from io import BytesIO
from datetime import date, timedelta

import streamlit as st
from openai import OpenAI
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn

# PowerPoint support for Method Statement PRO
from pptx import Presentation
from pptx.util import Pt as PPTPt

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
MOS_PRO_TEMPLATE = os.path.join(BASE_DIR, "Templates", "MOS New.pptx")
RA_TEMPLATE = os.path.join(BASE_DIR, "Templates", "RA Template.docx")
LP_TEMPLATE = os.path.join(BASE_DIR, "Templates", "Lifting Plan Template.docx")


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



# =====================
# METHOD STATEMENT PRO - POWERPOINT HELPERS
# =====================
def mos_pro_template_path():
    """Return the preferred MOS PRO PowerPoint template path."""
    candidates = [
        MOS_PRO_TEMPLATE,
        os.path.join(BASE_DIR, "MOS New.pptx"),
    ]
    for path in candidates:
        if os.path.exists(path):
            return path
    return MOS_PRO_TEMPLATE


def mos_pro_date(value):
    """Format dates to match the existing EWMT PowerPoint style."""
    try:
        return value.strftime("%d.%m.%Y")
    except Exception:
        return str(value)


def mos_pro_safe_filename(value):
    name = re.sub(r"[^A-Za-z0-9_-]+", "_", str(value or "Project").strip())
    return (name[:60] or "Project").strip("_")


def mos_pro_set_paragraph_text(paragraph, new_text):
    """Replace paragraph text while retaining the first run's formatting."""
    if paragraph.runs:
        paragraph.runs[0].text = str(new_text)
        for run in paragraph.runs[1:]:
            run.text = ""
    else:
        paragraph.text = str(new_text)


def mos_pro_replace_label_in_text_frame(text_frame, patterns):
    """
    Replace only the variable value after a known label.
    `patterns` is a list of (compiled_regex, replacement_value).
    """
    changed = 0
    for paragraph in text_frame.paragraphs:
        original = "".join(run.text for run in paragraph.runs) if paragraph.runs else paragraph.text
        if not original:
            continue

        new_text = original
        for regex, value in patterns:
            if regex.search(new_text):
                new_text = regex.sub(lambda m: m.group(1) + str(value), new_text)
                break

        if new_text != original:
            mos_pro_set_paragraph_text(paragraph, new_text)
            changed += 1
    return changed


def mos_pro_update_project_fields(prs, project):
    """
    Update the repeated project information throughout MOS New.pptx.
    The standard content, drawings, pictures and page layouts are left untouched.
    """
    patterns = [
        (re.compile(r"^(\s*Customer\s*/\s*Tenant\s*Company\s*:\s*).*$", re.I), project.get("customer", "")),
        (re.compile(r"^(\s*Site\s*Location\s*:\s*).*$", re.I), project.get("site", "")),
        (re.compile(r"^(\s*Process\s*:\s*).*$", re.I), project.get("process", "")),
        (re.compile(r"^(\s*Prepared\s+by\s*:\s*).*$", re.I), project.get("prepared_by", "")),
        (re.compile(r"^(\s*Approved\s+By\s*:\s*).*$", re.I), project.get("approved_by", "")),
        (re.compile(r"^(\s*Last\s+Review\s+Date\s*:\s*).*$", re.I), project.get("last_review", "")),
        (re.compile(r"^(\s*Next\s+Review\s+Date\s*:\s*).*$", re.I), project.get("next_review", "")),
    ]

    changed = 0
    for slide in prs.slides:
        for shape in slide.shapes:
            # Text boxes / placeholders
            if getattr(shape, "has_text_frame", False):
                changed += mos_pro_replace_label_in_text_frame(shape.text_frame, patterns)

            # Tables - this is where the repeated EWMT project block is located.
            if getattr(shape, "has_table", False):
                for row in shape.table.rows:
                    for cell in row.cells:
                        changed += mos_pro_replace_label_in_text_frame(cell.text_frame, patterns)

            # Grouped shapes - safe recursive handling for future template edits.
            if getattr(shape, "shape_type", None) == 6:  # GROUP
                for subshape in shape.shapes:
                    if getattr(subshape, "has_text_frame", False):
                        changed += mos_pro_replace_label_in_text_frame(subshape.text_frame, patterns)
                    if getattr(subshape, "has_table", False):
                        for row in subshape.table.rows:
                            for cell in row.cells:
                                changed += mos_pro_replace_label_in_text_frame(cell.text_frame, patterns)

    return changed


def mos_pro_generate_work_steps(description, machine, equipment, site_notes):
    """Generate optional project-specific work steps for the existing Work Method slide."""
    response = client.responses.create(
        model="gpt-5.4",
        input=f"""
Create a concise project-specific work method for Eric Wong Machinery Transportation Pte Ltd.
This content will be inserted into the existing PowerPoint Method Statement work-method page.

Description of work:
{description}

Machine / load:
{machine}

Equipment intended:
{equipment}

Site / access notes:
{site_notes}

Requirements:
- Singapore machinery moving / lifting contractor wording.
- 10 to 16 clear sequential work steps.
- Use only equipment actually stated above.
- Include barricading / exclusion zone, pre-use checks and toolbox briefing where relevant.
- Include trial lift only when crane / lorry crane lifting is part of the stated work.
- Include controlled landing / shifting / positioning relevant to the actual description.
- Include housekeeping and completion.
- Do not invent dimensions, crane capacities, registration numbers, personnel names or certificates.
- Return JSON only.
""",
        tools=[{
            "type": "file_search",
            "vector_store_ids": [MS_VECTOR_STORE_ID]
        }],
        text={
            "format": {
                "type": "json_schema",
                "name": "mos_pro_work_steps",
                "schema": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "steps": {
                            "type": "array",
                            "items": {"type": "string"},
                            "minItems": 8,
                            "maxItems": 16
                        }
                    },
                    "required": ["steps"]
                }
            }
        }
    )
    return json.loads(response.output_text).get("steps", [])


def mos_pro_update_work_method_page(prs, steps):
    """
    Replace only the existing 'Work Method Statement for Lifting Operation' text box.
    All other slides and objects remain untouched.
    """
    if not steps:
        return False

    target = None
    for slide in prs.slides:
        for shape in slide.shapes:
            if getattr(shape, "has_text_frame", False) and "Work Method Statement for Lifting Operation" in shape.text:
                target = shape
                break
        if target is not None:
            break

    if target is None:
        return False

    tf = target.text_frame
    tf.clear()
    tf.word_wrap = True

    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "Work Method Statement for Lifting Operation:"
    r.font.size = PPTPt(11)
    r.font.bold = True

    p = tf.add_paragraph()
    r = p.add_run()
    r.text = "The following work shall be carried out for the planned operation:"
    r.font.size = PPTPt(10.5)

    for idx, step in enumerate(steps, start=1):
        p = tf.add_paragraph()
        r = p.add_run()
        r.text = f"{idx}. {str(step).strip()}"
        r.font.size = PPTPt(10.2)

    return True


def mos_pro_build_powerpoint(project, replace_work_method=False, work_inputs=None):
    """Open the user's master PowerPoint, edit only selected fields, and return a PPTX buffer."""
    template_path = mos_pro_template_path()
    if not os.path.exists(template_path):
        raise FileNotFoundError(
            "MOS New.pptx was not found. Put the PowerPoint in Templates/MOS New.pptx in GitHub."
        )

    prs = Presentation(template_path)
    change_count = mos_pro_update_project_fields(prs, project)

    steps = []
    work_updated = False
    if replace_work_method:
        work_inputs = work_inputs or {}
        steps = mos_pro_generate_work_steps(
            work_inputs.get("description", ""),
            work_inputs.get("machine", ""),
            work_inputs.get("equipment", ""),
            work_inputs.get("site_notes", ""),
        )
        work_updated = mos_pro_update_work_method_page(prs, steps)

    output = BytesIO()
    prs.save(output)
    output.seek(0)
    return output, change_count, work_updated, steps, len(prs.slides)


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

    # Separate BETA module: this does not replace or modify the existing Word Method Statement.
    pro_col, pro_blank1, pro_blank2 = st.columns(3)
    with pro_col:
        dashboard_card(
            "📑 Method Statement PRO",
            "Build a copy of the professional EWMT PowerPoint MOS template while keeping the current Word Method Statement fully operational.",
            METHOD_IMAGE,
            "POWERPOINT BETA"
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
        ms_date_input = st.date_input("Date", value=date.today(), key="ms_date_input")
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
                    "{{date}}": str(ms_date_input),
                    "{{location}}": safe_text(ms_location),
                    "{{description_of_work}}": safe_text(ms_description),
                    "{{machine_spec}}": safe_text(ms_machine),
                    "{{equipment}}": safe_text(data["equipment"]),
                    "{{safety_aspect}}": safe_text(data["safety_aspect"]),
                    "{{job_scope}}": safe_text(data["job_scope"]),
                    "{{risk_assessment_note}}": "Refer as attached",
                    "{{operation_date}}": str(ms_date_input),
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
# METHOD STATEMENT PRO - POWERPOINT BETA
# ======================================================
if page == "📑 Method Statement PRO":
    st.markdown("## 📑 Method Statement PRO")
    st.caption("PowerPoint-based professional Method Statement. Your existing Word Method Statement remains completely separate and unchanged.")

    st.info(
        "BETA approach: the system opens your original `MOS New.pptx`, keeps the existing slides/layout, "
        "updates the repeated project-information block, and downloads a new editable PowerPoint copy. "
        "You can improve this PRO module gradually without affecting the current working Method Statement."
    )

    template_path = mos_pro_template_path()
    if os.path.exists(template_path):
        try:
            template_prs = Presentation(template_path)
            st.success(f"MOS PRO template found • {len(template_prs.slides)} slides • `{os.path.relpath(template_path, BASE_DIR)}`")
        except Exception as exc:
            st.error(f"MOS New.pptx exists but cannot be opened: {exc}")
    else:
        st.error("MOS PRO template not found.")
        st.code("Templates/MOS New.pptx")
        st.warning("Upload the exact `MOS New.pptx` PowerPoint into the Templates folder in your GitHub repository. The existing Word Method Statement will continue to work even if this PRO template is missing.")

    with st.expander("1. Project Information — updates all repeated MOS header/footer blocks", expanded=True):
        p1, p2 = st.columns(2)
        with p1:
            pro_customer = st.text_input("Customer / Tenant Company", key="pro_customer")
            pro_site = st.text_input("Site Location", key="pro_site")
            pro_process = st.text_input(
                "Process",
                value="Lifting and Moving of Machinery",
                key="pro_process"
            )
            pro_prepared_by = st.text_input("Prepared By", value="Kevin Wong", key="pro_prepared_by")
        with p2:
            pro_approved_by = st.text_input("Approved By", value="Eric Wong (Director)", key="pro_approved_by")
            pro_last_review = st.date_input("Last Review Date", value=date.today(), key="pro_last_review")
            pro_next_review = st.date_input(
                "Next Review Date",
                value=date.today() + timedelta(days=730),
                key="pro_next_review"
            )
            pro_file_name = st.text_input(
                "Output File Name",
                value="EWMT_Method_Statement_PRO",
                key="pro_file_name"
            )

        st.caption(
            "These values replace Customer, Site, Process, Prepared By, Approved By and review dates throughout the PowerPoint. "
            "The standard wording, diagrams, pictures and slide positions are not intentionally rewritten."
        )

    with st.expander("2. Optional — Replace Existing Work Method Page with AI Job Steps", expanded=False):
        pro_replace_work = st.checkbox(
            "Replace the existing 'Work Method Statement for Lifting Operation' page",
            value=False,
            key="pro_replace_work"
        )
        st.caption("Leave this OFF while you are developing the template. With it OFF, the original work-method page remains exactly as supplied in MOS New.pptx.")

        if pro_replace_work:
            pro_description = st.text_area(
                "Description of Work",
                height=110,
                key="pro_description",
                placeholder="Example: Unload one machining centre from low-bed trailer, shift into factory and position at designated location."
            )
            pro_machine = st.text_input(
                "Machine / Load Details",
                key="pro_machine",
                placeholder="Example: Mazak Variaxis i700T, 8,500 kg"
            )
            pro_equipment = st.text_area(
                "Equipment Actually Intended",
                height=90,
                key="pro_equipment",
                placeholder="Example: 70T lorry crane, 10T forklift, hydraulic jacks, machine skates, webbing slings, shackles"
            )
            pro_site_notes = st.text_area(
                "Site / Access Notes",
                height=90,
                key="pro_site_notes",
                placeholder="Example: Ground-floor concrete access, enter via Roller Shutter 2, protect epoxy flooring with plywood."
            )
        else:
            pro_description = ""
            pro_machine = ""
            pro_equipment = ""
            pro_site_notes = ""

    with st.expander("3. What This BETA Version Will / Will Not Change", expanded=False):
        st.markdown(
            """
**Will change**
- Customer / Tenant Company across the deck
- Site Location across the deck
- Process across the deck
- Prepared By / Approved By
- Last Review / Next Review dates
- Optional existing work-method text box only when you explicitly enable it

**Will not intentionally change**
- The existing Word Method Statement module
- Standard MOS wording on the fixed pages
- EWMT logo / company heading
- Existing diagrams / photographs / certificates
- Slide order and slide numbers
- Risk Assessment and Lifting Plan modules
            """
        )

    generate_pro = st.button(
        "📑 Generate Method Statement PRO PowerPoint",
        key="generate_method_statement_pro",
        type="primary"
    )

    if generate_pro:
        if not os.path.exists(template_path):
            st.error("Cannot generate MOS PRO until `Templates/MOS New.pptx` exists in the repository.")
        elif not pro_customer.strip() or not pro_site.strip():
            st.error("Please enter at least Customer / Tenant Company and Site Location.")
        elif pro_replace_work and not pro_description.strip():
            st.error("Please enter Description of Work before replacing the work-method page.")
        else:
            try:
                project_data = {
                    "customer": pro_customer.strip(),
                    "site": pro_site.strip(),
                    "process": pro_process.strip(),
                    "prepared_by": pro_prepared_by.strip(),
                    "approved_by": pro_approved_by.strip(),
                    "last_review": mos_pro_date(pro_last_review),
                    "next_review": mos_pro_date(pro_next_review),
                }

                with st.spinner("Preparing editable MOS PRO PowerPoint from the master template..."):
                    output, updated_fields, work_updated, generated_steps, slide_count = mos_pro_build_powerpoint(
                        project=project_data,
                        replace_work_method=pro_replace_work,
                        work_inputs={
                            "description": pro_description,
                            "machine": pro_machine,
                            "equipment": pro_equipment,
                            "site_notes": pro_site_notes,
                        }
                    )

                st.success(
                    f"MOS PRO generated successfully • {slide_count} slides • "
                    f"{updated_fields} project-information fields updated."
                )

                if pro_replace_work:
                    if work_updated:
                        st.success("The existing Work Method Statement page was replaced with project-specific steps.")
                    else:
                        st.warning("The PowerPoint was generated, but the Work Method Statement text box could not be located. Its original content was kept.")

                    if generated_steps:
                        with st.expander("Preview generated work steps"):
                            for i, step in enumerate(generated_steps, start=1):
                                st.write(f"{i}. {step}")

                output_name = mos_pro_safe_filename(pro_file_name)
                st.download_button(
                    "⬇️ Download Method Statement PRO (.pptx)",
                    output,
                    file_name=f"{output_name}.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True
                )

            except Exception as e:
                st.error("Method Statement PRO generation failed")
                st.exception(e)


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
        lp_date_input = st.date_input("Date", value=date.today(), key="lp_date_input")
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
        crane_renew = st.text_input("Date of Last Certification", key="crane_renew")
        crane_expiry = st.text_input("Expiry Date of Certificate", key="crane_expiry")
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

        lg_expiry = st.text_input("Expiry Date of Lifting Gear Certificate", key="lg_expiry")

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
                    "{{date}}": str(lp_date_input),
                    "{{operation_date}}": str(lp_date_input),
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
                    "{{crane_renew}}": safe_text(crane_renew),
                    "{{crane_expiry}}": safe_text(crane_expiry),
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
                    "{{lg_expiry}}": safe_text(lg_expiry),

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
        ra_date_input = st.date_input("Date", value=date.today(), key="ra_date_input")
        ra_due_date_input = st.date_input("Due Date", value=date.today(), key="ra_due_date_input")

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
Date: {ra_date_input}
Due Date: {ra_due_date_input}

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
    "due_date":"{ra_due_date_input}",
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
                    "{{date}}": str(ra_date_input),
                    "{{due_date}}": str(ra_due_date_input)
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
                            str(ra_due_date_input),
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
                        "Expiry Date": str(found_date),
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

            st.info("For this to work, put expiry date inside the certificate filename, example: 3 Ton Shackle Expiry 2026-06-30.pdf")


# ======================================================
# SETTINGS
# ======================================================
if page == "⚙️ Settings":
    st.markdown("## ⚙️ Settings")
    st.info("Method Statement PRO uses `Templates/MOS New.pptx`. Keep this file in the Templates folder. The original Word Method Statement continues to use `Templates/Method of statement Template.docx`.")
    st.info("This module can be added next: manage default names, templates and company settings.")

    st.markdown("### Method Statement Placeholder Guide")

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

    st.markdown("### Lifting Plan Placeholder Guide")

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
