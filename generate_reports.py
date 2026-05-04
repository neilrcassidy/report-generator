import pandas as pd
import os
import re
import json
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Flowable, Table, TableStyle, Indenter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.enums import TA_JUSTIFY, TA_CENTER
from reportlab.graphics.shapes import Drawing, Wedge
from reportlab.lib.units import cm

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def config_path(filename="config.json"):
    """
    Returns the path to config/config.json located next to the .exe 
    (or next to .py if running as script)
    """
    if getattr(sys, "frozen", False):  # running as .exe
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, "config", filename)

CONFIG_FILE = config_path()
if not os.path.exists(CONFIG_FILE):
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror("Error", f"Configuration file missing!\n\nPlease make sure the 'config' folder and 'config.json' are located here:\n{os.path.dirname(CONFIG_FILE)}")
    sys.exit(1)

with open(CONFIG_FILE, "r", encoding="utf-8") as f:
    config = json.load(f)

# Colors
COLOR_GREEN_DARK = colors.HexColor(config["colors"]["green_dark"])
COLOR_GREEN_LIGHT = colors.HexColor(config["colors"]["green_light"])
COLOR_YELLOW = colors.HexColor(config["colors"]["yellow"])
COLOR_RED = colors.HexColor(config["colors"]["red"])
COLOR_BLACK = colors.HexColor(config["colors"]["black"])

# Hardcoded Base Columns
BASE_COLUMNS = ["STUDENT", "EXAMS DONE", "CERTIFICATION"]

# Texts
TITLE_TEXT = config["texts"]["title"]
SUBTITLE_TEXT = config["texts"]["subtitle"]
INTRO_TEXT = config["texts"]["intro_text"]
PASS_CONCLUSION_TEXT = config["texts"]["pass_conclusion_text"]
AVERAGE_CONCLUSION_TEXT = config["texts"]["average_conclusion_text"]
FAIL_CONCLUSION_TEXT = config["texts"]["fail_conclusion_text"]
DOUBT_TEXT = config["texts"].get("doubt_text", "")

CONTACTS = config["contacts"]

# -------------------------
# Utility Functions
# -------------------------
def clean_filename(name):
    return re.sub(r'[\/:*?"<>|]', "", name)

def read_input_file(file_path):
    # 1. Read the file
    ext = os.path.splitext(file_path)[1].lower()
    
    try:
        if ext == ".csv":
            df = pd.read_csv(file_path)
        elif ext in [".xlsx", ".xls"]:
            df = pd.read_excel(file_path)
        elif ext == ".ods":
            df = pd.read_excel(file_path, engine="odf")
        else:
            raise ValueError(f"Unsupported file format '{ext}'. Please use .xlsx, .xls, .csv, or .ods.")
    except Exception as e:
        raise ValueError(f"Could not read the file. Ensure it is formatted correctly.\nDetails: {str(e)}")
    
    # 2. Extract ALL unique skills from the config
    all_possible_skills = set()
    for cert_set in config["certification_skills_sets"]:
        all_possible_skills.update(cert_set["skills"])
    all_possible_skills = list(all_possible_skills)

    # 3. Check for base columns
    missing_base = [c for c in BASE_COLUMNS if c not in df.columns]
    if missing_base:
        raise ValueError(f"The following basic columns are missing in the file:\n{', '.join(missing_base)}")

    # --- NEW CODE: DROP GHOST ROWS ---
    # This automatically deletes any hidden spreadsheet rows where the Student, 
    # Exams Done, and Certification are all completely blank.
    df = df.dropna(subset=BASE_COLUMNS, how='all')
    # ---------------------------------

    # 4. Check for empty global data (Name, Exams Done, Cert)
    if df[BASE_COLUMNS].isnull().values.any():
        problem_rows = df[df[BASE_COLUMNS].isnull().any(axis=1)]
        student_names = problem_rows["STUDENT"].fillna("Missing name").tolist()
        raise ValueError(f"There is missing basic data (Name, Exams Done, or Certification) for these students:\n{', '.join(str(s) for s in student_names)}")

    # 5. Check for valid certifications
    valid_certs = []
    for cert_set in config["certification_skills_sets"]:
        valid_certs.extend(cert_set["certifications"])

    invalid_certs_mask = ~df["CERTIFICATION"].isin(valid_certs)
    if invalid_certs_mask.any():
        invalid_students = df[invalid_certs_mask]["STUDENT"].tolist()
        raise ValueError(
            f"Invalid certification for the following students: {', '.join(str(s) for s in invalid_students)}.\n"
            f"Allowed certifications in the configuration are: {', '.join(valid_certs)}"
        )

    # 6. ROW-BY-ROW VALIDATION using certification_skills_sets
    problem_students = []
    for index, row in df.iterrows():
        student = row["STUDENT"]
        cert = row["CERTIFICATION"]
        
        # Get the specific skills required for THIS student's certification
        required_skills = []
        for cert_set in config["certification_skills_sets"]:
            if cert in cert_set["certifications"]:
                required_skills = cert_set["skills"]
                break
        
        for skill in required_skills:
            if skill not in df.columns:
                raise ValueError(f"Column '{skill}' is required for {cert} students but is missing from the file.")
            
            cell_val = row[skill]
            
            # If the required cell is empty
            if pd.isna(cell_val) or str(cell_val).strip() == "":
                problem_students.append(str(student))
                break 
            
            # If the required cell is not a valid number
            try:
                float(cell_val)
            except ValueError:
                problem_students.append(str(student))
                break

    if problem_students:
        raise ValueError(f"Missing or invalid grades in REQUIRED skills for the following students:\n{', '.join(problem_students)}")

    # 7. Filter the DataFrame cleanly (Base Columns + All Possible Skills)
    columns_to_keep = BASE_COLUMNS + [c for c in all_possible_skills if c in df.columns]
    df = df[columns_to_keep].copy()

    # Convert all skill columns to numeric safely
    for col in all_possible_skills:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')

    return df

# -------------------------
# Custom Flowables
# -------------------------
class SectionHeader(Flowable):
    def __init__(self, text, width=200, height=30):
        super().__init__()
        self.text = text
        self.width = width
        self.height = height
    def draw(self):
        self.canv.setFillColor(COLOR_GREEN_DARK)
        self.canv.roundRect(0, 0, self.width, self.height, 10, fill=1, stroke=0)
        self.canv.setFillColor(colors.white)
        self.canv.setFont("Helvetica-Bold", 12)
        self.canv.drawCentredString(self.width/2, self.height/2 - 4, self.text)

class RoundedProgressBar(Flowable):
    def __init__(self, value, limit_pass, limit_average, width=200, height=15):
        super().__init__()
        try: self.value = float(value)
        except: self.value = 0.0
        self.value = max(0.0, min(100.0, self.value))
        self.width = width
        self.height = height
        self.radius = height/2
        self.bar_color = (
            COLOR_GREEN_LIGHT if self.value >= limit_pass else
            COLOR_YELLOW if self.value >= limit_average else
            COLOR_RED
        )
        self.background_color = COLOR_BLACK
    def draw(self):
        canv = self.canv
        canv.saveState()
        canv.setFillColor(self.background_color)
        canv.roundRect(0, 0, self.width, self.height, self.radius, fill=1, stroke=0)
        fill_width = (self.value / 100.0) * self.width
        if fill_width > 0:
            canv.setFillColor(self.bar_color)
            fill_radius = min(self.radius, fill_width/2)
            canv.roundRect(0, 0, fill_width, self.height, fill_radius, fill=1, stroke=0)
        canv.setFillColor(colors.black)
        canv.setFont("Helvetica", 10)
        canv.drawString(self.width+8, (self.height/2)-3, f"{int(round(self.value))} %")
        canv.restoreState()

class SemiCircleGauge(Flowable):
    def __init__(self, value, limit_pass, limit_average, width=180, height=120, radius=70, arc_thickness=14):
        super().__init__()
        try: self.value = float(value)
        except: self.value = 0.0
        self.value = max(0.0, min(100.0, self.value))
        self.width = width
        self.height = height
        self.radius = radius
        self.arc_thickness = arc_thickness
        self.inner_radius = max(1, radius - arc_thickness + 3)
        self.limit_pass = limit_pass
        self.limit_average = limit_average
    def draw(self):
        d = Drawing(self.width, self.height)
        cx, cy = self.width/2, self.arc_thickness+10
        gauge_color = (
            COLOR_GREEN_LIGHT if self.value >= self.limit_pass else
            COLOR_YELLOW if self.value >= self.limit_average else
            COLOR_RED
        )
        progress_angle = 180.0 * (self.value/100.0)
        d.add(Wedge(cx, cy, self.radius, 0, 180, fillColor=COLOR_BLACK, strokeColor=None))
        if progress_angle > 0:
            d.add(Wedge(cx, cy, self.radius, 0, progress_angle, fillColor=gauge_color, strokeColor=None))
            d.add(Wedge(cx, cy, self.inner_radius, 0, progress_angle, fillColor=colors.white, strokeColor=None))
        d.add(Wedge(cx, cy, self.inner_radius, 0, 180, fillColor=colors.white, strokeColor=None))
        self.canv.saveState()
        self.canv.translate(self.width,0)
        self.canv.scale(-1,1)
        d.drawOn(self.canv,0,0)
        self.canv.restoreState()
        self.canv.setFont("Helvetica-Bold", 28)
        self.canv.setFillColor(colors.black)
        self.canv.drawCentredString(self.width/2, cy+5, f"{int(self.value)} %")

# -------------------------
# Header & Footer
# -------------------------
def draw_header(c, doc):
    width, height = A4
    c.saveState()

    c.setFillColor(COLOR_GREEN_DARK)
    c.setFont("Helvetica-Bold", 20)

    # Logo
    logo_path = resource_path("icons/logo.png")
    if os.path.exists(logo_path):
        c.drawImage(
            logo_path,
            40,
            height - 85,
            width=140,
            height=50,
            preserveAspectRatio=True,
            mask='auto'
        )

    # Subtitle
    c.setFont("Helvetica", 14)
    c.setFillColor(COLOR_BLACK)
    subtitle_width = c.stringWidth(SUBTITLE_TEXT, "Helvetica", 14)
    subtitle_x = width - subtitle_width - 40
    c.drawString(subtitle_x, height - 80, SUBTITLE_TEXT)

    # Line
    c.setStrokeColor(COLOR_GREEN_DARK)
    c.setLineWidth(2)
    c.line(40, height - 85, width - 40, height - 85)

    c.restoreState()

def draw_footer(c, doc):
    width, height = A4
    margin_bottom = 0
    FOOTER_HEIGHT = 60
    SIDE_MARGIN = 50
    ICON_WIDTH = 20
    ICON_HEIGHT = 20
    TEXT_ICON_PADDING = 5
    TEXT_LINE_HEIGHT = 12

    phone_numbers = [CONTACTS["phone"].get("landline"), CONTACTS["phone"].get("mobile")]
    phone_lines = [num for num in phone_numbers if num]

    contacts = [
        (resource_path("icons/phone.png"), phone_lines, "tel"),
        (resource_path("icons/email.png"), [CONTACTS["email"]], "mailto"),
        (resource_path("icons/web.png"), [CONTACTS["website"]], "url")
    ]

    c.setFillColor(COLOR_GREEN_DARK)
    c.rect(0, margin_bottom, width, FOOTER_HEIGHT, fill=1, stroke=0)
    footer_center_y = margin_bottom + FOOTER_HEIGHT / 2

    for i, (icon_path, text_lines, link_type) in enumerate(contacts):
        if i == 0:
            icon_x = SIDE_MARGIN
        elif i == 1:
            total_text_width = max(c.stringWidth(line, "Helvetica", 10) for line in text_lines)
            est_width = ICON_WIDTH + TEXT_ICON_PADDING + total_text_width
            icon_x = (width / 2) - (est_width / 2)
        else:
            total_text_width = max(c.stringWidth(line, "Helvetica", 10) for line in text_lines)
            total_block_width = ICON_WIDTH + TEXT_ICON_PADDING + total_text_width
            icon_x = width - SIDE_MARGIN - total_block_width

        icon_y = footer_center_y - ICON_HEIGHT / 2

        if os.path.exists(icon_path):
            c.drawImage(icon_path, icon_x, icon_y, width=ICON_WIDTH, height=ICON_HEIGHT,
                        preserveAspectRatio=True, mask='auto')
        else:
            c.setFillColor(colors.black)
            c.circle(icon_x + ICON_WIDTH / 2, icon_y + ICON_HEIGHT / 2, ICON_WIDTH / 2, fill=1, stroke=0)

        c.setFont("Helvetica", 10)
        c.setFillColor(colors.white)
        num_lines = len(text_lines)
        total_block_height = TEXT_LINE_HEIGHT * num_lines
        text_start_y = footer_center_y + (total_block_height / 2) - TEXT_LINE_HEIGHT + 2
        text_x = icon_x + ICON_WIDTH + TEXT_ICON_PADDING

        for j, line in enumerate(text_lines):
            line_y = text_start_y - j * TEXT_LINE_HEIGHT
            c.drawString(text_x, line_y, line)
            if link_type == "tel":
                link_url = f"tel:{line}"
            elif link_type == "mailto":
                link_url = f"mailto:{line}"
            else:
                link_url = line if line.startswith("http") else f"https://{line}"
            text_width = c.stringWidth(line, "Helvetica", 10)
            text_height = TEXT_LINE_HEIGHT
            c.linkURL(link_url, (text_x, line_y, text_x + text_width, line_y + text_height), relative=0)

# -------------------------
# Generate PDF per student
# -------------------------
def generate_student_pdf(student_row, output_folder, index):
    student = student_row["STUDENT"]
    exam_amount = student_row["EXAMS DONE"]
    certification = student_row["CERTIFICATION"]

    skills = []
    limit_pass = 70
    limit_average = 60
    
    for certification_skills_set in config["certification_skills_sets"]:
        if certification in certification_skills_set["certifications"]:
            skills = certification_skills_set["skills"]
            limit_pass = certification_skills_set.get("pass_threshold", 70)
            limit_average = certification_skills_set.get("average_threshold", 60)
            
    safe_name = clean_filename(str(student))
    output_filename = os.path.join(output_folder, f"{index+1:03d}_{safe_name}.pdf")

    doc = SimpleDocTemplate(
        output_filename,
        pagesize=A4,
        rightMargin=40,
        leftMargin=40,
        topMargin=100,
        bottomMargin=70
    )

    styles = getSampleStyleSheet()
    style_info_label = ParagraphStyle('InfoLabel', parent=styles['Normal'], fontSize=12, leading=13)
    style_info_data = ParagraphStyle('InfoData', parent=styles['Heading3'], fontSize=14, leading=16, spaceAfter=5)
    style_justified = ParagraphStyle('justify', parent=styles['Normal'], alignment=TA_JUSTIFY, leading=12, fontSize=11)

    elements = []
    if INTRO_TEXT:
        formatted_intro = INTRO_TEXT.replace("{min_pass}", str(limit_pass))
        elements.append(Paragraph(formatted_intro, style_justified))
        elements.append(Spacer(1, 10))

    data_info = [
        [Paragraph("<b>ALUMNO/A:</b>", style_info_label),
        Paragraph("<b>CERTIFICACIÓN:</b>", style_info_label)],
        [Paragraph(str(student), style_info_data),
         Paragraph(str(certification), style_info_data)]
    ]
    t_info = Table(data_info, colWidths=[300, 200])
    t_info.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('LEFTPADDING', (0,0), (-1,-1), 0),
        ('BOTTOMPADDING', (0,1), (-1,1), 12)
    ]))
    elements.append(t_info)
    elements.append(Spacer(1, 5))

    data_info = [
        [Paragraph("<b>EXÁMENES REALIZADOS:</b>", style_info_label)],
        [Paragraph(str(exam_amount), style_info_data)]
    ]
    t_info = Table(data_info, colWidths=[500])
    t_info.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('LEFTPADDING', (0,0), (-1,-1), 0),
        ('BOTTOMPADDING', (0,1), (-1,1), 12),
        ('LINEBELOW', (0,1), (-1,1), 1.2, colors.grey),
    ]))
    elements.append(t_info)
    elements.append(Spacer(1, 5))

    notes = [student_row[skill] for skill in skills if pd.notna(student_row.get(skill))]
    avg_note = sum(notes) / len(notes) if notes else 0

    left_content = [SectionHeader("CALIFICACIÓN INDIVIDUAL", width=240, height=25), Spacer(1, 8)]
    for skill, note in zip(skills, notes):
        left_content.append(Paragraph(skill, ParagraphStyle('label', parent=styles['Normal'], leading=12, fontSize=11)))
        left_content.append(Spacer(1, 3))
        left_content.append(RoundedProgressBar(note, limit_pass, limit_average, width=210))
        left_content.append(Spacer(1, 15))

    # Determine status text based on the three tiers
    if avg_note >= limit_pass:
        status_text = "APTO/A"
    elif avg_note >= limit_average:
        status_text = "APTO/A<br/>(con observaciones)"
    else:
        status_text = "NO APTO/A"

    right_content = [
        SectionHeader("CALIFICACIÓN TOTAL", width=200, height=25),
        Indenter(10, 0),
        SemiCircleGauge(avg_note, limit_pass, limit_average),
        Spacer(1, 2),
        Indenter(-10, 0),
        SectionHeader("ESTADO", width=200, height=20),
        Paragraph(status_text, ParagraphStyle('status', parent=styles['Heading3'], alignment=TA_CENTER, fontSize=16)),
    ]

    main_table_data = [[left_content, right_content]]
    main_table = Table(main_table_data, colWidths=[270, 230])
    main_table.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('LEFTPADDING', (1,0), (1,0), 20)
    ]))
    elements.append(main_table)
    elements.append(Spacer(1, 10))

    mobile_num = CONTACTS["phone"].get("mobile", "")
    clean_num = mobile_num.replace(" ", "")
    
    formatted_doubt_text = ""
    if DOUBT_TEXT and mobile_num:
        formatted_doubt_text = DOUBT_TEXT.format(mobile_clean=clean_num, mobile_display=mobile_num)

    if avg_note >= limit_pass and PASS_CONCLUSION_TEXT:
        elements.append(Paragraph(PASS_CONCLUSION_TEXT + formatted_doubt_text, style_justified))
    elif avg_note < limit_pass and avg_note >= limit_average and AVERAGE_CONCLUSION_TEXT:
        elements.append(Paragraph(AVERAGE_CONCLUSION_TEXT + formatted_doubt_text, style_justified))
    elif avg_note < limit_average and FAIL_CONCLUSION_TEXT:
        elements.append(Paragraph(FAIL_CONCLUSION_TEXT + formatted_doubt_text, style_justified))

    doc.build(
        elements,
        onFirstPage=lambda c,d: (draw_header(c,d), draw_footer(c,d)),
        onLaterPages=lambda c,d: (draw_header(c,d), draw_footer(c,d))
    )

# -------------------------
# GUI File Selection & Run
# -------------------------
def choose_file_and_run():
    root = tk.Tk()
    root.withdraw()

    file_path = filedialog.askopenfilename(
        title="Select input file",
        filetypes=[
            ("All Supported Files", "*.xlsx *.xls *.csv *.ods"),
            ("Excel Files", "*.xlsx *.xls"),
            ("CSV Files", "*.csv"),
            ("OpenDocument Spreadsheet", "*.ods")
        ]
    )

    if not file_path:
        messagebox.showinfo("No file selected", "You must select a valid data file to continue.")
        return

    output_folder = "reports"
    os.makedirs(output_folder, exist_ok=True)

    try:
        df = read_input_file(file_path)
        for index, row in df.iterrows():
            generate_student_pdf(row, output_folder, index)
        messagebox.showinfo("Success", f"All PDFs generated in '{output_folder}'!")
    except Exception as e:
        messagebox.showerror("Error", str(e))

if __name__ == "__main__":
    choose_file_and_run()