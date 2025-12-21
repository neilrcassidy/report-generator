import pandas as pd
import argparse
import os
import re
import json
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Flowable, Table, TableStyle, Indenter, KeepTogether
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.enums import TA_JUSTIFY, TA_CENTER
from reportlab.graphics.shapes import Drawing, Wedge, Rect
from reportlab.lib.units import cm

# -------------------------
# Load Config
# -------------------------
with open("config.json", "r", encoding="utf-8") as f:
    config = json.load(f)

# Colors
COLOR_GREEN_DARK = colors.HexColor(config["colors"]["green_dark"])
COLOR_GREEN_LIGHT = colors.HexColor(config["colors"]["green_light"])
COLOR_YELLOW = colors.HexColor(config["colors"]["yellow"])
COLOR_RED = colors.HexColor(config["colors"]["red"])
COLOR_BLACK = colors.HexColor(config["colors"]["black"])

# Thresholds
LIMIT_GREEN = config["thresholds"]["green"]
LIMIT_YELLOW = config["thresholds"]["yellow"]

# Skills & Expected Columns
SKILLS = config["skills"]
EXPECTED_COLUMNS = config["expected_columns"]

# Texts
TITLE_TEXT = config["texts"]["title"]
SUBTITLE_TEXT = config["texts"]["subtitle"]
INTRO_TEXT = config["texts"]["intro_text"]
POSITIVE_CONCLUSION_TEXT = config["texts"]["positive_conclusion_text"]
NEGATIVE_CONCLUSION_TEXT = config["texts"]["negative_conclusion_text"]

# Contacts
CONTACTS = config["contacts"]

# Passing note
PASSING_NOTE = config["min_passing_note"]

# -------------------------
# Utility Functions
# -------------------------
def sanitize_filename(name):
    return re.sub(r'[\/:*?"<>|]', "", name)

def read_input_file(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".csv":
        df = pd.read_csv(file_path)
    elif ext in [".xlsx", ".xls"]:
        df = pd.read_excel(file_path)
    else:
        raise ValueError("Unsupported file type. Use .csv or .xlsx")
    missing = [c for c in EXPECTED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"Missing expected columns: {missing}")
    return df[EXPECTED_COLUMNS]

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
    def __init__(self, value, width=200, height=15):
        super().__init__()
        try: self.value = float(value)
        except: self.value = 0.0
        self.value = max(0.0, min(100.0, self.value))
        self.width = width
        self.height = height
        self.radius = height/2
        self.bar_color = (
            COLOR_GREEN_LIGHT if self.value >= LIMIT_GREEN else
            COLOR_YELLOW if self.value >= LIMIT_YELLOW else
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
        canv.drawString(self.width+8, (self.height/2)-3, f"{int(self.value)} %")
        canv.restoreState()

class SemiCircleGauge(Flowable):
    def __init__(self, value, width=180, height=120, radius=70, arc_thickness=14):
        super().__init__()
        try: self.value = float(value)
        except: self.value = 0.0
        self.value = max(0.0, min(100.0, self.value))
        self.width = width
        self.height = height
        self.radius = radius
        self.arc_thickness = arc_thickness
        self.inner_radius = max(1, radius - arc_thickness + 3)
    def draw(self):
        d = Drawing(self.width, self.height)
        cx, cy = self.width/2, self.arc_thickness+10
        gauge_color = (
            COLOR_GREEN_LIGHT if self.value >= LIMIT_GREEN else
            COLOR_YELLOW if self.value >= LIMIT_YELLOW else
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
        self.canv.setFont("Helvetica-Bold", 28)  # Hardcoded font size
        self.canv.setFillColor(colors.black)
        self.canv.drawCentredString(self.width/2, cy+5, f"{int(self.value)} %")

# -------------------------
# Header & Footer
# -------------------------
def draw_header(c, doc):
    width, height = A4
    c.saveState()
    c.setFillColor(COLOR_GREEN_DARK)
    c.setFont("Helvetica-Bold", 20)  # Hardcoded
    c.drawString(40, height-50, TITLE_TEXT)
    c.setFillColor(colors.black)
    c.setFont("Helvetica", 14)  # Hardcoded
    c.drawString(40, height-75, SUBTITLE_TEXT)
    c.setStrokeColor(COLOR_GREEN_DARK)
    c.setLineWidth(2)
    c.line(40, height-85, width-40, height-85)
    c.restoreState()

def draw_footer(c, doc):
    """Draws the footer with icons, contact info, and clickable links."""
    width, height = A4
    margin_bottom = 0
    FOOTER_HEIGHT = 60
    SIDE_MARGIN = 50
    ICON_WIDTH = 20
    ICON_HEIGHT = 20
    TEXT_ICON_PADDING = 5
    TEXT_LINE_HEIGHT = 12

    # Contact info from CONFIG
    contacts = [
        ("icons/phone.png", CONTACTS["phone"], "tel"),      # phones
        ("icons/email.png", [CONTACTS["email"]], "mailto"), # email
        ("icons/web.png", [CONTACTS["website"]], "url")     # website
    ]

    # Draw green background
    c.setFillColor(COLOR_GREEN_DARK)
    c.rect(0, margin_bottom, width, FOOTER_HEIGHT, fill=1, stroke=0)

    footer_center_y = margin_bottom + FOOTER_HEIGHT / 2

    for i, (icon_path, text_lines, link_type) in enumerate(contacts):
        # HORIZONTAL POSITIONING
        if i == 0:  # Left
            icon_x = SIDE_MARGIN
        elif i == 1:  # Center
            total_text_width = max(c.stringWidth(line, "Helvetica", 10) for line in text_lines)
            est_width = ICON_WIDTH + TEXT_ICON_PADDING + total_text_width
            icon_x = (width / 2) - (est_width / 2)
        else:  # Right
            total_text_width = max(c.stringWidth(line, "Helvetica", 10) for line in text_lines)
            total_block_width = ICON_WIDTH + TEXT_ICON_PADDING + total_text_width
            icon_x = width - SIDE_MARGIN - total_block_width

        # VERTICAL POSITIONING
        icon_y = footer_center_y - ICON_HEIGHT / 2

        # Draw icon if exists
        if os.path.exists(icon_path):
            c.drawImage(icon_path, icon_x, icon_y, width=ICON_WIDTH, height=ICON_HEIGHT,
                        preserveAspectRatio=True, mask='auto')
        else:
            c.setFillColor(colors.black)
            c.circle(icon_x + ICON_WIDTH / 2, icon_y + ICON_HEIGHT / 2, ICON_WIDTH / 2, fill=1, stroke=0)

        # Draw text
        c.setFont("Helvetica", 10)
        c.setFillColor(colors.white)

        num_lines = len(text_lines)
        total_block_height = TEXT_LINE_HEIGHT * num_lines
        text_start_y = footer_center_y + (total_block_height / 2) - TEXT_LINE_HEIGHT + 2
        text_x = icon_x + ICON_WIDTH + TEXT_ICON_PADDING

        for j, line in enumerate(text_lines):
            line_y = text_start_y - j * TEXT_LINE_HEIGHT
            c.drawString(text_x, line_y, line)

            # Add clickable link
            if link_type == "tel":
                link_url = f"tel:{line}"
            elif link_type == "mailto":
                link_url = f"mailto:{line}"
            else:  # url
                link_url = line if line.startswith("http") else f"https://{line}"

            text_width = c.stringWidth(line, "Helvetica", 10)
            text_height = TEXT_LINE_HEIGHT
            # linkRect(x1, y1, x2, y2, url)
            c.linkURL(link_url,
                      (text_x, line_y, text_x + text_width, line_y + text_height),
                      relative=0)

# -------------------------
# Generate PDF per student
# -------------------------
def generate_student_pdf(student_row, output_folder):
    alumno = student_row["Alumno/a"]
    certificacion = student_row["Certificación"]
    safe_name = sanitize_filename(alumno)
    output_filename = os.path.join(output_folder, f"{safe_name}.pdf")

    doc = SimpleDocTemplate(
        output_filename,
        pagesize=A4,
        rightMargin=40,
        leftMargin=40,
        topMargin=100,
        bottomMargin=70
    )

    styles = getSampleStyleSheet()
    elements = []

    # Condensed styles
    style_info_label = ParagraphStyle('InfoLabel', parent=styles['Normal'], fontSize=12, leading=13)
    style_info_data = ParagraphStyle('InfoData', parent=styles['Heading3'], fontSize=14, leading=16, spaceAfter=5)
    style_justified = ParagraphStyle('justify', parent=styles['Normal'], alignment=TA_JUSTIFY, leading=12, fontSize=11)

    # Intro Text
    if INTRO_TEXT != "" and INTRO_TEXT != None:
        elements.append(Paragraph(INTRO_TEXT, style_justified))
        elements.append(Spacer(1, 10))

    # Student Info Table
    data_info = [
        [Paragraph("<b>ALUMNO/A:</b>", style_info_label),
         Paragraph("<b>CERTIFICACIÓN:</b>", style_info_label)],
        [Paragraph(alumno, style_info_data),
         Paragraph(certificacion, style_info_data)]
    ]
    t_info = Table(data_info, colWidths=[300, 200])
    t_info.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('LEFTPADDING', (0,0), (-1,-1), 0),
        ('BOTTOMPADDING', (0,1), (-1,1), 12),
        ('LINEBELOW', (0,1), (-1,1), 1.2, colors.grey),
    ]))
    elements.append(t_info)
    elements.append(Spacer(1, 5))

    # Main Dashboard
    notes = [student_row[skill] for skill in SKILLS]
    avg_note = sum(notes) / len(notes)

    # ----- LEFT COLUMN -----
    left_content = [SectionHeader("CALIFICACIÓN INDIVIDUAL", width=240, height=25), Spacer(1, 8)]

    for skill, note in zip(SKILLS, notes):
        left_content.append(Paragraph(skill, ParagraphStyle('label', parent=styles['Normal'], leading=12, fontSize=11)))
        left_content.append(Spacer(1, 3))
        left_content.append(RoundedProgressBar(note, width=210))
        left_content.append(Spacer(1, 15))

    # ----- RIGHT COLUMN -----
    right_content = [
        SectionHeader("CALIFICACIÓN TOTAL", width=200, height=25),
        Indenter(10, 0),    
        SemiCircleGauge(avg_note),
        Spacer(1, 2),
        Indenter(-10, 0),
        SectionHeader("ESTADO", width=200, height=20),
        Paragraph(f"{'APTO/A' if avg_note >= PASSING_NOTE else 'NO APTO/A'}", ParagraphStyle('status', parent=styles['Heading3'], alignment=TA_CENTER, fontSize=16)),
    ]

    # Table to align both columns
    main_table_data = [[left_content, right_content]]
    main_table = Table(main_table_data, colWidths=[270, 230])
    main_table.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('LEFTPADDING', (1,0), (1,0), 20)
    ]))
    elements.append(main_table)
    elements.append(Spacer(1, 10))

    # Teacher Report
    if avg_note >= PASSING_NOTE and POSITIVE_CONCLUSION_TEXT != None and POSITIVE_CONCLUSION_TEXT != "":
        elements.append(Paragraph(POSITIVE_CONCLUSION_TEXT, style_justified))
    elif avg_note < PASSING_NOTE and NEGATIVE_CONCLUSION_TEXT != None and NEGATIVE_CONCLUSION_TEXT != "":
        elements.append(Paragraph(NEGATIVE_CONCLUSION_TEXT, style_justified))

    # Build PDF with header & footer
    doc.build(
        elements,
        onFirstPage=lambda c,d: (draw_header(c,d), draw_footer(c,d)),
        onLaterPages=lambda c,d: (draw_header(c,d), draw_footer(c,d))
    )

    print(f"Created: {output_filename}")

# -------------------------
# Main
# -------------------------
def main():
    parser = argparse.ArgumentParser(description="Generate one PDF per student with progress bars.")
    parser.add_argument("input_file", help="CSV or Excel input file")
    args = parser.parse_args()

    output_folder = "reports"
    os.makedirs(output_folder, exist_ok=True)

    try:
        df = read_input_file(args.input_file)
        for _, row in df.iterrows():
            generate_student_pdf(row, output_folder)
        print("\nAll student PDFs generated successfully!")
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    main()
