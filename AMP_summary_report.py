from pathlib import Path
import re
import html
from turtle import title
import pandas as pd

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate,
    LongTable,
    TableStyle,
    Paragraph,
    Spacer,
)
from reportlab.pdfgen import canvas as canvas_module
from pypdf import PdfReader, PdfWriter


# ============================================================
# FILE SETTINGS
# ============================================================

BASE_DIR = Path(r"C:\Users\PatrickBaecher\Documents\Python\AMOS")

EXCEL_FILE = BASE_DIR / "OMP_summary_report_20260612.xlsx"
LOGO_FILE = BASE_DIR / "ASLB.svg.png"

OUTPUT_DIR = BASE_DIR / "output"
CHAPTER_DIR = OUTPUT_DIR / "chapters"

FINAL_PDF = OUTPUT_DIR / "OMP_Maintenance_Program_Summary.pdf"


# ============================================================
# CHAPTER RULES
# Order matters: first match wins.
# Example:
# "PART 1A; SECTION 1" will be assigned to "PART 1A"
# ============================================================

CHAPTER_RULES = [
    ("01", "PART 1A", "PART 1A"),
    ("02", "PART 1B", "PART 1B"),
    ("03", "PART 1C", "PART 1C"),
    ("04", "PART 1D", "PART 1D"),
    ("05", "PART 1E", "PART 1E"),
    ("06", "PART 2", "PART 2"),
    ("07", "PART 3", "PART 3"),
]


# ============================================================
# PAGE / LAYOUT SETTINGS
# ============================================================

PAGE_SIZE = landscape(A4)

LEFT_MARGIN = 8 * mm
RIGHT_MARGIN = 8 * mm
TOP_MARGIN = 35 * mm
BOTTOM_MARGIN = 10 * mm

# Available width on landscape A4 with the above margins:
# 842 pt - margins = approx. 785 pt
COLUMN_WIDTHS = [
    55,   # Rev Task
    90,   # Taskcard
    70,   # Zone(s)/MSG
    80,   # Frequency
    120,  # Reference(s)
    140,  # Effectivity
    230,  # Task Description
]


# ============================================================
# TEXT STYLES
# ============================================================

STYLE_CELL = ParagraphStyle(
    name="Cell",
    fontName="Helvetica",
    fontSize=6.2,
    leading=7.4,
    alignment=TA_LEFT,
    spaceAfter=0,
)

STYLE_CELL_BOLD = ParagraphStyle(
    name="CellBold",
    fontName="Helvetica-Bold",
    fontSize=6.2,
    leading=7.4,
    alignment=TA_LEFT,
    spaceAfter=0,
)

STYLE_HEADER = ParagraphStyle(
    name="TableHeader",
    fontName="Helvetica-Bold",
    fontSize=7,
    leading=8,
    alignment=TA_CENTER,
)

STYLE_CHAPTER_TITLE = ParagraphStyle(
    name="ChapterTitle",
    fontName="Helvetica-Bold",
    fontSize=14,
    leading=16,
    alignment=TA_CENTER,
)


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def clean_value(value) -> str:
    """Convert Excel cell value into clean text."""
    if pd.isna(value):
        return ""

    text = str(value)

    # Convert common HTML breaks from AMOS text into line breaks
    text = re.sub(r"<br\s*/?>", "\n", text, flags=re.IGNORECASE)

    # Remove other basic HTML tags if present
    text = re.sub(r"<[^>]+>", "", text)

    # Decode HTML entities such as &nbsp;
    text = html.unescape(text)

    # Normalize non-breaking spaces
    text = text.replace("\xa0", " ")

    # Normalize line endings
    text = text.replace("\r\n", "\n").replace("\r", "\n")

    # Remove excessive empty lines
    text = re.sub(r"\n{3,}", "\n\n", text)

    return text.strip()


def list_break(value) -> str:
    """
    Convert semicolon-separated list values into line breaks.
    Use only for fields where semicolon means list separator.
    """
    text = clean_value(value)
    if not text:
        return ""
    return re.sub(r"\s*;\s*", "\n", text)

def task_description_paragraph(title, workstep_text) -> Paragraph:
    """
    Create Task Description cell:
    - title in bold
    - empty line
    - workstep_text normal
    """
    title_text = clean_value(title)
    workstep = clean_value(workstep_text)

    parts = []

    if title_text:
        title_text = html.escape(title_text).replace("\n", "<br/>")
        parts.append(f"<b>{title_text}</b>")

    if workstep:
        workstep = html.escape(workstep).replace("\n", "<br/>")

        if parts:
            parts.append("<br/><br/>" + workstep)
        else:
            parts.append(workstep)

    return Paragraph("".join(parts), STYLE_CELL)


def paragraph(text: str, style=STYLE_CELL) -> Paragraph:
    """Create ReportLab Paragraph with line breaks preserved."""
    text = clean_value(text)

    # Escape XML-sensitive characters for ReportLab Paragraph
    text = html.escape(text)

    # Convert line breaks to ReportLab line breaks
    text = text.replace("\n", "<br/>")

    return Paragraph(text, style)

def references_paragraph(mrbno, mpdno, reference, combined_special) -> Paragraph:
    """
    Create Reference(s) cell:
    - MRB: in bold + value from mrbno
    - MPD: in bold + value from mpdno
    - General References: in bold, value underneath
    - Special Code: in bold, value underneath
    """
    mrb = clean_value(mrbno)
    mpd = clean_value(mpdno)
    general_refs = list_break(reference)
    special_code = list_break(combined_special)

    parts = []

    # 1. MRB
    parts.append(f"<b>MRB:</b> {html.escape(mrb)}")

    # 2. MPD
    parts.append(f"<b>MPD:</b> {html.escape(mpd)}")

    # 3. General References
    if general_refs:
        general_refs = html.escape(general_refs).replace("\n", "<br/>")
        parts.append(f"<b>General References:</b><br/>{general_refs}")
    else:
        parts.append("<b>General References:</b>")

    # 4. Special Code
    if special_code:
        special_code = html.escape(special_code).replace("\n", "<br/>")
        parts.append(f"<b>Special Code:</b><br/>{special_code}")
    else:
        parts.append("<b>Special Code:</b>")

    return Paragraph("<br/>".join(parts), STYLE_CELL)

def combine_lines(*lines) -> str:
    """Join non-empty lines with line breaks."""
    cleaned = [clean_value(x) for x in lines if clean_value(x)]
    return "\n".join(cleaned)


def get_chapter_info(section_code: str) -> tuple[str, str]:
    """
    Return chapter_sort and chapter_title based on section_code.
    """
    value = clean_value(section_code).upper()

    for sort_key, match_text, display_title in CHAPTER_RULES:
        if match_text == "OTHER":
            continue
        if match_text.upper() in value:
            return sort_key, display_title

    return "99", "OTHER"


def safe_first_value(df: pd.DataFrame, column_name: str) -> str:
    """Return first non-empty value from a column."""
    if column_name not in df.columns:
        return ""

    values = df[column_name].dropna().astype(str)
    values = [v.strip() for v in values if v.strip()]

    return values[0] if values else ""


# ============================================================
# PAGE NUMBER CANVAS
# This allows "Page x / y" inside each chapter PDF.
# ============================================================

class NumberedCanvas(canvas_module.Canvas):
    def __init__(self, *args, chapter_title="", mpno="", revision="", logo_file=None, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []
        self.chapter_title = chapter_title
        self.mpno = mpno
        self.revision = revision
        self.logo_file = logo_file

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        page_count = len(self._saved_page_states)

        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_header_footer(page_count)
            super().showPage()

        super().save()

    def draw_header_footer(self, page_count):
        page_width, page_height = PAGE_SIZE

        # Outer border
        self.setLineWidth(1)
        self.rect(
            6 * mm,
            6 * mm,
            page_width - 12 * mm,
            page_height - 12 * mm,
        )

        # Header border line
        self.setLineWidth(0.8)
        self.line(6 * mm, page_height - 27 * mm, page_width - 6 * mm, page_height - 27 * mm)

        # Logo
        if self.logo_file and Path(self.logo_file).exists():
            try:
                self.drawImage(
                    str(self.logo_file),
                    9 * mm,
                    page_height - 24 * mm,
                    width=38 * mm,
                    height=16 * mm,
                    preserveAspectRatio=True,
                    mask="auto",
                )
            except Exception:
                # If logo fails, continue without breaking PDF generation
                pass

        # Center title = chapter title
        self.setFont("Helvetica-Bold", 16)
        self.drawCentredString(
            page_width / 2,
            page_height - 16 * mm,
            self.chapter_title,
        )

        # Right page number
        self.setFont("Helvetica-Bold", 8)
        self.drawRightString(
            page_width - 10 * mm,
            page_height - 12 * mm,
            f"Page {self._pageNumber} / {page_count}",
        )

        # Sub-header info
        self.setFont("Helvetica-Bold", 9)
        header_info = (
            f"Maintenance Program: {self.mpno}    "
            f"Issue/Revision: {self.revision}"
        )
        self.drawString(
            9 * mm,
            page_height - 32 * mm,
            header_info,
        )

        # Footer
        self.setFont("Helvetica", 6)
        self.drawRightString(
            page_width - 8 * mm,
            4 * mm,
            "Generated by Python prototype",
        )


# ============================================================
# TABLE ROW BUILDING
# ============================================================

def build_report_row(row) -> list:
    """
    Build one PDF table row from one Excel row.
    """

    # Column 1: Rev Task
    rev_task = combine_lines(
        row.get("omp_revision_no", ""),
        row.get("section_code", ""),
    )

    # Column 2: Taskcard
    taskcard_lines = [
        row.get("taskcardno", ""),
        row.get("single_running", ""),
    ]

    linked_check = clean_value(row.get("linked_check", ""))
    if linked_check:
        taskcard_lines.append(f"Linked Check: {linked_check}")

    task_type_code = clean_value(row.get("task_type_code", ""))
    if task_type_code:
        taskcard_lines.append(f"Type Code: {task_type_code}")

    taskcard = combine_lines(*taskcard_lines)

    # Column 3: Zone(s)/MSG
    zone = list_break(row.get("zone", ""))
    msg_code = clean_value(row.get("msg_code", ""))

    zone_msg = zone
    if msg_code:
        zone_msg = combine_lines(zone_msg, f"MSG: {msg_code}")

    # Column 4: Frequency
    thr = clean_value(row.get("thr", ""))
    interval = clean_value(row.get("INT", ""))

    if thr == "T:":
        thr = ""
    if interval == "I:":
        interval = ""

    frequency = combine_lines(thr, interval)

    # Column 5: Reference(s)
    references = references_paragraph(
        mrbno=row.get("mrbno", ""),
        mpdno=row.get("mpdno", ""),
        reference=row.get("reference", ""),
        combined_special=row.get("combined_special", ""),
)

    # Column 6: Effectivity
    applicability = list_break(row.get("applicability", ""))
    eff_notes = clean_value(row.get("eff_notes", ""))

    effectivity = combine_lines(applicability, eff_notes)

    # Column 7: Task Description
    title = row.get("title", "")
    workstep_text = row.get("workstep_text", "")

    task_description = task_description_paragraph(title, workstep_text)

    return [
        paragraph(rev_task),
        paragraph(taskcard),
        paragraph(zone_msg),
        paragraph(frequency),
        references,
        paragraph(effectivity),
        task_description,
    ]


# ============================================================
# PDF GENERATION
# ============================================================

def generate_chapter_pdf(chapter_df: pd.DataFrame, chapter_title: str, output_pdf: Path):
    """
    Generate one PDF for one chapter/section.
    """
    mpno = safe_first_value(chapter_df, "omp_mpno")
    revision = safe_first_value(chapter_df, "omp_revision_no")

    doc = SimpleDocTemplate(
        str(output_pdf),
        pagesize=PAGE_SIZE,
        leftMargin=LEFT_MARGIN,
        rightMargin=RIGHT_MARGIN,
        topMargin=TOP_MARGIN,
        bottomMargin=BOTTOM_MARGIN,
    )

    story = []

    table_header = [
        paragraph("Rev Task", STYLE_HEADER),
        paragraph("Taskcard", STYLE_HEADER),
        paragraph("Zone(s)/MSG", STYLE_HEADER),
        paragraph("Frequency", STYLE_HEADER),
        paragraph("Reference(s)", STYLE_HEADER),
        paragraph("Effectivity", STYLE_HEADER),
        paragraph("Task Description", STYLE_HEADER),
    ]

    table_data = [table_header]

    for _, row in chapter_df.iterrows():
        table_data.append(build_report_row(row))

    table = LongTable(
        table_data,
        colWidths=COLUMN_WIDTHS,
        repeatRows=1,
        splitByRow=1,
    )

    table.setStyle(TableStyle([
        # Header row
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("ALIGN", (0, 0), (-1, 0), "CENTER"),

        # Grid
        ("GRID", (0, 0), (-1, -1), 0.35, colors.black),

        # Body cells
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 2),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))

    story.append(table)

    doc.build(
        story,
        canvasmaker=lambda *args, **kwargs: NumberedCanvas(
            *args,
            chapter_title=chapter_title,
            mpno=mpno,
            revision=revision,
            logo_file=LOGO_FILE,
            **kwargs,
        )
    )


def merge_pdfs(chapter_files: list[Path], final_pdf: Path):
    """
    Merge all chapter PDFs into one final PDF using PdfWriter.
    """
    writer = PdfWriter()

    for pdf_file in chapter_files:
        reader = PdfReader(str(pdf_file))

        for page in reader.pages:
            writer.add_page(page)

    with open(final_pdf, "wb") as output_file:
        writer.write(output_file)


def main():
    OUTPUT_DIR.mkdir(exist_ok=True)
    CHAPTER_DIR.mkdir(exist_ok=True)

    if not EXCEL_FILE.exists():
        raise FileNotFoundError(f"Excel file not found: {EXCEL_FILE}")

    print(f"Reading Excel file: {EXCEL_FILE}")

    df = pd.read_excel(EXCEL_FILE, dtype=str)
    df = df.fillna("")

    required_columns = [
        "omp_mpno",
        "omp_issue_no",
        "omp_revision_no",
        "section_code",
        "reference",
        "taskcardno",
        "single_running",
        "linked_check",
        "mpdno",
        "mrbno",
        "task_type_code",
        "msg_code",
        "zone",
        "combined_special",
        "title",
        "workstep_text",
        "eff_notes",
        "applicability",
        "linked_partno",
        "thr",
        "INT",
    ]

    missing = [col for col in required_columns if col not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns in Excel file: {missing}")

    # Assign chapter key and title from section_code
    chapter_info = df["section_code"].apply(get_chapter_info)
    df["chapter_sort"] = chapter_info.apply(lambda x: x[0])
    df["chapter_title"] = chapter_info.apply(lambda x: x[1])

    # Sort report
    df = df.sort_values(
        by=["chapter_sort", "taskcardno"],
        ascending=[True, True],
        kind="stable",
    )

    chapter_files = []

    for (chapter_sort, chapter_title), chapter_df in df.groupby(["chapter_sort", "chapter_title"], sort=True):
        safe_title = re.sub(r"[^A-Za-z0-9_-]+", "_", chapter_title)
        chapter_pdf = CHAPTER_DIR / f"{chapter_sort}_{safe_title}.pdf"

        print(f"Creating chapter PDF: {chapter_pdf.name} ({len(chapter_df)} rows)")

        generate_chapter_pdf(
            chapter_df=chapter_df,
            chapter_title=chapter_title,
            output_pdf=chapter_pdf,
        )

        chapter_files.append(chapter_pdf)

    print("Merging chapter PDFs...")
    merge_pdfs(chapter_files, FINAL_PDF)

    print(f"Done: {FINAL_PDF}")


if __name__ == "__main__":
    main()
