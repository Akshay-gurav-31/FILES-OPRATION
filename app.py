import os
import uuid
from flask import Flask, request, render_template, send_file, jsonify
from werkzeug.utils import secure_filename

# PDF processing
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.colors import HexColor
import PyPDF2
import io

# DOCX processing
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"
ALLOWED_EXTENSIONS = {"pdf", "docx"}

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def unique_filename(folder, original_name):
    """Generate a unique filename to avoid overwriting existing files."""
    name, ext = os.path.splitext(original_name)
    unique = f"{name}_{uuid.uuid4().hex[:8]}{ext}"
    return os.path.join(folder, unique)


# ─── PDF Processing ───────────────────────────────────────────────────────────

def add_footer_to_pdf(input_path, output_path, footer_text):
    """
    Overlay a footer on every page of the PDF using ReportLab + PyPDF2.
    Strategy: create a one-page 'footer stamp' PDF in memory for each page,
    then merge it onto the original page.
    """
    reader = PyPDF2.PdfReader(input_path)
    writer = PyPDF2.PdfWriter()
    total_pages = len(reader.pages)

    for page_num, page in enumerate(reader.pages, start=1):
        # Get page dimensions
        media_box = page.mediabox
        page_width = float(media_box.width)
        page_height = float(media_box.height)

        # Build the footer stamp in memory
        packet = io.BytesIO()
        c = canvas.Canvas(packet, pagesize=(page_width, page_height))

        # Compose the footer line
        full_footer = f"{footer_text}    |    Page {page_num} of {total_pages}"

        # Draw a thin separator line just above the footer text
        c.setStrokeColor(HexColor("#AAAAAA"))
        c.setLineWidth(0.5)
        margin = 50
        y_line = 38
        c.line(margin, y_line, page_width - margin, y_line)

        # Footer text
        c.setFont("Helvetica", 9)
        c.setFillColor(HexColor("#555555"))
        c.drawCentredString(page_width / 2, 22, full_footer)

        c.save()
        packet.seek(0)

        # Stamp the footer onto the original page
        stamp_reader = PyPDF2.PdfReader(packet)
        stamp_page = stamp_reader.pages[0]
        page.merge_page(stamp_page)
        writer.add_page(page)

    with open(output_path, "wb") as f:
        writer.write(f)


# ─── DOCX Processing ──────────────────────────────────────────────────────────

def add_footer_to_docx(input_path, output_path, footer_text):
    """
    Add (or replace) the footer in every section of the DOCX file.
    Includes a page-number field and the custom footer text.
    """
    doc = Document(input_path)

    for section in doc.sections:
        section.different_first_page_header_footer = False

        footer = section.footer
        # Clear any existing footer content
        for para in footer.paragraphs:
            for run in para.runs:
                run.text = ""

        # Use the first paragraph (always present) or add one
        if footer.paragraphs:
            para = footer.paragraphs[0]
        else:
            para = footer.add_paragraph()

        para.clear()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Add custom footer text
        run_text = para.add_run(f"{footer_text}  |  Page ")
        run_text.font.size = Pt(9)
        run_text.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

        # Insert PAGE field (current page number)
        _add_page_number_field(run_text)

        run_of = para.add_run(" of ")
        run_of.font.size = Pt(9)
        run_of.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

        # Insert NUMPAGES field (total pages)
        _add_num_pages_field(para)

    doc.save(output_path)


def _add_page_number_field(run):
    """Insert a PAGE field inside the given run element."""
    fld_char1 = OxmlElement("w:fldChar")
    fld_char1.set(qn("w:fldCharType"), "begin")

    instr_text = OxmlElement("w:instrText")
    instr_text.set(qn("xml:space"), "preserve")
    instr_text.text = "PAGE"

    fld_char2 = OxmlElement("w:fldChar")
    fld_char2.set(qn("w:fldCharType"), "end")

    run._r.append(fld_char1)
    run._r.append(instr_text)
    run._r.append(fld_char2)


def _add_num_pages_field(para):
    """Insert a NUMPAGES field as a new run inside para."""
    run = para.add_run()
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

    fld_char1 = OxmlElement("w:fldChar")
    fld_char1.set(qn("w:fldCharType"), "begin")

    instr_text = OxmlElement("w:instrText")
    instr_text.set(qn("xml:space"), "preserve")
    instr_text.text = "NUMPAGES"

    fld_char2 = OxmlElement("w:fldChar")
    fld_char2.set(qn("w:fldCharType"), "end")

    run._r.append(fld_char1)
    run._r.append(instr_text)
    run._r.append(fld_char2)


# ─── Routes ───────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/process", methods=["POST"])
def process():
    # Validate file presence
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded."}), 400

    file = request.files["file"]
    footer_text = request.form.get("footer_text", "").strip()

    if file.filename == "":
        return jsonify({"error": "No file selected."}), 400

    if not footer_text:
        return jsonify({"error": "Footer text cannot be empty."}), 400

    if not allowed_file(file.filename):
        return jsonify({"error": "Only PDF and DOCX files are supported."}), 400

    # Save uploaded file with a safe, unique name
    original_name = secure_filename(file.filename)
    upload_path = unique_filename(UPLOAD_FOLDER, original_name)
    file.save(upload_path)

    ext = original_name.rsplit(".", 1)[1].lower()
    output_name = original_name.replace(f".{ext}", f"_with_footer.{ext}")
    output_path = unique_filename(OUTPUT_FOLDER, output_name)

    try:
        if ext == "pdf":
            add_footer_to_pdf(upload_path, output_path, footer_text)
        elif ext == "docx":
            add_footer_to_docx(upload_path, output_path, footer_text)
    except Exception as e:
        return jsonify({"error": f"Processing failed: {str(e)}"}), 500

    return send_file(
        output_path,
        as_attachment=True,
        download_name=output_name,
        mimetype=(
            "application/pdf" if ext == "pdf"
            else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        ),
    )


if __name__ == "__main__":
    app.run(debug=True)
