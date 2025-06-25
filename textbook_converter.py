# [all imports]
import os
import re
import fitz
import tempfile
from pptx import Presentation
from PIL import Image
from reportlab.platypus import SimpleDocTemplate, Paragraph, Image as RLImage, Spacer, Preformatted, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.pagesizes import A4
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.units import inch
from reportlab.lib import colors
import streamlit as st
import google.generativeai as genai
import time

# =========================
# GEMINI CONFIGURATION
# =========================
genai.configure(api_key="AIzaSyBQC4WS3nxPaneINL3V1nuVsZd4QMa5Ra4")
model = genai.GenerativeModel("gemini-pro")

# ================
# CHUNKING UTILITY
# ================
def chunk_text(text, max_len=1500):
    words = text.split()
    chunks = []
    chunk = []
    for word in words:
        chunk.append(word)
        if len(" ".join(chunk)) >= max_len:
            chunks.append(" ".join(chunk))
            chunk = []
    if chunk:
        chunks.append(" ".join(chunk))
    return chunks

# ====================
# GEMINI PROMPT LOGIC
# ====================
def call_gemini_prompt(raw_text, retries=3):
    prompt = f"""
You are a professional academic typesetter.

## Format Instructions:
- Use `#`, `##`, `###` for heading levels: topics, subtopics, and sub-subtopics.
- Maintain numbered sections like `1 INTRODUCTION`, `1.1 Motivation`, etc.
- Keep paragraph breaks.
- Wrap code with triple backticks.
- Use Markdown tables where applicable.
- Output only Markdown, nothing else.

## Input Text:
\"\"\"
{raw_text}
\"\"\"
"""
    for attempt in range(retries):
        try:
            response = model.generate_content(prompt)
            return response.text
        except Exception as e:
            print(f"[Gemini Retry {attempt+1}] Error: {e}")
            time.sleep(1 + attempt)
    return raw_text

def clean_text(raw_text):
    chunks = chunk_text(raw_text)
    return "\n\n".join(call_gemini_prompt(chunk) for chunk in chunks)

# =========================
# STYLES
# =========================
def create_styles(font_family='Times-Roman'):
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name='TextbookTitle', fontName=font_family, fontSize=20, alignment=TA_CENTER, spaceAfter=20))
    styles.add(ParagraphStyle(name='TextbookHeading1', fontName=font_family, fontSize=16, spaceAfter=12, spaceBefore=12))
    styles.add(ParagraphStyle(name='TextbookHeading2', fontName=font_family, fontSize=14, spaceAfter=10, spaceBefore=10))
    styles.add(ParagraphStyle(name='TextbookHeading3', fontName=font_family, fontSize=12, spaceAfter=8, spaceBefore=8))
    styles.add(ParagraphStyle(name='TextbookBody', fontName=font_family, fontSize=11, spaceAfter=6))
    styles.add(ParagraphStyle(name='TextbookListItem', fontName=font_family, fontSize=11, leftIndent=20, spaceAfter=4))
    styles.add(ParagraphStyle(name='TextbookItalic', fontName=font_family, fontSize=11, spaceAfter=6, italic=True))
    styles.add(ParagraphStyle(name='TextbookCode', fontName='Courier', fontSize=10, leading=12, spaceAfter=8))
    return styles

# =========================
# MARKDOWN PARSER W/ TABLE SUPPORT
# =========================
def parse_markdown_to_flowables(text, styles):
    flowables = []
    in_code_block = False
    code_lines = []
    table_lines = []
    in_table = False

    for line in text.split("\n"):
        line = line.rstrip()

        # Code blocks
        if line.strip().startswith("```"):
            if not in_code_block:
                in_code_block = True
                code_lines = []
            else:
                in_code_block = False
                flowables.append(Preformatted("\n".join(code_lines), styles["TextbookCode"]))
        elif in_code_block:
            code_lines.append(line)
        # Table detection
        elif "|" in line and re.match(r"^\|", line.strip()):
            in_table = True
            table_lines.append(line)
        elif in_table and (not line.strip() or "---" not in line):
            # Parse and end table
            data = [re.split(r"\s*\|\s*", row.strip("| ")) for row in table_lines if row.strip()]
            table = Table(data, hAlign="LEFT")
            table.setStyle(TableStyle([
                ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
                ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
                ('FONTNAME', (0,0), (-1,-1), 'Helvetica'),
                ('FONTSIZE', (0,0), (-1,-1), 10),
                ('ALIGN', (0,0), (-1,-1), 'LEFT'),
            ]))
            flowables.append(table)
            flowables.append(Spacer(1, 12))
            in_table = False
            table_lines = []
            if line.strip():
                flowables.append(Paragraph(line, styles["TextbookBody"]))
        elif in_table:
            table_lines.append(line)
        # Headings and body
        elif not line.strip():
            flowables.append(Spacer(1, 12))
        elif line.startswith("### "):
            flowables.append(Paragraph(line[4:], styles["TextbookHeading3"]))
        elif line.startswith("## "):
            flowables.append(Paragraph(line[3:], styles["TextbookHeading2"]))
        elif line.startswith("# "):
            flowables.append(Paragraph(line[2:], styles["TextbookHeading1"]))
        elif line.startswith("- "):
            flowables.append(Paragraph(f"• {line[2:]}", styles["TextbookListItem"]))
        elif re.match(r"^\d+\.", line):
            flowables.append(Paragraph(line, styles["TextbookListItem"]))
        elif line.startswith("**Note:**") or line.startswith("**Example:**"):
            flowables.append(Paragraph(f"<i>{line}</i>", styles["TextbookItalic"]))
        else:
            flowables.append(Paragraph(line, styles["TextbookBody"]))
    return flowables

# =====================
# FILE EXTRACTION UTILS
# =====================
def extract_pptx(pptx_path, img_dir, progress_bar=None, progress_text=None):
    prs = Presentation(pptx_path)
    content = []
    for i, slide in enumerate(prs.slides):
        slide_text = []
        images = []
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                slide_text.append(shape.text)
            if shape.shape_type == 13:
                img = shape.image
                img_bytes = img.blob
                img_path = os.path.join(img_dir, f"slide_{i}.png")
                with open(img_path, "wb") as f:
                    f.write(img_bytes)
                if os.path.getsize(img_path) > 20 * 1024:
                    images.append(img_path)
        content.append({"text": "\n".join(slide_text), "images": images})
        if progress_bar:
            progress_bar.progress((i + 1) / len(prs.slides))
    return content

def extract_pdf(pdf_path, img_dir, progress_bar=None, progress_text=None):
    doc = fitz.open(pdf_path)
    content = []
    for i, page in enumerate(doc):
        text = page.get_text()
        images = []
        for img_index, img in enumerate(page.get_images(full=True)):
            base_image = doc.extract_image(img[0])
            image_bytes = base_image["image"]
            img_path = os.path.join(img_dir, f"page_{i}_{img_index}.png")
            with open(img_path, "wb") as f:
                f.write(image_bytes)
            if os.path.getsize(img_path) > 20 * 1024:
                images.append(img_path)
        content.append({"text": text, "images": images})
        if progress_bar:
            progress_bar.progress((i + 1) / len(doc))
    return content

# =====================
# PDF CREATION FUNCTION
# =====================
def create_textbook_pdf(content, output_path, font_family='Times-Roman', progress_bar=None, progress_text=None):
    styles = create_styles(font_family)
    doc = SimpleDocTemplate(output_path, pagesize=A4, rightMargin=72, leftMargin=72, topMargin=72, bottomMargin=72)
    story = []

    for i, block in enumerate(content):
        raw_text = block.get("text", "")
        structured_text = clean_text(raw_text)
        flowables = parse_markdown_to_flowables(structured_text, styles)
        story.extend(flowables)

        for img_path in block.get("images", []):
            story.append(RLImage(img_path, width=4 * inch, height=3 * inch))
            story.append(Spacer(1, 12))

        if progress_bar:
            progress_bar.progress((i + 1) / len(content))

    doc.build(story)

# =====================
# STREAMLIT WEB INTERFACE
# =====================
def main():
    st.set_page_config(page_title="Textbook Converter", layout="centered")
    st.title("📘 Lecture-to-Textbook Converter")

    uploaded_file = st.file_uploader("Upload a PDF or PPTX lecture file", type=["pdf", "pptx"])
    font_choice = st.selectbox("Choose a font", ["Times-Roman", "Helvetica", "Courier"])

    if uploaded_file and st.button("Convert to Textbook"):
        with tempfile.TemporaryDirectory() as tmpdir:
            input_path = os.path.join(tmpdir, uploaded_file.name)
            with open(input_path, "wb") as f:
                f.write(uploaded_file.read())

            img_dir = os.path.join(tmpdir, "images")
            os.makedirs(img_dir, exist_ok=True)

            progress_bar = st.progress(0)
            progress_text = st.empty()

            if uploaded_file.name.endswith(".pdf"):
                content = extract_pdf(input_path, img_dir, progress_bar, progress_text)
            else:
                content = extract_pptx(input_path, img_dir, progress_bar, progress_text)

            output_path = os.path.join(tmpdir, "output_textbook.pdf")
            create_textbook_pdf(content, output_path, font_choice, progress_bar, progress_text)

            st.success("✅ Textbook PDF generated successfully!")
            with open(output_path, "rb") as f:
                st.download_button("📥 Download Textbook PDF", f, file_name="textbook.pdf", mime="application/pdf")

if __name__ == "__main__":
    main()
