from docx import Document
from zipfile import ZipFile
from lxml import etree
from docx.shared import Pt
from docx.oxml.ns import qn

# --------------------------------------------------
# Load theme fonts (major / minor)
# --------------------------------------------------
def extract_theme_fonts(docx_path):
    with ZipFile(docx_path) as z:
        theme_xml = z.read("word/theme/theme1.xml")

    root = etree.fromstring(theme_xml)
    ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}

    major = root.find(".//a:majorFont/a:latin", ns)
    minor = root.find(".//a:minorFont/a:latin", ns)

    def print_document_xml(docx_path):
        with ZipFile(docx_path) as z:
            xml_bytes = z.read("word/document.xml")
            xml_str = xml_bytes.decode("utf-8")
            print(xml_str)

    return {
        "major": major.get("typeface") if major is not None else None,
        "minor": minor.get("typeface") if minor is not None else None
    }


# --------------------------------------------------
# Resolve font name (run → style → theme)
# --------------------------------------------------
def resolve_font_name(run, paragraph, theme_fonts):
    # 1️⃣ Run-level
    if run.font.name:
        return run.font.name

    # 2️⃣ Paragraph style
    style = paragraph.style
    if style and style.font and style.font.name:
        return style.font.name

    # 3️⃣ Heading vs body → theme
    if style and style.name.lower().startswith("heading"):
        return theme_fonts["major"]

    return theme_fonts["minor"]


# --------------------------------------------------
# Resolve font size (run → style → defaults)
# --------------------------------------------------
def resolve_font_size(run, paragraph):
    # 1️⃣ Run-level
    if run.font.size:
        return run.font.size.pt

    # 2️⃣ Style-level
    style = paragraph.style
    if style and style.font and style.font.size:
        return style.font.size.pt

    # 3️⃣ Word default (Normal ≈ 11pt)
    return 11.0


# --------------------------------------------------
# Extract highlight / background
# --------------------------------------------------
def extract_background(run):
    rPr = run._element.rPr
    if rPr is None:
        return {
            "highlight": None,
            "shading": None
        }

    # Highlight
    highlight = run.font.highlight_color
    highlight_val = highlight._value if highlight else None

    # Shading (safe XML lookup)
    shd = rPr.find(qn("w:shd"))

    shading_val = shd.get(qn("w:fill")) if shd is not None else None

    return {
        "highlight": highlight_val,
        "shading": shading_val
    }

# --------------------------------------------------
# MAIN EXTRACTOR
# --------------------------------------------------
def extract_docx_styles(docx_path):
    doc = Document(docx_path)
    theme_fonts = extract_theme_fonts(docx_path)
    elements = []

    for para in doc.paragraphs:
        para_data = {
            "text": para.text,
            "style": para.style.name if para.style else None,
            "alignment": para.alignment,
            "space_before": para.paragraph_format.space_before.pt if para.paragraph_format.space_before else None,
            "space_after": para.paragraph_format.space_after.pt if para.paragraph_format.space_after else None,
            "line_spacing": para.paragraph_format.line_spacing,
            "runs": []
        }

        for run in para.runs:
            font_name = resolve_font_name(run, para, theme_fonts)
            font_size = resolve_font_size(run, para)

            para_data["runs"].append({
                "text": run.text,
                "font": font_name,
                "font_size_pt": font_size,
                "bold": run.bold,
                "italic": run.italic,
                "background": extract_background(run)
            })

        elements.append(para_data)

    return elements
