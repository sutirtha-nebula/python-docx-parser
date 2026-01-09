from docx import Document

def extract_docx_styles(docx_path):
    doc = Document(docx_path)
    elements = []

    for para in doc.paragraphs:
        elements.append({
            "text": para.text,
            "style": para.style.name,
            "alignment": para.alignment,
            "space_before": para.paragraph_format.space_before,
            "space_after": para.paragraph_format.space_after,
            "line_spacing": para.paragraph_format.line_spacing,
            "runs": [
                {
                    "text": run.text,
                    "font": run.font.name,
                    "size": run.font.size,
                    "bold": run.bold,
                    "italic": run.italic
                }
                for run in para.runs
            ]
        })

    return elements
