from parse_doc import extract_docx_styles
from docx_pdf import docx_to_pdf
from extract_layout import extract_pdf_layout
from reconcile import reconcile
import json
import os

def build_style_layout_template(docx_path, workdir):
    styles = extract_docx_styles(docx_path)
    pdf = docx_to_pdf(docx_path, workdir)
    layout = extract_pdf_layout(pdf)
    template = reconcile(styles, layout)
    return template

if __name__ == "__main__":
    input_docx = "./2026_01_Precision_AI_UFA_Template.docx"
    working_dir = "./tmp"
    os.makedirs(working_dir, exist_ok=True)

    template = build_style_layout_template(input_docx, working_dir)

    output_file = "layout_template.json"
    with open(output_file, "w", encoding="utf-8") as f:
        json.dump(template, f, indent=2, ensure_ascii=False)

    print(f"Template saved to {output_file}")
