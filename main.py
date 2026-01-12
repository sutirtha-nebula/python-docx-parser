from functions.parse_doc import extract_docx_styles
from functions.docx_pdf import docx_to_pdf
from functions.extract_layout import extract_pdf_layout
from functions.reconcile import reconcile
import json
import os


def build_style_layout_template(docx_path, workdir):
    os.makedirs(workdir, exist_ok=True)

    # 1. Extract DOCX styles
    styles = extract_docx_styles(docx_path)

    styles_path = os.path.join(workdir, "styles.json")
    with open(styles_path, "w", encoding="utf-8") as f:
        json.dump(styles, f, indent=2, ensure_ascii=False)

    # 2. Convert DOCX → PDF
    pdf_path = docx_to_pdf(docx_path, workdir)

    # 3. Extract PDF layout
    layout = extract_pdf_layout(pdf_path)

    layout_path = os.path.join(workdir, "layout.json")
    with open(layout_path, "w", encoding="utf-8") as f:
        json.dump(layout, f, indent=2, ensure_ascii=False)

    # 4. Reconcile styles + layout
    template = reconcile(styles, layout)

    template_path = os.path.join(workdir, "template.json")
    with open(template_path, "w", encoding="utf-8") as f:
        json.dump(template, f, indent=2, ensure_ascii=False)

    return {
        "styles": styles_path,
        "layout": layout_path,
        "template": template_path
    }


if __name__ == "__main__":
    input_docx = "./data/2026_01_Precision_AI_UFA_Template.docx"
    working_dir = "./tmp"

    result = build_style_layout_template(input_docx, working_dir)

    print("Artifacts generated:")
    for k, v in result.items():
        print(f" - {k}: {v}")