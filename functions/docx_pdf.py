import subprocess
import os

def docx_to_pdf(docx_path, output_dir):
    subprocess.run([
        "libreoffice",
        "--headless",
        "--convert-to", "pdf",
        "--outdir", output_dir,
        docx_path
    ], check=True)

    pdf_path = os.path.join(
        output_dir,
        os.path.basename(docx_path).replace(".docx", ".pdf")
    )

    return pdf_path
