import fitz  # PyMuPDF

def extract_pdf_layout(pdf_path):
    doc = fitz.open(pdf_path)
    blocks = []

    for page_num, page in enumerate(doc):
        for block in page.get_text("blocks"):
            x0, y0, x1, y1, text, *_ = block
            blocks.append({
                "page": page_num,
                "bbox": [x0, y0, x1, y1],
                "text": text.strip()
            })

    return blocks
