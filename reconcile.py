def reconcile(docx_elements, pdf_blocks):
    template = []

    for para in docx_elements:
        for block in pdf_blocks:
            if para["text"] and para["text"] in block["text"]:
                template.append({
                    "text": para["text"],
                    "style": para["style"],
                    "runs": para["runs"],
                    "layout": {
                        "page": block["page"],
                        "bbox": block["bbox"]
                    }
                })
                break

    return template
