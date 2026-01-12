import xmltodict
from zipfile import ZipFile

def xml_to_dict(docx_path):
    with ZipFile(docx_path) as z:
        xml_bytes = z.read("word/document.xml")
    xml_str = xml_bytes.decode("utf-8")
    return xmltodict.parse(xml_str)

doc_dict = xml_to_dict("./data/sample3.docx")

print(doc_dict)