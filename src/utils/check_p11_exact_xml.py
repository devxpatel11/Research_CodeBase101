import docx
import lxml.etree as ET

doc_t = docx.Document('deliverables/Access-Template-2024 (1).docx')
p11 = doc_t.paragraphs[11]
print(ET.tostring(p11._element, pretty_print=True).decode('utf-8'))
