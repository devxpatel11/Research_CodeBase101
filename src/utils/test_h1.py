import docx
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

doc_t = docx.Document('deliverables/Access-Template-2024 (1).docx')
p = doc_t.add_paragraph("I. INTRODUCTION", style="H1")
print(f"p style: {p.style.name}, font: {p.style.font.name}, size: {p.style.font.size}")
