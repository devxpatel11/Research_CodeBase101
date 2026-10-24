import docx
import win32com.client
import os
import pymupdf
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

doc_t = docx.Document('deliverables/Access-Template-2024 (1).docx')
# Let's inspect the sections in doc_t
print("doc_t sections:", len(doc_t.sections))
for i, s in enumerate(doc_t.sections):
    print(f"Section {i}: start_type={s.start_type}")
