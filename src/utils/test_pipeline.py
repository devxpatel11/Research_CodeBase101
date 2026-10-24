"""
Test script to see how Word renders:
1. Clean Section 0 to Section 1 continuous break
2. Images placed in 2-column flow (width=3.35 in)
3. Tables tested both ways
"""

import docx
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
import win32com.client
import pymupdf
import os
import copy

TEMPLATE_PATH = r"deliverables/Access-Template-2024 (1).docx"
SOURCE_PATH = r"deliverables/update_v4_edited.docx"
TEST_DOCX = r"deliverables/test_access_v3.docx"
TEST_PDF = r"deliverables/test_access_v3.pdf"

doc_t = docx.Document(TEMPLATE_PATH)
doc_s = docx.Document(SOURCE_PATH)

# Let's inspect the sections in doc_t
print(f"Original template sections: {len(doc_t.sections)}")
for i, s in enumerate(doc_t.sections):
    print(f"Section {i}: cols={s._sectPr.xpath('./w:cols/@w:num')}, start={s.start_type}")
