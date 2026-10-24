"""
Test Clean 2-Column Architecture
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
TEST_DOCX = r"deliverables/test_clean_2col.docx"
TEST_PDF = r"deliverables/test_clean_2col.pdf"

doc_t = docx.Document(TEMPLATE_PATH)
doc_s = docx.Document(SOURCE_PATH)

body_t = doc_t._body._element
body_s = doc_s._body._element

# Paragraph 2: Paper Title
p_title = doc_t.paragraphs[2]
p_title.text = ""
p_title.add_run(doc_s.paragraphs[0].text.strip())
p_title.style = "Paper Title"

# Paragraph 3: Authors (AU)
p_auth = doc_t.paragraphs[3]
p_auth.text = ""
p1_lines = doc_s.paragraphs[1].text.strip().split('\n')
auth_lines = [l.strip().rstrip(',') for l in p1_lines[:-1] if l.strip()]
affil_line = p1_lines[-1].strip()
p_auth.add_run(", ".join(auth_lines))
p_auth.style = "AU"

# Paragraph 4: Affiliation (PI)
p_aff = doc_t.paragraphs[4]
p_aff.text = ""
p_aff.add_run(affil_line)
p_aff.style = "PI"

# In the template, paragraphs are:
# 0: DOP, 1: DOI, 2: Paper Title, 3: AU, 4: PI_No Space, 5: PI_No Space, 6: PI, 7: PI, 8: footnote text, 9: Abstract, 10: IT, 11: IT (has sectPr)
# Let's remove template dummy paragraphs 5, 6, 7 (excess affiliations)
body_t.remove(doc_t.paragraphs[5]._element)
body_t.remove(doc_t.paragraphs[6]._element)
body_t.remove(doc_t.paragraphs[7]._element)
# Now footnote text is paragraph 5, Abstract is 6, IT is 7, IT (sectPr) is 8.

# Replace Abstract (P6)
p_abs = doc_t.paragraphs[6]
p_abs._element.clear_content()
for child in doc_s.paragraphs[2]._element:
    tag = child.tag.split('}')[-1]
    if tag in ('r', 'oMath', 'oMathPara'):
        p_abs._element.append(copy.deepcopy(child))
p_abs.style = "Abstract"

# Replace IT (P7)
p_it = doc_t.paragraphs[7]
p_it._element.clear_content()
for child in doc_s.paragraphs[3]._element:
    tag = child.tag.split('}')[-1]
    if tag in ('r', 'oMath', 'oMathPara'):
        p_it._element.append(copy.deepcopy(child))
p_it.style = "IT"

# Paragraph 8 is the paragraph that contains the sectPr for Section 0!
p_sect = doc_t.paragraphs[8]
print(f"P_sect has sectPr: {bool(p_sect._element.xpath('.//w:sectPr'))}")

# Now, we remove everything AFTER p_sect (except the final body sectPr)
# In body_t, p_sect is at index 8. The final sectPr is at the end.
while len(body_t) > 10: # keep 0..8 and the final sectPr (index 9)
    body_t.remove(body_t[9])

print(f"Body children after cleaning dummy text: {len(body_t)}")
for i, c in enumerate(body_t):
    tag = c.tag.split('}')[-1]
    sect = bool(c.xpath('.//w:sectPr')) or tag == 'sectPr'
    print(f"  Child {i}: tag={tag}, sect={sect}")

# Now save and check
doc_t.save(TEST_DOCX)
print("Saved base test template.")
