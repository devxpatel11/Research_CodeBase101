import docx
from docx.enum.section import WD_SECTION_START
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import win32com.client
import pymupdf
import os

doc = docx.Document('deliverables/Access-Template-2024 (1).docx')
# Keep first 12 paragraphs (Section 0)
body = doc._body._element
while len(body) > 9:
    body.remove(body[8])

# We are now in Section 1 (2 columns)
doc.add_paragraph("This is text in two columns. " * 20, style="PARA")

# Now add a section for a wide table (1 column)
s_wide = doc.add_section(WD_SECTION_START.CONTINUOUS)
# Set cols to 1
cols_elem = s_wide._sectPr.xpath('./w:cols')
if cols_elem:
    s_wide._sectPr.remove(cols_elem[0])
s_wide._sectPr.append(parse_xml(f'<w:cols {nsdecls("w")} w:space="720"/>'))

p_wide = doc.add_paragraph("TABLE I: WIDE TABLE TITLE ACROSS PAGE", style="Table Title")
tbl = doc.add_table(rows=3, cols=6)
tbl.alignment = docx.enum.table.WD_TABLE_ALIGNMENT.CENTER
for r in tbl.rows:
    for c in r.cells:
        c.text = "Sample Data"
        c.width = docx.shared.Inches(1.1)

# Now switch back to 2 columns
s_2col = doc.add_section(WD_SECTION_START.CONTINUOUS)
cols_elem = s_2col._sectPr.xpath('./w:cols')
if cols_elem:
    s_2col._sectPr.remove(cols_elem[0])
s_2col._sectPr.append(parse_xml(f'<w:cols {nsdecls("w")} w:num="2" w:space="400"/>'))

doc.add_paragraph("This is text back in two columns after the wide table. " * 20, style="PARA")

test_path = 'deliverables/test_section_switch.docx'
test_pdf = 'deliverables/test_section_switch.pdf'
doc.save(test_path)
print(f"Saved {test_path}")

word = win32com.client.Dispatch("Word.Application")
doc_w = word.Documents.Open(os.path.abspath(test_path))
doc_w.ExportAsFixedFormat(os.path.abspath(test_pdf), 17)
doc_w.Close(False)
word.Quit()
print(f"Exported {test_pdf}")

# Check rendering
doc_pdf = pymupdf.open(test_pdf)
print("Total pages:", len(doc_pdf))
page1 = doc_pdf[0]
pix = page1.get_pixmap(dpi=150)
pix.save('deliverables/test_section_switch_p1.png')
print("Rendered test_section_switch_p1.png")
