import docx

doc_t = docx.Document('deliverables/Access-Template-2024 (1).docx')
body = doc_t._body._element

p_idx = 0
for elem in body:
    if elem.tag.endswith('p'):
        sectPr = elem.xpath('./w:pPr/w:sectPr')
        p = docx.text.paragraph.Paragraph(elem, doc_t)
        if sectPr:
            cols = sectPr[0].xpath('./w:cols')
            col_num = cols[0].get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}num') if cols else '1'
            print(f"Section break after p{p_idx} (text='{p.text[:40]}...'): cols={col_num}")
        p_idx += 1
    elif elem.tag.endswith('tbl'):
        pass
