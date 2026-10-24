import docx

doc = docx.Document('deliverables/update_v4_edited.docx')
body = doc._body._element

print("=== BODY ELEMENTS IN update_v4_edited.docx ===")
elem_idx = 0
for elem in body:
    tag = elem.tag.split('}')[-1]
    if tag == 'p':
        p = docx.text.paragraph.Paragraph(elem, doc)
        text = p.text.strip().replace('\n', ' ')
        has_omml = bool(elem.xpath('.//m:oMath'))
        has_img = bool(elem.xpath('.//a:blip'))
        print(f"[{elem_idx:3d}] P (omml={has_omml}, img={has_img}): {text[:70]}")
    elif tag == 'tbl':
        t = docx.table.Table(elem, doc)
        has_omml = bool(elem.xpath('.//m:oMath'))
        print(f"[{elem_idx:3d}] TBL (rows={len(t.rows)}, cols={len(t.columns)}, omml={has_omml}): first_cell='{t.cell(0,0).text.strip()[:40]}'")
    elif tag == 'sectPr':
        print(f"[{elem_idx:3d}] SECTPR (end of doc)")
    elem_idx += 1
