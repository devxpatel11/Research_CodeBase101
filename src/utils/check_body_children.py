import docx

doc_t = docx.Document('deliverables/Access-Template-2024 (1).docx')
body = doc_t._body._element

print(f"Total children in template body: {len(body)}")
for i, child in enumerate(body[:16]):
    tag = child.tag.split('}')[-1]
    text = child.text.strip().replace('\n', ' ') if child.text else ""
    if tag == 'p':
        p = docx.text.paragraph.Paragraph(child, doc_t)
        has_sect = bool(child.xpath('./w:pPr/w:sectPr'))
        print(f"[{i:02d}] P (style={p.style.name}, sect={has_sect}): {p.text[:60]}")
    elif tag == 'tbl':
        print(f"[{i:02d}] TBL")
    else:
        print(f"[{i:02d}] {tag}")
