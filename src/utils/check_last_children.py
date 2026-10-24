import docx

doc_t = docx.Document('deliverables/Access-Template-2024 (1).docx')
body = doc_t._body._element

print("=== LAST 5 CHILDREN IN TEMPLATE BODY ===")
for i in range(len(body)-5, len(body)):
    elem = body[i]
    tag = elem.tag.split('}')[-1]
    if tag == 'p':
        p = docx.text.paragraph.Paragraph(elem, doc_t)
        has_sect = bool(elem.xpath('./w:pPr/w:sectPr'))
        print(f"[{i}] P (style={p.style.name}, sect={has_sect}): {p.text[:40]}")
    else:
        print(f"[{i}] {tag}")
