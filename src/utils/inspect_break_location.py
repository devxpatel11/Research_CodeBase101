import docx

doc_t = docx.Document('deliverables/Access-Template-2024 (1).docx')
for i in range(8, 16):
    p = doc_t.paragraphs[i]
    sect = p._element.xpath('.//w:sectPr')
    print(f'P{i:2d} | style={p.style.name:20s} | has_sectPr={bool(sect)} | text={repr(p.text[:40])}')
