import docx

doc_t = docx.Document('deliverables/Access-Template-2024 (1).docx')
for i, p in enumerate(doc_t.paragraphs):
    if p._element.xpath('.//a:blip'):
        print(f"Template Image at p{i}, style={p.style.name}, text='{p.text}'")
        # Check surrounding section breaks
        sect_prev = p._element.xpath('preceding-sibling::w:p[w:pPr/w:sectPr]')
        sect_inside = p._element.xpath('.//w:sectPr')
        print(f"   sect inside p: {bool(sect_inside)}")
