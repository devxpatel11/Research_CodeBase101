import docx

doc_t = docx.Document('deliverables/Access-Template-2024 (1).docx')
for i, s in enumerate(doc_t.sections):
    p_in_s = [p.text[:40] for p in doc_t.paragraphs if p._element in s._sectPr.xpath('preceding-sibling::w:p')]
    print(f'=== Section {i} ===')
    print('Start type:', s.start_type)
    print('Cols:', s._sectPr.xpath('./w:cols/@w:num'), s._sectPr.xpath('./w:cols/@w:space'))
    # Print the paragraph that ends this section
    parent_p = s._sectPr.getparent()
    print('Parent tag:', parent_p.tag.split('}')[-1])
    if parent_p.tag.split('}')[-1] == 'p':
        p_obj = docx.text.paragraph.Paragraph(parent_p, doc_t)
        print('Paragraph text ending section:', repr(p_obj.text[:60]))
