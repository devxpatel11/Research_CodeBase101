import docx

doc_t = docx.Document('deliverables/Access-Template-2024 (1).docx')
print('Number of sections in template:', len(doc_t.sections))
for i, s in enumerate(doc_t.sections):
    num = s._sectPr.xpath('./w:cols/@w:num')
    space = s._sectPr.xpath('./w:cols/@w:space')
    print(f'Section {i}: start_type={s.start_type}, cols={num}, space={space}')

body = doc_t._body._element
print('\nBody children count:', len(body))
for i, child in enumerate(body):
    tag = child.tag.split('}')[-1]
    sectPr = child.xpath('.//w:sectPr')
    if sectPr or tag == 'sectPr':
        print(f'Child {i}: tag={tag}, has sectPr! cols={sectPr[0].xpath("./w:cols/@w:num") if sectPr else "final"}')
