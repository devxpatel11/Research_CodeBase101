import docx

doc_t = docx.Document('deliverables/Access-Template-2024 (1).docx')
body = doc_t._body._element
final_sectPr = body.xpath('./w:sectPr')
if final_sectPr:
    print('Final sectPr in body:')
    print(final_sectPr[0].xml)
