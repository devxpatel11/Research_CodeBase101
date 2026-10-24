import docx

doc_t = docx.Document('deliverables/Access-Template-2024 (1).docx')
s_h1 = doc_t.styles['H1']
s_h1_list = doc_t.styles['H1_List (Space)']
print("H1 numPr:", bool(s_h1._element.xpath('.//w:numPr')))
print("H1_List numPr:", bool(s_h1_list._element.xpath('.//w:numPr')))
