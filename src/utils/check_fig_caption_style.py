import docx

doc_t = docx.Document('deliverables/Access-Template-2024 (1).docx')
st = doc_t.styles['Fig Caption']
print('Fig Caption pPr:', st._element.pPr.xml if st._element.pPr is not None else None)
p41 = doc_t.paragraphs[41]
print('P41 xml:', p41._element.xml)
