import docx

doc_t = docx.Document('deliverables/Access-Template-2024 (1).docx')
st = doc_t.styles['Table Title']
print('Table Title pPr:', st._element.pPr.xml if st._element.pPr is not None else None)
