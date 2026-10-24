import docx

doc_t = docx.Document('deliverables/Access-Template-2024 (1).docx')
p40 = doc_t.paragraphs[40]
extent = p40._element.xpath('.//wp:extent')
if extent:
    cx = int(extent[0].get('cx'))
    cy = int(extent[0].get('cy'))
    print(f"Template p40 image size: {cx/914400.0:.2f} x {cy/914400.0:.2f} inches")
