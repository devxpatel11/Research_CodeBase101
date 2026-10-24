import docx

doc = docx.Document('deliverables/update_v4_edited.docx')
print("=== PICTURE SIZES IN update_v4_edited.docx ===")
for i, p in enumerate(doc.paragraphs):
    blips = p._element.xpath('.//a:blip')
    if blips:
        extent = p._element.xpath('.//wp:extent')
        if extent:
            cx = int(extent[0].get('cx'))
            cy = int(extent[0].get('cy'))
            cx_in = cx / 914400.0
            cy_in = cy / 914400.0
            print(f"P{i} image size: {cx_in:.2f} x {cy_in:.2f} inches ({cx} x {cy} emu)")
