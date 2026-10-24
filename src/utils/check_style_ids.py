import docx

doc_t = docx.Document('deliverables/Access-Template-2024 (1).docx')
print("=== TEMPLATE STYLE ID vs NAME ===")
for s in doc_t.styles:
    if s.type == docx.enum.style.WD_STYLE_TYPE.PARAGRAPH:
        print(f"Style ID: '{s.style_id:25}' | Name: '{s.name}'")
