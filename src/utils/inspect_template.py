import docx

doc_t = docx.Document('deliverables/Access-Template-2024 (1).docx')
print("=== TEMPLATE STYLES AND FONTS ===")
for s in doc_t.styles:
    if s.type == docx.enum.style.WD_STYLE_TYPE.PARAGRAPH:
        font_name = s.font.name if s.font else None
        font_size = s.font.size.pt if s.font and s.font.size else None
        bold = s.font.bold if s.font else None
        italic = s.font.italic if s.font else None
        print(f"Style: '{s.name}' | font: {font_name} | size: {font_size}pt | bold: {bold} | italic: {italic}")
