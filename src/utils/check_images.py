import docx

doc = docx.Document('deliverables/update_v4_edited.docx')
print("=== IMAGES IN update_v4_edited.docx ===")
for r_id, rel in doc.part.rels.items():
    if "image" in rel.target_ref:
        img_part = rel.target_part
        print(f"Rel ID: {r_id} -> {rel.target_ref}, size: {len(img_part.blob)} bytes")
