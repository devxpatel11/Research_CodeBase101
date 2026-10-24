import docx
import os

doc = docx.Document('deliverables/update_v4_edited.docx')
print("=== EXTRACTING IMAGES FROM update_v4_edited.docx ===")
os.makedirs('deliverables/extracted_images', exist_ok=True)
for r_id, rel in doc.part.rels.items():
    if "image" in rel.target_ref:
        img_part = rel.target_part
        filename = os.path.basename(rel.target_ref)
        out_path = os.path.join('deliverables/extracted_images', filename)
        with open(out_path, 'wb') as f:
            f.write(img_part.blob)
        print(f"Saved {filename} ({len(img_part.blob)} bytes)")

print("\n=== COMPARING WITH figures/ DIRECTORY ===")
for f in os.listdir('figures'):
    if f.endswith('.png'):
        f_path = os.path.join('figures', f)
        print(f"{f}: {os.path.getsize(f_path)} bytes")
