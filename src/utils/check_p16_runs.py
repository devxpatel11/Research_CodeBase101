import docx

doc = docx.Document('deliverables/update_v4_edited.docx')
for i in [16, 17, 18, 19, 20]:
    p = doc.paragraphs[i]
    print(f"\n=== P{i} ({len(p.runs)} runs) ===")
    for r in p.runs[:4]:
        print(f"  run: '{r.text}' | bold={r.bold} | italic={r.italic}")
