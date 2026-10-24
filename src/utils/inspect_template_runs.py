import docx
import sys

sys.stdout.reconfigure(encoding='utf-8')
doc_t = docx.Document('deliverables/Access-Template-2024 (1).docx')

print("=== TEMPLATE FIRST 20 PARAGRAPHS WITH EXACT DETAILS ===")
for i, p in enumerate(doc_t.paragraphs[:25]):
    text = p.text.strip().replace('\n', ' ')
    if not text:
        continue
    runs_detail = []
    for r in p.runs:
        runs_detail.append(f"'{r.text}'(font={r.font.name}, sz={r.font.size.pt if r.font.size else None}, b={r.font.bold}, i={r.font.italic})")
    print(f"P{i} [{p.style.name}] (align={p.alignment}): {text[:60]}")
    print(f"   runs: {' '.join(runs_detail[:4])}")
