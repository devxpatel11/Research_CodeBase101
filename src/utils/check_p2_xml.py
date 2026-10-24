import docx

doc = docx.Document('deliverables/update_v4_edited.docx')
p2 = doc.paragraphs[2]

print("=== P2 XML CHILDREN ===")
for child in p2._element:
    tag = child.tag.split('}')[-1]
    if tag == 'r':
        print(f"Run: '{child.text}'")
    elif tag in ('oMath', 'oMathPara'):
        print(f"OMML Math: tag={tag}")
    else:
        print(f"Other: tag={tag}")
