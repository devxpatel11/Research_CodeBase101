import docx

doc_t = docx.Document('deliverables/Access-Template-2024 (1).docx')
print("=== TEMPLATE SECTIONS DETAIL ===")
for i, s in enumerate(doc_t.sections):
    print(f"--- Section {i} ---")
    print(f"  start_type: {s.start_type}")
    print(f"  page_width: {s.page_width.pt} pt ({s.page_width.inches} in)")
    print(f"  page_height: {s.page_height.pt} pt ({s.page_height.inches} in)")
    print(f"  top_margin: {s.top_margin.pt} pt ({s.top_margin.inches} in)")
    print(f"  bottom_margin: {s.bottom_margin.pt} pt ({s.bottom_margin.inches} in)")
    print(f"  left_margin: {s.left_margin.pt} pt ({s.left_margin.inches} in)")
    print(f"  right_margin: {s.right_margin.pt} pt ({s.right_margin.inches} in)")
    print(f"  header_distance: {s.header_distance.pt} pt")
    print(f"  footer_distance: {s.footer_distance.pt} pt")
    print(f"  different_first_page_header_footer: {s.different_first_page_header_footer}")
    
    # check cols in xml
    sectPr = s._sectPr
    cols = sectPr.xpath('./w:cols')
    if cols:
        num = cols[0].get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}num')
        space = cols[0].get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}space')
        print(f"  cols: num={num}, space={space}")
    else:
        print("  cols: default (1)")
