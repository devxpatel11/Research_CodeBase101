import docx
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def create_1col_break():
    p = parse_xml(
        f'<w:p {nsdecls("w")}>'
        f'<w:pPr>'
        f'<w:sectPr>'
        f'<w:type w:val="continuous"/>'
        f'<w:cols w:space="720"/>'
        f'</w:sectPr>'
        f'</w:pPr>'
        f'</w:p>'
    )
    return p

def create_2col_break():
    p = parse_xml(
        f'<w:p {nsdecls("w")}>'
        f'<w:pPr>'
        f'<w:sectPr>'
        f'<w:type w:val="continuous"/>'
        f'<w:cols w:num="2" w:space="400"/>'
        f'</w:sectPr>'
        f'</w:pPr>'
        f'</w:p>'
    )
    return p

print("Section break helper functions defined.")
