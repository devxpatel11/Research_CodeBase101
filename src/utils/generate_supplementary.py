"""
Supplementary Methods Generator:
Produces deliverables/supplementary_methods.docx and deliverables/supplementary_methods.pdf
Adheres strictly to IEEE formatting: Times New Roman, 1-inch margins, 1.15 line spacing.
Renders all mathematical formulas, powers, fractions, and under-roots natively via Office Math (OMML).
Contains:
  - Complete mathematical derivations (RankNet gradients, Delta-NDCG, Newton leaf updates, monotonicity proof).
  - Under-root standard error derivation.
  - Non-compensatory gate proof and numerical verification.
  - Complete audit ledgers (TenderKG, Spanish BOE, Latvian CFLA).
  - Algorithmic pseudocode for all 4 core pipeline stages.
  - Hyperparameter optimization and tuning specifications.
  - Industrial manufacturing case study walk-through (aerospace machining).
  - Complete feature lineage and data dictionary tables.
"""

import os
import sys
import re
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
import win32com.client
import lxml.etree as ET
from latex2mathml.converter import convert as l2m

sys.stdout.reconfigure(encoding='utf-8')

DELIVERABLES_DIR = "deliverables"
XSLT_PATH = r"C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL"

# Load XSLT stylesheet for MathML to OMML conversion
xslt_tree = ET.parse(XSLT_PATH)
omml_transform = ET.XSLT(xslt_tree)

def latex_to_omml(latex_code):
    if not latex_code:
        return None
    try:
        mathml = l2m(latex_code.strip())
        dom = ET.fromstring(mathml)
        omml = omml_transform(dom)
        return ET.tostring(omml, encoding='utf-8').decode('utf-8')
    except Exception as e:
        return None

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_table_borders_none(table):
    tblPr = table._tbl.tblPr
    tblBorders = parse_xml(f'<w:tblBorders {nsdecls("w")}><w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/><w:insideH w:val="none"/><w:insideV w:val="none"/></w:tblBorders>')
    tblPr.append(tblBorders)

def add_styled_heading(doc, text, level):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.name = "Times New Roman"
    run.bold = True
    run.font.color.rgb = RGBColor(0, 0, 0)
    if level == 1:
        run.font.size = Pt(14)
        p.paragraph_format.space_before = Pt(16)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.keep_with_next = True
    elif level == 2:
        run.font.size = Pt(12)
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.keep_with_next = True
    elif level == 3:
        run.font.size = Pt(12)
        run.italic = True
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.keep_with_next = True
    return p

def add_body_paragraph(doc, text, bold_prefix=None, space_after=6):
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.15
    p.paragraph_format.space_after = Pt(space_after)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    if bold_prefix:
        r_pre = p.add_run(bold_prefix)
        r_pre.font.name = "Times New Roman"
        r_pre.font.size = Pt(12)
        r_pre.bold = True
        r_pre.font.color.rgb = RGBColor(0, 0, 0)

    # Parse and cleanly render inline math expressions marked by $...$
    tokens = re.split(r'(\$[^\$]+\$)', text)
    for token in tokens:
        if not token:
            continue
        if token.startswith('$') and token.endswith('$') and len(token) > 2:
            math_content = token[1:-1].strip()
            omml = latex_to_omml(math_content)
            if omml:
                p._p.append(parse_xml(omml))
            else:
                r = p.add_run(math_content)
                r.font.name = "Times New Roman"
                r.font.size = Pt(12)
                r.italic = True
                r.font.color.rgb = RGBColor(0, 0, 0)
        else:
            r = p.add_run(token)
            r.font.name = "Times New Roman"
            r.font.size = Pt(12)
            r.font.color.rgb = RGBColor(0, 0, 0)
    return p

def add_equation_paragraph(doc, eq_text, eq_num=None):
    omml_str = latex_to_omml(eq_text)
    tbl = doc.add_table(rows=1, cols=2)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders_none(tbl)
    
    cell0 = tbl.rows[0].cells[0]
    cell1 = tbl.rows[0].cells[1]
    cell0.width = Inches(5.8)
    cell1.width = Inches(0.7)
    set_cell_margins(cell0, top=40, bottom=40, left=40, right=40)
    set_cell_margins(cell1, top=40, bottom=40, left=40, right=40)
    
    # Cell 0: Centered Equation with clear horizontal fraction bars, powers, and roots
    p0 = cell0.paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p0.paragraph_format.line_spacing = 1.15
    p0.paragraph_format.space_before = Pt(4)
    p0.paragraph_format.space_after = Pt(4)
    if omml_str:
        p0._p.append(parse_xml(omml_str))
    else:
        r = p0.add_run(eq_text)
        r.font.name = "Times New Roman"
        r.font.size = Pt(11)
        r.italic = True
        
    # Cell 1: Right-aligned Equation Number
    p1 = cell1.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p1.paragraph_format.line_spacing = 1.15
    p1.paragraph_format.space_before = Pt(4)
    p1.paragraph_format.space_after = Pt(4)
    if eq_num:
        r_num = p1.add_run(f"({eq_num})")
        r_num.font.name = "Times New Roman"
        r_num.font.size = Pt(11)
        r_num.bold = True
        r_num.font.color.rgb = RGBColor(0, 0, 0)
        
    p_sp = doc.add_paragraph()
    p_sp.paragraph_format.space_before = Pt(0)
    p_sp.paragraph_format.space_after = Pt(4)
    p_sp.paragraph_format.line_spacing = 1.0
    return tbl

def add_code_block(doc, code_text):
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.0
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.left_indent = Inches(0.25)
    r = p.add_run(code_text)
    r.font.name = "Courier New"
    r.font.size = Pt(9.5)
    r.font.color.rgb = RGBColor(30, 30, 30)
    return p

def format_table(table, col_widths, headers, data, align_cols=None):
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr_cells = table.rows[0].cells
    for i, title in enumerate(headers):
        hdr_cells[i].text = title
        set_cell_background(hdr_cells[i], "1F4E79")
        set_cell_margins(hdr_cells[i], top=120, bottom=120, left=150, right=150)
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in p.runs:
            r.font.name = "Times New Roman"
            r.font.size = Pt(10)
            r.bold = True
            r.font.color.rgb = RGBColor(255, 255, 255)

    for row_idx, row_data in enumerate(data):
        row_cells = table.rows[row_idx + 1].cells
        bg_color = "F2F5F9" if row_idx % 2 == 1 else "FFFFFF"
        for col_idx, cell_value in enumerate(row_data):
            cell = row_cells[col_idx]
            cell.text = "" # Clear default text
            set_cell_background(cell, bg_color)
            set_cell_margins(cell, top=80, bottom=80, left=150, right=150)
            p = cell.paragraphs[0]
            if align_cols and col_idx < len(align_cols):
                p.alignment = align_cols[col_idx]
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT

            # Check if cell contains math $...$
            tokens = re.split(r'(\$[^\$]+\$)', str(cell_value))
            for token in tokens:
                if not token:
                    continue
                if token.startswith('$') and token.endswith('$') and len(token) > 2:
                    math_content = token[1:-1].strip()
                    omml = latex_to_omml(math_content)
                    if omml:
                        p._p.append(parse_xml(omml))
                    else:
                        r = p.add_run(math_content)
                        r.font.name = "Times New Roman"
                        r.font.size = Pt(9.5)
                        r.italic = True
                else:
                    r = p.add_run(token)
                    r.font.name = "Times New Roman"
                    r.font.size = Pt(9.5)
                    r.font.color.rgb = RGBColor(0, 0, 0)

    for row in table.rows:
        for idx, width in enumerate(col_widths):
            row.cells[idx].width = Inches(width)

def build_supplementary_docx():
    doc = docx.Document()
    
    # Page setup: Standard Letter, 1-inch margins
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
        section.page_width = Inches(8.5)
        section.page_height = Inches(11.0)

    # Title
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(8)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run("Supplementary Methods: Mathematical Derivations, Empirical Audit Ledgers, and Technical Verification")
    r_title.font.name = "Times New Roman"
    r_title.font.size = Pt(18)
    r_title.bold = True
    r_title.font.color.rgb = RGBColor(0, 0, 0)

    # Subtitle
    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(14)
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = p_sub.add_run("Accompanying Material for: Evidence-Aware Compliance-Constrained Learning-to-Rank for B2B Supplier Matching")
    r_sub.font.name = "Times New Roman"
    r_sub.font.size = Pt(12)
    r_sub.italic = True
    r_sub.font.color.rgb = RGBColor(50, 50, 50)

    # Author
    p_auth = doc.add_paragraph()
    p_auth.paragraph_format.space_before = Pt(0)
    p_auth.paragraph_format.space_after = Pt(20)
    p_auth.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_auth = p_auth.add_run("Dev Arpit Patel\nSchool of Computer Science and Engineering, Vellore Institute of Technology, Vellore 632014, India\nEmail: dev.patel2024@vitstudent.ac.in | ORCID: 0009-0002-8412-9104")
    r_auth.font.name = "Times New Roman"
    r_auth.font.size = Pt(11)
    r_auth.font.color.rgb = RGBColor(40, 40, 40)

    # Preamble
    add_body_paragraph(doc, 
        "This supplementary document provides complete mathematical derivations, proof of monotonicity, non-compensatory boundary checks, exact reproducible audit ledgers for TenderKG v1, Spanish BOE v3, and Latvian CFLA, formal algorithmic procedures, detailed hyperparameter configurations, an end-to-end industrial manufacturing case walk-through, and comprehensive data dictionary and feature lineage tables. All code, datasets, scripts, logs, and generated tables correspond directly to the accompanying experimental repository.",
        bold_prefix="Overview and Purpose: ")

    # Section S1
    add_styled_heading(doc, "Section S1: Detailed Mathematical Derivations and Analytical Proofs", 1)
    
    add_styled_heading(doc, "S1.1 Derivation of RankNet Pairwise Loss and LambdaGradients", 2)
    add_body_paragraph(doc,
        "In pairwise learning-to-rank, each procurement query $q$ is associated with candidate supplier pairs $(i, j)$ where candidate $i$ possesses ground-truth relevance label $y_i$ and candidate $j$ possesses label $y_j$. Let $f_i = f_\\theta(x(q, s_i))$ and $f_j = f_\\theta(x(q, s_j))$ represent the real-valued model outputs. In the probabilistic ranking formulation (Burges et al., 2005), the modeled probability that candidate $i$ is ranked higher than candidate $j$ is parameterized via the logistic sigmoid function:")
    add_equation_paragraph(doc, r"P_{ij} = P(s_i \succ s_j) = \frac{1}{1 + e^{-\sigma (f_i - f_j)}} = \sigma_{\text{logit}}(\sigma (f_i - f_j))", "S1")
    add_body_paragraph(doc,
        "where $\\sigma > 0$ denotes the inverse temperature parameter governing sigmoid slope steepness. Let $\\bar{P}_{ij}$ denote the observed posterior probability based on true labels $y_i$ and $y_j$. When labels are ordinal integers (with $y_i > y_j$), $\\bar{P}_{ij} = 1$; when $y_i = y_j$, $\\bar{P}_{ij} = 0.5$; and when $y_i < y_j$, $\\bar{P}_{ij} = 0$. The cross-entropy cost $C_{ij}$ for pair $(i, j)$ is defined as:")
    add_equation_paragraph(doc, r"C_{ij} = - \bar{P}_{ij} \ln P_{ij} - (1 - \bar{P}_{ij}) \ln(1 - P_{ij})", "S2")
    add_body_paragraph(doc,
        "For pairs where candidate $i$ has strictly higher relevance than candidate $j$ ($\\bar{P}_{ij} = 1$), the loss simplifies to $C_{ij} = -\\ln P_{ij} = \\ln(1 + e^{-\\sigma (f_i - f_j)})$. Differentiating $C_{ij}$ with respect to model score $f_i$ yields:")
    add_equation_paragraph(doc, r"\frac{\partial C_{ij}}{\partial f_i} = \frac{\partial}{\partial f_i} \ln(1 + e^{-\sigma(f_i - f_j)}) = \frac{-\sigma e^{-\sigma(f_i - f_j)}}{1 + e^{-\sigma(f_i - f_j)}} = - \frac{\sigma}{1 + e^{\sigma(f_i - f_j)}} = -\sigma (1 - P_{ij})", "S3")
    add_body_paragraph(doc,
        "By anti-symmetry, $\\frac{\\partial C_{ij}}{\\partial f_j} = - \\frac{\\partial C_{ij}}{\\partial f_i} = +\\sigma(1 - P_{ij})$. In gradient-boosted decision tree optimization, the negative gradient (often called the pseudo-residual or virtual force) $\\lambda_{ij}$ indicates the direction and magnitude of score adjustment required for candidate $i$:")
    add_equation_paragraph(doc, r"\lambda_{ij} = - \frac{\partial C_{ij}}{\partial f_i} = \frac{\sigma}{1 + e^{\sigma(f_i - f_j)}} = \sigma(1 - P_{ij})", "S4")

    add_styled_heading(doc, "S1.2 Integration of Metric Weighting: LambdaMART Formulation", 2)
    add_body_paragraph(doc,
        "Direct optimization of NDCG@k is computationally intractable due to the non-differentiable sorting operator. LambdaMART resolves this by scaling the pairwise RankNet force $\\lambda_{ij}$ by the absolute change in the target metric $\\Delta \\text{NDCG}_{ij}$ that would result from swapping the rank positions of candidate $i$ and candidate $j$ within query $q$:")
    add_equation_paragraph(doc, r"\lambda_{ij}^{\text{MART}} = \Delta \text{NDCG}_{ij} \cdot \lambda_{ij} = |\text{NDCG}_{\text{swap}(i,j)} - \text{NDCG}_{\text{orig}}| \cdot \frac{\sigma}{1 + e^{\sigma(f_i - f_j)}}", "S5")
    add_body_paragraph(doc,
        "For a query with ideal DCG normalizer $\\text{IDCG}_q$, the rank positions $r_i$ and $r_j$ induce discount factors $D(r) = \\frac{1}{\\log_2(r + 1)}$. The absolute difference when swapping rank positions $r_i$ and $r_j$ evaluates directly to:")
    add_equation_paragraph(doc, r"\Delta \text{NDCG}_{ij} = \frac{|2^{y_i} - 2^{y_j}| \cdot |D(r_i) - D(r_j)|}{\text{IDCG}_q}, \quad D(r) = \frac{1}{\log_2(r+1)}", "S6")
    add_body_paragraph(doc,
        "Summing the pairwise forces across all valid within-query pairs where $i$ is paired with any other candidate $j$ (either as the preferred candidate in $\\mathcal{P}_q^+$ or the dispreferred candidate in $\\mathcal{P}_q^-$), the total virtual gradient $g_i$ accumulated on candidate $i$ is:")
    add_equation_paragraph(doc, r"g_i = \sum_{j: (i,j) \in \mathcal{P}_q^+} \lambda_{ij}^{\text{MART}} - \sum_{j: (j,i) \in \mathcal{P}_q^-} \lambda_{ji}^{\text{MART}}", "S7")

    add_styled_heading(doc, "S1.3 Second Derivative (Hessian) and Regularized Newton Leaf Updates", 2)
    add_body_paragraph(doc,
        "To perform second-order tree splitting and optimal leaf value estimation, LightGBM computes the diagonal elements of the Hessian matrix. Differentiating $\\lambda_{ij}^{\\text{MART}}$ with respect to $f_i$:")
    add_equation_paragraph(doc, r"h_{ij} = \frac{\partial^2 C_{ij}}{\partial f_i^2} = \frac{\partial}{\partial f_i} \left[ \sigma (1 - P_{ij}) \right] = \sigma^2 P_{ij} (1 - P_{ij}) \cdot \Delta \text{NDCG}_{ij}", "S8")
    add_body_paragraph(doc,
        "The total Hessian $h_i$ accumulated for candidate $i$ is given by summing over all active pairs: $h_i = \\sum_{j} h_{ij}$. When constructing tree split terminal leaves, each leaf region $R_{lm}$ (m-th leaf of tree $l$) collects a set of training instances. The optimal leaf value $\\gamma_{lm}$ minimizing the second-order Taylor expansion of the objective, subject to L2 leaf regularization parameter $\\lambda_{\\text{reg}}$, is given by:")
    add_equation_paragraph(doc, r"\gamma_{lm} = - \frac{\sum_{i \in R_{lm}} g_i}{\sum_{i \in R_{lm}} h_i + \lambda_{\text{reg}}}", "S9")
    add_body_paragraph(doc,
        "This formulation guarantees numerical stability even when candidate relevance differences are small or when leaf partitions contain few training observations.")

    add_styled_heading(doc, "S1.4 Proof of Monotonicity for the Evidence-Aware Utility Function", 2)
    add_body_paragraph(doc,
        "Theorem 1 (Compliance Monotonicity): Let $U(q, s)$ be the final utility score defined in Section IV-D of the main manuscript: $U(q, s) = (1 - \\beta) a(q, s) + \\beta \\frac{\\text{RCI}^-(q, s)}{100}$, where $\\beta \\in [0, 1]$, $a(q, s) = \\sigma_{\\text{logit}}((f(q,s) - \\mu)/T)$, and $\\text{RCI}^-(q,s) = 100 \\sum_{j=1}^m w_j l_j$ with $w_j \\ge 0$ and $\\sum_{j=1}^m w_j = 1$. Then, for any optional compliance criterion $j$, the marginal utility with respect to the verified compliance lower bound $l_j$ is strictly non-negative:")
    add_equation_paragraph(doc, r"\frac{\partial U(q, s)}{\partial l_j} = \frac{\beta}{100} \frac{\partial \text{RCI}^-}{\partial l_j} = \frac{\beta}{100} \cdot 100 w_j = \beta w_j \ge 0", "S10")
    add_body_paragraph(doc,
        "Proof: By definition, the learned relevance score $f(q, s)$ is computed from past transactional history and commercial attributes at decision time $t$, which is independent of the prospective optional compliance evidence submitted for query $q$. Thus $\\frac{\\partial a(q, s)}{\\partial l_j} = 0$. Furthermore, $\\text{RCI}^-$ is a linear combination of lower bounds $l_j$ weighted by non-negative importance weights $w_j$. Therefore, $\\frac{\\partial \\text{RCI}^-}{\\partial l_j} = 100 w_j$. Multiplying by $\\frac{\\beta}{100}$ yields $\\beta w_j$. Since $\\beta \\ge 0$ and $w_j \\ge 0$, the derivative is non-negative for all $j$. When $\\beta > 0$ and $w_j > 0$, the derivative is strictly positive ($\\beta w_j > 0$). Q.E.D.")
    add_body_paragraph(doc,
        "Corollary 1 (Incentive Compatibility): A supplier submitting verified positive compliance evidence (increasing $l_j$ from 0 to 1) strictly improves or preserves their composite utility $U(q, s)$. Conversely, withholding or failing to verify evidence (leaving $l_j = 0$) cannot increase utility. Crucially, the upper bound $u_j$ does not enter $U(q, s)$, preventing suppliers from inflating their utility through unverified claims (since $\\frac{\\partial U}{\\partial u_j} = 0$).")

    add_styled_heading(doc, "S1.5 Non-Compensatory Proof of the Mandatory Eligibility Gate", 2)
    add_body_paragraph(doc,
        "Theorem 2 (Non-Compensatory Mandatory Clearance): Let $G(q, s, t) \\in \\{\\text{REJECT}, \\text{REVIEW}, \\text{ELIGIBLE}\\}$ denote the mandatory eligibility decision function. For any supplier $s$ and query $q$, if there exists at least one applicable mandatory requirement $j \\in \\mathcal{R}_q^{\\text{mand}}$ such that $r_j(q, s, t) = \\text{FAIL}$, then $s \\notin \\mathcal{F}(q, t)$ regardless of the magnitude of learned relevance $f(q, s) \\in \\mathbb{R}$ or optional compliance $\\text{RCI}^-(q, s) \\in [0, 100]$.")
    add_body_paragraph(doc,
        "Proof: Feasible candidate set membership is defined as $\\mathcal{F}(q, t) = \\{s \\in \\mathcal{C}(q, t) : G(q, s, t) = \\text{ELIGIBLE}\\}$. The decision rule specifies that $G(q, s, t) = \\text{REJECT}$ if $\\exists j \\in \\mathcal{R}_q^{\\text{mand}} : r_j(q, s, t) = \\text{FAIL}$. Since $\\text{REJECT} \\ne \\text{ELIGIBLE}$, $s \\notin \\mathcal{F}(q, t)$. Even in the asymptotic limit $f(q, s) \\to +\\infty$, the mapping $a(q, s) = \\sigma_{\\text{logit}}((f - \\mu)/T) \\to 1$, yet candidate $s$ is pruned prior to ranking. Hence, no finite or infinite relevance score can compensate for a mandatory statutory failure. Q.E.D.")

    add_styled_heading(doc, "S1.6 Cluster Bootstrap Standard Error and Under-Root Formulation", 2)
    add_body_paragraph(doc,
        "To evaluate statistical uncertainty without violating tender grouping, we draw $B = 2,000$ cluster bootstrap resamples grouped by parent tender. The empirical standard error $\\text{SE}(\\hat{\\theta})$ incorporates a clear radical under-root over the resample variance:")
    add_equation_paragraph(doc, r"\text{SE}(\hat{\theta}) = \sqrt{\frac{1}{B-1} \sum_{b=1}^B (\hat{\theta}_b - \bar{\theta})^2}", "S11")

    add_styled_heading(doc, "S1.7 Numerical Boundary and Verification Checks", 2)
    add_body_paragraph(doc,
        "To verify numerical stability and conformance with Section 10 of the Research Guide, the implementation was tested against reference boundary cases. Table S1 summarizes the test suite results.")

    # Table S1: Numerical Verification with native math formatting
    t_num = doc.add_table(rows=8, cols=4)
    format_table(t_num, [2.2, 2.0, 1.3, 1.0],
        ["Test Case / Property", "Analytical Formulation", "Evaluated Value", "Status"],
        [
            ["Optional RCI Lower Bound", r"$\text{RCI}^- = 100 \sum_{j=1}^m w_j l_j$", "50.0000", "PASS"],
            ["Optional RCI Upper Bound", r"$\text{RCI}^+ = 100 \sum_{j=1}^m w_j u_j$", "80.0000", "PASS"],
            ["Evidence Completeness Index", r"$Q = \sum_{j=1}^m w_j (1 - (u_j - l_j))$", "0.7000", "PASS"],
            ["Utility Mapping (beta = 0.2)", r"$U = (1-0.2) \cdot 0.75 + 0.2 \cdot \frac{50}{100}$", "0.7000", "PASS"],
            ["Pairwise RankNet Gradient", r"$\lambda_{ij} = \frac{\sigma}{1 + e^{\sigma \Delta f}}$", "+0.184535", "PASS"],
            ["Second-order Hessian", r"$h_{ij} = \sigma^2 P_{ij} (1 - P_{ij})$", "0.092268", "PASS"],
            ["Utility Monotonicity Check", r"$\frac{\partial U}{\partial l_j} = \beta w_j \ge 0$", "+0.060000", "PASS"]
        ],
        align_cols=[WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER])

    # Section S2
    add_styled_heading(doc, "Section S2: Complete Empirical Dataset Audits and Quality Ledgers", 1)
    add_body_paragraph(doc,
        "To ensure total reproducibility and transparent documentation of real-world procurement data imperfections, we detail the complete audit ledgers for the three primary empirical corpora examined in this study.")

    add_styled_heading(doc, "S2.1 TenderKG v1 Forensic Audit Ledger", 2)
    add_body_paragraph(doc,
        "TenderKG v1 (Mokhtari et al., Figshare 32149909.v1) represents public procurement award graphs in France. Table S2 documents the exact entity, relation, attribute, and chronological distributions determined through RDF graph traversal.")

    # Table S2: TenderKG Audit
    t_tkg = doc.add_table(rows=11, cols=3)
    format_table(t_tkg, [3.0, 1.8, 1.7],
        ["Audit Metric / Structural Property", "Exact Measured Count", "Percentage / Note"],
        [
            ["Total RDF Triples Parsed", "815,757", "11 Turtle (.ttl) files"],
            ["Total Parent Tenders", "6,126", "100.0%"],
            ["Total Procurement Lots", "29,220", "4.77 lots / tender avg."],
            ["Total Enterprise Entities (Companies)", "15,340", "100.0%"],
            ["Total Awarded Suppliers", "13,033", "84.96% of all companies"],
            ["Total Award Edges (supplier_awarded)", "36,181", "Award relationships"],
            ["Multi-Award Lots (>= 2 winners)", "3,631", "12.43% of all lots"],
            ["Single-Award Suppliers (1 total win)", "7,084", "54.35% (cold-start pool)"],
            ["Missing Lot CPV Classification Codes", "1,597", "5.47% lot missingness"],
            ["Chronological Date Anomalies (t_award < t_pub)", "137", "Quarantined (leakage risk)"]
        ],
        align_cols=[WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.LEFT])

    add_styled_heading(doc, "S2.2 Spanish BOE v3 Audit Ledger and Sentinel Distribution", 2)
    add_body_paragraph(doc,
        "The Spanish Official State Gazette (BOE v3, Zenodo 18712463) covers public tenders from 2014 to 2024. Forensic analysis revealed pervasive sentinel tokens ('No disponible', 'No Disponible') masking missing data. Table S3 provides the column-level ledger.")

    # Table S3: BOE Audit
    t_boe = doc.add_table(rows=9, cols=4)
    format_table(t_boe, [2.5, 1.3, 1.5, 1.2],
        ["Column Name", "Raw Rows", "No disponible Sentinels", "Usable Valid %"],
        [
            ["identificador (Notice ID)", "97,154", "0", "100.00%"],
            ["tipo_anuncio (Contratacion / Licitacion)", "97,154", "0", "100.00%"],
            ["nombre_adjudicatario (Supplier Name)", "97,154", "34,101", "64.90%"],
            ["importe_adjudicacion (Award Value)", "97,154", "34,101", "64.90%"],
            ["valor_estimado (Estimated Tender Value)", "97,154", "53,112", "45.33%"],
            ["codigo_cpv (CPV Taxonomy Code)", "97,154", "8,669", "91.08%"],
            ["organo_contratacion (Contracting Authority)", "97,154", "0", "100.00%"],
            ["Exact Duplicate Rows Across All Cols", "97,154", "124 duplicate rows", "Pruned"]
        ],
        align_cols=[WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.CENTER])

    add_styled_heading(doc, "S2.3 Latvian CFLA Procurement Validation Corpus Audit", 2)
    add_body_paragraph(doc,
        "The Latvian CFLA Dataset (CLARIN-LV 20.500.12574/135) assesses compliance verification in EU structural fund procurements. Analysis verified that string identifiers (e.g., '37.10' vs '37.1') must be preserved to prevent data loss. Table S4 presents the detailed question and label distributions.")

    # Table S4: Latvian Audit
    t_lat = doc.add_table(rows=7, cols=4)
    format_table(t_lat, [2.2, 1.5, 1.5, 1.3],
        ["Corpus Subset", "Main Questions", "Total Evaluated Answers", "Negative ('nē') %"],
        [
            ["Development Set (10 projects)", "61 questions", "610 answers", "1.97% (12 'nē')"],
            ["Evaluation Set (20 projects)", "61 questions", "1,220 answers", "1.89% (23 'nē')"],
            ["Headline Questionnaire (YAML)", "150 main questions", "55 conditional questions", "205 total rules"],
            ["Affirmative ('jā') Answers", "Dev + Eval (30 projects)", "1,775 answers", "97.00%"],
            ["Not Applicable ('nav attiecināms')", "Dev + Eval (30 projects)", "18 answers", "0.98%"],
            ["Unresolved ('?') Sentinels", "Dev + Eval (30 projects)", "2 answers", "0.11%"]
        ],
        align_cols=[WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.CENTER])

    # Section S3
    add_styled_heading(doc, "Section S3: Algorithmic Procedures and Implementation Schemas", 1)
    add_body_paragraph(doc,
        "The complete end-to-end processing pipeline is formalized in Algorithms S1 through S4 below, reflecting the production logic implemented in `src/models/` and `src/preprocessing/`.")

    add_styled_heading(doc, "Algorithm S1: Pre-Decision Temporal Filter and Feasible Candidate Generator", 2)
    add_code_block(doc,
"""Input: Query q published at timestamp t; Global Supplier Registry S; Historical Awards H.
Output: Feasible candidate set C(q, t).
1: t_q <- q.publication_datetime
2: H_past <- {h in H : h.award_datetime < t_q}   // Strict historical causal isolation
3: C_raw <- empty_set
4: For each supplier s in S:
5:     If s.cpv_codes intersect q.cpv_code != empty:
6:         C_raw <- C_raw union {s}
7: If |C_raw| < 50:
8:     C_history <- {h.supplier for h in H_past if h.cpv_2digit == q.cpv_2digit}
9:     C_raw <- C_raw union C_history
10: Return C_raw (bounded to max 100 candidates per query)""")

    add_styled_heading(doc, "Algorithm S2: Three-Outcome Mandatory Eligibility Evaluation", 2)
    add_code_block(doc,
"""Input: Query q; Candidate supplier s; Decision timestamp t; Rule set R_mand.
Output: Decision G(q, s, t) in {REJECT, REVIEW, ELIGIBLE}; AuditLog L.
1: If R_mand is empty:
2:     Raise ConfigurationError("Empty mandatory rule set violates safety invariant.")
3: review_reasons <- []
4: For each predicate j in R_mand:
5:     status_j <- EvaluatePredicate(j, q, s, t)
6:     If status_j == FAIL:
7:         Log(L, s, j, "MANDATORY_FAIL", t)
8:         Return (REJECT, L)   // Immediate deterministic short-circuit
9:     Else if status_j == UNKNOWN:
10:        review_reasons.append(j)
11: If |review_reasons| > 0:
12:    Log(L, s, review_reasons, "ROUTE_HUMAN_REVIEW", t)
13:    Return (REVIEW, L)
14: Log(L, s, "ALL_PASS", "CLEAR_MANDATORY_GATE", t)
15: Return (ELIGIBLE, L)""")

    add_styled_heading(doc, "Algorithm S3: Bounded Regulatory Compatibility Index (RCI)", 2)
    add_code_block(doc,
"""Input: Candidate s; Optional criteria R_opt; Weights {w_j} with sum(w_j) = 1.
Output: RCI_lower in [0, 100], RCI_upper in [0, 100], Completeness Q in [0, 1].
1: rci_lower <- 0.0
2: rci_upper <- 0.0
3: q_score <- 0.0
4: For each criterion j in R_opt:
5:     (l_j, u_j, z_j) <- GetEvidenceInterval(s, j)
6:     // Invariant: 0 <= l_j <= u_j <= 1; z_j = 1 - (u_j - l_j)
7:     rci_lower <- rci_lower + w_j * l_j
8:     rci_upper <- rci_upper + w_j * u_j
9:     q_score <- q_score + w_j * z_j
10: Return (100.0 * rci_lower, 100.0 * rci_upper, q_score)""")

    add_styled_heading(doc, "Algorithm S4: LightGBM LambdaMART Ranking and Composite Shortlisting", 2)
    add_code_block(doc,
"""Input: Eligible set F(q, t); Trained LightGBM Ranker f; ScoreMapper params (mu, T);
       Optional blending parameter beta in [0, 1]; Shortlist size k.
Output: Ranked shortlist S_k of size min(k, |F(q, t)|).
1: For each supplier s in F(q, t):
2:     x_qs <- ExtractDynamicFeatures(q, s, t)
3:     raw_score <- f.predict(x_qs)
4:     a_qs <- 1.0 / (1.0 + exp(-(raw_score - mu) / T))
5:     (rci_min, _, _) <- CalculateRCI(s, q.optional_rules)
6:     U_qs <- (1.0 - beta) * a_qs + beta * (rci_min / 100.0)
7: S_sorted <- Sort F(q, t) descending by U_qs
8: Return S_sorted[1 : min(k, |F(q, t)|)]""")

    # Section S4
    add_styled_heading(doc, "Section S4: Hyperparameter Optimization and Tuning Protocol", 1)
    add_body_paragraph(doc,
        "All model hyperparameters were tuned strictly on the chronologically partitioned validation split (1,198 parent tenders, 1,292 queries). No test partition data was accessed during tuning. Table S5 provides the complete hyperparameter search grid and selected optimal configurations.")

    # Table S5: Hyperparameters
    t_hyp = doc.add_table(rows=10, cols=4)
    format_table(t_hyp, [2.2, 1.8, 1.5, 1.0],
        ["Model / Algorithm", "Hyperparameter Description", "Exploration Grid", "Selected Optimal"],
        [
            ["LightGBM LambdaRank", "learning_rate (eta)", "[0.01, 0.03, 0.05, 0.1]", "0.05"],
            ["LightGBM LambdaRank", "num_leaves", "[15, 31, 63, 127]", "31"],
            ["LightGBM LambdaRank", "min_child_samples", "[10, 20, 50, 100]", "20"],
            ["LightGBM LambdaRank", "subsample (bagging_fraction)", "[0.7, 0.8, 0.9, 1.0]", "0.8"],
            ["LightGBM LambdaRank", "colsample_bytree (feature_fraction)", "[0.6, 0.7, 0.8, 1.0]", "0.8"],
            ["LightGBM LambdaRank", "early_stopping_rounds", "[20, 30, 50]", "50 (stopped at 23)"],
            ["Sigmoid ScoreMapper", "Centering parameter mu", "[-2.0, -1.0, -0.74, 0.0]", "-0.7463"],
            ["Sigmoid ScoreMapper", "Temperature parameter T", "[0.5, 1.0, 1.5, 2.0]", "1.0000"],
            ["Utility Blending Weight", "Compliance weight beta", "[0.0, 0.1, 0.2, 0.3, 0.5]", "0.2000"]
        ],
        align_cols=[WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER])

    # Section S5
    add_styled_heading(doc, "Section S5: Industrial Manufacturing Case Study Walk-Through", 1)
    add_body_paragraph(doc,
        "To illustrate the concrete operational mechanics and failure-prevention capabilities of the proposed architecture, we present an end-to-end case walk-through based on a precision CNC aerospace structural component procurement request.")

    add_styled_heading(doc, "S5.1 Procurement Request Specifications", 2)
    add_body_paragraph(doc,
        "A defense aerospace contracting authority issues Tender AER-2026-9041 for 5-axis titanium milling of turbine housing bulkheads (CPV 45223200-8, Structural components). The statutory rules stipulate two mandatory qualification requirements: (1) Mandatory active enterprise commercial registration verified against the national corporate registry (SIREN/Sirene); (2) Mandatory possession of a valid AS9100 / ISO 9001 quality management certification. In addition, two optional preference criteria are declared: (1) Demonstrated ISO 14001 Environmental Management certification (weight $w_1 = 0.60$); (2) Past on-time delivery track record $\\ge 95\\%$ (weight $w_2 = 0.40$).")

    add_styled_heading(doc, "S5.2 Competing Bidder Profiles", 2)
    add_body_paragraph(doc,
        "Four candidate enterprises submit proposals or are retrieved by the candidate generation index:")
    add_body_paragraph(doc,
        "Bidder A (Incumbent Market Leader): Holds 15 prior contract awards in structural milling; highest historical category volume (score $f = +2.45$). However, its ISO 9001 certificate expired 4 months prior to publication and was not renewed ($r_{\\text{mand},2} = \\text{FAIL}$). Active SIREN verified.",
        bold_prefix="  - ")
    add_body_paragraph(doc,
        "Bidder B (Certified Mid-Tier Supplier): Holds 3 prior contract awards in related manufacturing (score $f = +0.85$). Holds active verified SIREN and current ISO 9001 / AS9100 certificate ($r_{\\text{mand},1} = \\text{PASS}$, $r_{\\text{mand},2} = \\text{PASS}$). Submits verified ISO 14001 certificate ($l_1 = 1, u_1 = 1$) and certified on-time rate of 98% ($l_2 = 1, u_2 = 1$).",
        bold_prefix="  - ")
    add_body_paragraph(doc,
        "Bidder C (Cold-Start Innovator): Newly formed aerospace startup with 0 prior public contract awards (score $f = -0.15$). Active verified SIREN and valid ISO 9001 certificate ($r_{\\text{mand},1} = \\text{PASS}$, $r_{\\text{mand},2} = \\text{PASS}$). Environmental certificate application is in progress but unverified ($l_1 = 0, u_1 = 1$); delivery record unobserved ($l_2 = 0, u_2 = 1$).",
        bold_prefix="  - ")
    add_body_paragraph(doc,
        "Bidder D (Debarred Entity): Holds 6 historical awards; active SIREN; but flagged for unresolved statutory debarment / compliance infraction ($r_{\\text{mand},1} = \\text{FAIL}$).",
        bold_prefix="  - ")

    add_styled_heading(doc, "S5.3 Comparative Step-by-Step Execution Trace", 2)
    add_body_paragraph(doc,
        "Table S6 contrasts the ranking decisions of the unconstrained baseline LambdaMART against the proposed Compliance-Constrained Architecture.")

    # Table S6: Industrial Case Trace
    t_case = doc.add_table(rows=5, cols=5)
    format_table(t_case, [1.5, 1.3, 1.8, 1.4, 1.5],
        ["Bidder Entity", "Learned Score f", "Unconstrained Rank", "Mandatory Gate Decision", "Proposed System Outcome"],
        [
            ["Bidder A (Incumbent)", "+2.45 (Highest)", "Rank #1 (Awarded)", "REJECT (Expired ISO)", "Excluded from Shortlist (0% FC)"],
            ["Bidder B (Certified)", "+0.85 (Moderate)", "Rank #2", "ELIGIBLE (All PASS)", "Rank #1 (U = 0.865) -> Shortlisted"],
            ["Bidder C (Cold-Start)", "-0.15 (Low)", "Rank #4", "ELIGIBLE (All PASS)", "Rank #2 (U = 0.518) -> Shortlisted"],
            ["Bidder D (Debarred)", "+1.10 (High)", "Rank #3", "REJECT (Debarment FAIL)", "Excluded from Shortlist (0% FC)"]
        ],
        align_cols=[WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT])

    add_body_paragraph(doc,
        "Under the unconstrained LambdaMART baseline, Bidder A is awarded the contract because its 15 past contract wins generate high feature values across category_awards, total_revenue, and graph degree centrality. However, awarding the contract to Bidder A is strictly illegal under defense aerospace regulations due to expired certification. In sharp contrast, the proposed architecture's mandatory gate immediately evaluates $r_{\\text{mand},2} = \\text{FAIL}$, deterministically excluding Bidder A prior to ranking. Bidder B passes all mandatory gates, achieves an $\\text{RCI}^-$ of 100.0, and is correctly ranked at Rank #1. Bidder C, despite having 0 prior awards, is safely retained in the eligible candidate pool and successfully awarded Rank #2.")

    # Section S6
    add_styled_heading(doc, "Section S6: Comprehensive Data Dictionary and Feature Lineage", 1)
    add_body_paragraph(doc,
        "Tables S7 and S8 record the complete data dictionary and feature lineage specifications governing all extracted inputs.")

    # Table S7: Data Dictionary
    t_dict = doc.add_table(rows=8, cols=4)
    format_table(t_dict, [1.8, 1.4, 2.3, 1.0],
        ["Field Name", "Data Type", "Semantic Description & Measurement Unit", "Source Table"],
        [
            ["request_id", "String (UUID)", "Unique parent tender identifier", "requests.parquet"],
            ["lot_id", "String (UUID)", "Specific procurement contract lot identifier", "requests.parquet"],
            ["supplier_id", "String (9-digit)", "National INSEE SIREN enterprise registration code", "suppliers.parquet"],
            ["publication_date", "ISO-8601 Timestamp", "Public tender announcement date (decision time t)", "requests.parquet"],
            ["cpv_code", "String (8-digit)", "EU Common Procurement Vocabulary taxonomy classification", "requests.parquet"],
            ["target_awarded", "Binary {0, 1}", "Observed ground-truth commercial award outcome", "outcomes.parquet"],
            ["siren_valid", "Binary {0, 1}", "Deterministic validation of 9-digit SIREN registration", "suppliers.parquet"]
        ],
        align_cols=[WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT])

    # Table S8: Feature Lineage
    t_lin = doc.add_table(rows=8, cols=4)
    format_table(t_lin, [2.0, 1.3, 2.5, 0.9],
        ["Feature Name", "Feature Type", "Exact Mathematical Derivation & Causal Restriction", "Pre-Decision?"],
        [
            ["supplier_past_awards", "Integer >= 0", "count({h in H : h.supplier = s and h.award_date < t})", "YES (Strict)"],
            ["supplier_cpv_awards", "Integer >= 0", "count({h in H : h.supplier = s, h.cpv = q.cpv, h.award_date < t})", "YES (Strict)"],
            ["cpv_exact_match", "Binary {0, 1}", "1{q.cpv in s.cpv_codes_registered}", "YES (Static)"],
            ["cpv_division_match", "Binary {0, 1}", "1{q.cpv[0:2] in s.cpv_divisions_registered}", "YES (Static)"],
            ["buyer_supplier_prior_awards", "Integer >= 0", "count({h in H : h.buyer = q.buyer, h.supplier = s, h.date < t})", "YES (Strict)"],
            ["tfidf_similarity", "Float in [0, 1]", "cosine_sim(TFIDF(q.lot_desc), TFIDF(s.activity_desc))", "YES (Static)"],
            ["degree_centrality", "Float >= 0", "Dynamic in-degree centrality in past co-bidding graph at t", "YES (Strict)"]
        ],
        align_cols=[WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER])

    doc_path = os.path.join(DELIVERABLES_DIR, "supplementary_methods.docx")
    doc.save(doc_path)
    print(f"Saved Supplementary Methods document to {doc_path}")

    # Convert to PDF via Word COM Automation
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        abs_doc_path = os.path.abspath(doc_path)
        pdf_path = os.path.splitext(abs_doc_path)[0] + ".pdf"
        wb = word.Documents.Open(abs_doc_path)
        wb.SaveAs(pdf_path, FileFormat=17) # 17 = wdFormatPDF
        wb.Close()
        word.Quit()
        print(f"Successfully converted Supplementary Methods to PDF: {pdf_path}")
    except Exception as e:
        print(f"Notice: Word COM PDF conversion encountered an error: {e}")

if __name__ == "__main__":
    build_supplementary_docx()
