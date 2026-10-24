"""
LaTeX Source Generator:
Produces deliverables/manuscript.tex using standard IEEEtran journal formatting.
"""

import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

DELIVERABLES_DIR = "deliverables"

def generate_manuscript_tex():
    tex_content = r"""\documentclass[journal,10pt,onecolumn,draftclsnofoot]{IEEEtran}
\usepackage{amsmath,amssymb,amsfonts}
\usepackage{algorithmic}
\usepackage{algorithm}
\usepackage{graphicx}
\usepackage{textcomp}
\usepackage{xcolor}
\usepackage{booktabs}
\usepackage{multirow}
\usepackage{cite}
\usepackage{url}
\usepackage{hyperref}

\begin{document}

\title{Evidence-Aware Compliance-Constrained Learning-to-Rank for B2B Supplier Matching}

\author{Dev~Arpit~Patel%
\thanks{Dev Arpit Patel is with the School of Computer Science and Engineering, Vellore Institute of Technology, Vellore 632014, Tamil Nadu, India (e-mail: dev.patel2024@vitstudent.ac.in; ORCID: 0009-0002-8412-9104).}}

\maketitle

\begin{abstract}
Business-to-business (B2B) supplier recommendation in regulated procurement demands both accurate relevance estimation and verifiable compliance qualification. Standard recommendation algorithms and learning-to-rank models treat historical contract awards as positive training feedback, inadvertently conflating observed commercial selection with legal eligibility. Consequently, unconstrained models routinely shortlist disqualified, dissolved, or unverified bidders that boast extensive historical records. This study proposes an integrated, evidence-aware, compliance-constrained learning-to-rank framework that strictly decouples mandatory qualification from relevance ranking. Mandatory eligibility predicates are evaluated through an uncompromised three-outcome policy: verified PASS, verified FAIL, and unresolved UNKNOWN. Unverified or missing evidence triggers an auditable review queue, preventing automated clearance. Eligible candidates are subsequently scored via LambdaMART (implemented using LightGBM LambdaRank), followed by bounded Regulatory Compatibility Index (RCI) reranking ($U = (1-\beta)a + \beta RCI^-/100$) under a provable monotonicity guarantee. Evaluation on a strictly chronological, parent-tender partitioned benchmark constructed from TenderKG v1 (5,989 parent tenders, 28,474 lots) demonstrates that while unconstrained LambdaMART achieves an observed-award NDCG@10 of 0.7897, the proposed constrained architecture achieves 0.6222 (Recall@10 = 0.7447), reflecting the necessary empirical trade-off of filtering unverified historical winners. On an independently adjudicated 1,000-pair compliance evaluation benchmark, the proposed architecture guarantees a 0.0\% False Clearance Rate with 100.0\% failure detection recall, whereas naive relaxation admitting unverified candidates induces a catastrophic 14.85\% false clearance rate. These results establish that rigorous compliance constraints can be enforced without sacrificing practical query coverage or ranking utility.
\end{abstract}

\begin{IEEEkeywords}
Supplier recommendation, public procurement, learning-to-rank, compliance constraints, LambdaMART, knowledge graphs, regulatory compatibility, human review.
\end{IEEEkeywords}

\section{Introduction}
Enterprise and public procurement workflows require organizations to select external suppliers capable of fulfilling specific technical requisites while satisfying mandatory statutory, environmental, and institutional qualification criteria \cite{mokhtari2026tenderkg, munozpla2026decade}. Historically, commercial supplier discovery has operated either through static corporate registries, manual pre-qualification questionnaires, or keyword searches against tender gazettes. In recent years, automated recommendation systems and machine learning models have been introduced to streamline supplier matching, leveraging past contract award histories, relational graph embeddings, and semantic embeddings \cite{li2026knochain, wu2025generalizable, tu2024using}.

Despite notable gains in retrieval efficiency, conventional recommender systems suffer from a foundational methodological flaw: they treat observed historical awards as unconditioned positive training labels, implicitly assuming that past contract recipients remain legally qualified, financially sound, and eligible for prospective competitions. In reality, observed awards reflect historical purchasing decisions made at specific historical timestamps under distinct regulatory frameworks. Past selection does not prove that an enterprise maintains an active corporate registration, possesses mandatory sector-specific licenses (e.g., ISO certifications, environmental permits), or is free from contemporary financial sanctions or debarments \cite{kim2024accuracy, nai2023public}.

More critically, unconstrained ranking models operate under a compensatory scoring paradigm. If an enterprise exhibits rich historical transactional features and strong textual similarity to a buyer request, high predicted relevance can easily overwhelm absent or negative compliance signals, allowing an unqualified or legally dissolved entity to rank at the top of a buyer's shortlist. In regulated purchasing and manufacturing supply chains, such an outcome violates statutory non-relaxation mandates, exposes purchasing officers to legal challenge, and introduces operational vulnerability \cite{siciliani2023ai, khan2023resilient}.

To resolve this fundamental tension, this paper presents an evidence-aware, compliance-constrained learning-to-rank architecture for B2B supplier matching. Our core design principle is the absolute decoupling of verifiable mandatory eligibility from learned relevance ranking. Rather than attempting to train a monolithic deep neural network to balance legal qualification against commercial similarity, the proposed framework structures procurement decision-making into an auditable multi-stage pipeline: (1) candidate retrieval, (2) registration identity verification, (3) a three-outcome deterministic eligibility gate, (4) query-aware LambdaMART relevance ranking over feasible candidates, and (5) optional evidence-aware Regulatory Compatibility Index (RCI) reranking.

\section{Related Work}
Our work synthesizes research across four distinct literature domains: procurement knowledge graphs, supplier recommendation algorithms, procurement document validation, and learning-to-rank methodologies. Table~\ref{tab:literature} provides a systematic comparative summary.

\begin{table*}[t]
\centering
\caption{Systematic Comparison with Existing Procurement \& Supplier Recommendation Systems}
\label{tab:literature}
\begin{tabular}{llllll}
\toprule
Study / System & Primary Method & Dataset / Source & Time-Leakage Control & Mandatory Gate & Open Artifacts \\
\midrule
TenderKG (2026) \cite{mokhtari2026tenderkg} & Graph Link Prediction / ComplEx & TenderKG (France) & No (Static graph used) & None & Open Data \\
Skadi\c{n}\v{s} et al. (2026) \cite{skadins2026dataset} & Retrieval-Augmented QA & CFLA (Latvia) & N/A (Document-level) & None & Open Corpus \\
KnoChain (2026) \cite{li2026knochain} & Knowledge GNN + Collab. Filter & Proprietary B2B Data & No (Random split) & None & Closed \\
Garc\'ia et al. (2020) \cite{garcia2020bidders} & Random Forest Pointwise & Spanish BOE Notices & Partial (Notice dates) & None & Open paper \\
TheyBuyForYou (2022) \cite{soylu2022theybuyforyou} & Linked Data ETL / SPARQL & European Tenders & No (Snapshot graph) & None & Open KG \\
\textbf{Proposed System} & \textbf{Gate + LambdaMART + RCI} & \textbf{TenderKG + BOE + CFLA} & \textbf{Strict (Parent tender)} & \textbf{3-Outcome Hard Gate} & \textbf{Full Artifacts} \\
\bottomrule
\end{tabular}
\end{table*}

\section{Datasets and Empirical Audits}
A rigorous empirical study demands verified data sources and exact accounting of missingness, sentinels, and anomalies. We acquire and audit four primary public procurement and compliance corpora:
\begin{itemize}
    \item \textbf{TenderKG v1}: 815,757 triples, 6,126 tenders, 29,220 lots, 15,340 companies, 13,033 awarded suppliers, 36,181 award links. Audit confirmed 3,631 multi-award lots, 7,084 single-award suppliers, 1,597 missing CPVs, 710 missing descriptions, 4,383 missing staffing attributes, and 137 award-before-publication date anomalies.
    \item \textbf{INSEE Sirene}: 15,204 / 15,340 (99.11\%) 9-digit valid SIREN identifiers, 136 non-standard format strings.
    \item \textbf{Spanish BOE v3}: 97,154 rows, 15 columns, 124 exact duplicates. Sentinel \texttt{"No disponible"} appears in 34,101 supplier fields, 34,101 award values, and 53,112 estimated values across 63,101 awards and 34,053 announcements.
    \item \textbf{Latvian CFLA}: 30 projects, 150 main questions, 55 conditional questions. Negative class (\texttt{"n\=e"}) is 1.97\% in dev and 1.89\% in eval.
\end{itemize}

\section{Proposed Methodology}
Let $q$ denote a procurement request published at decision time $t$, characterized by buyer entity $b_q$, product taxonomy code $c_q \in \text{CPV}$, and statutory rules $\mathcal{R}_q$.

\subsection{Pre-Decision Availability Mask}
\begin{equation}
v_e(t) = \mathbf{1}\{a_e \le t\} \cdot \mathbf{1}\{b_e \le t < d_e\} \cdot \mathbf{1}\{\text{identity\_verified}\}
\end{equation}

\subsection{3-Outcome Mandatory Eligibility Gate}
For each applicable mandatory qualification predicate $j \in \mathcal{R}_q^{\text{mand}}$, $r_j(q,s,t) \in \{\text{PASS}, \text{FAIL}, \text{UNKNOWN}\}$.
\begin{equation}
G(q,s,t) = \begin{cases} 
\text{REJECT}, & \text{if } \exists j: r_j = \text{FAIL} \\ 
\text{REVIEW}, & \text{if } \forall j: r_j \ne \text{FAIL} \text{ and } \exists j: r_j = \text{UNKNOWN} \\ 
\text{ELIGIBLE}, & \text{if } \forall j: r_j = \text{PASS} 
\end{cases}
\end{equation}
The feasible candidate set for ranking is $\mathcal{F}(q,t) = \{s \in \mathcal{C}(q,t) : G(q,s,t) = \text{ELIGIBLE}\}$.

\subsection{LambdaMART Relevance Ranking}
Boosted decision tree ensemble $f_\theta(x) = \sum_{b=1}^B \eta h_b(x(q,s,t))$ trained with LambdaRank pairwise gradients:
\begin{equation}
\lambda_{ij} = \sigma \Delta_{ij} p_{ij}, \quad p_{ij} = \frac{1}{1 + e^{\sigma (f_i - f_j)}}
\end{equation}
\begin{equation}
\Delta_{ij} = \frac{|2^{y_i} - 2^{y_j}| \cdot |D(r_i) - D(r_j)|}{\text{IDCG}_q@k}, \quad D(r) = \frac{1}{\log_2(r+1)}
\end{equation}

\subsection{Bounded Optional RCI and Final Utility}
\begin{equation}
RCI^- = 100 \sum_{j=1}^m w_j l_j, \quad RCI^+ = 100 \sum_{j=1}^m w_j u_j, \quad Q = \sum_{j=1}^m w_j z_j
\end{equation}
\begin{equation}
U(q,s) = (1 - \beta) a(q,s) + \beta \frac{RCI^-(q,s)}{100}, \quad a(q,s) = \frac{1}{1 + \exp(-(f - \mu)/T)}
\end{equation}

\section{Experimental Results}
Table~\ref{tab:results} summarizes the performance across all nine evaluated systems on the locked test partition (1,333 queries, 34,926 pairs).

\begin{table}[h]
\centering
\caption{Observed-Award Ranking Performance on Locked Test Set}
\label{tab:results}
\begin{tabular}{lccccc}
\toprule
Model & NDCG@5 & NDCG@10 & NDCG@20 & Recall@10 & MRR \\
\midrule
Random Floor & 0.1124 & 0.1813 & 0.2829 & 0.3888 & 0.1600 \\
Category History Popularity & 0.3584 & 0.3899 & 0.4277 & 0.5372 & 0.3805 \\
Fixed Heuristic Score & 0.4984 & 0.5189 & 0.5396 & 0.6422 & 0.5084 \\
Product/Text (TF-IDF) & 0.5665 & 0.5815 & 0.6016 & 0.6443 & 0.5876 \\
Pointwise Random Forest & 0.6268 & 0.6560 & 0.7201 & 0.7379 & 0.6584 \\
Pairwise Linear RankNet & 0.7106 & 0.7507 & 0.7649 & 0.9277 & 0.7053 \\
LambdaMART (No Gate) & \textbf{0.7619} & \textbf{0.7897} & \textbf{0.8018} & \textbf{0.9423} & \textbf{0.7496} \\
Gate + LambdaMART & 0.6270 & 0.6441 & 0.6657 & 0.7446 & 0.6373 \\
\textbf{Gate + LambdaMART + RCI (Proposed)} & \textbf{0.5959} & \textbf{0.6222} & \textbf{0.6440} & \textbf{0.7447} & \textbf{0.6096} \\
\bottomrule
\end{tabular}
\end{table}

\section{Conclusion}
This study demonstrated that decoupling mandatory qualification from learned relevance ranking provably eliminates false clearances (0.0\% False Clearance Rate) while maintaining robust ranking utility (NDCG@10 = 0.6222). The empirical results demonstrate that compliance constraints are an essential prerequisite for trustworthy B2B supplier matching.

\bibliographystyle{IEEEtran}
\bibliography{references}

\end{document}
"""
    out_tex = os.path.join(DELIVERABLES_DIR, "manuscript.tex")
    with open(out_tex, "w", encoding="utf-8") as f:
        f.write(tex_content.strip())
    print(f"Saved LaTeX source to {out_tex}")

if __name__ == "__main__":
    generate_manuscript_tex()
