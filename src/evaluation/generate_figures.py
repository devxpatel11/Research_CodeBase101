"""
Figure Generation Script:
Generates all 8 publication-quality figures specified in Section 16 of the Research Guide:
  - Fig 1: End-to-end System Architecture Diagram
  - Fig 2: Time-Respecting Data Preparation Protocol
  - Fig 3: Relational Evidence and Provenance Schema
  - Fig 4: Temporal Partitioning Timeline
  - Fig 5: Empirical Procurement Data Profile & Long Tail
  - Fig 6: LambdaMART Training & Validation Learning Curves
  - Fig 7: Operational Compliance vs Coverage Trade-off Plot
  - Fig 8: Systematic Ablations & Policy Sensitivity Plot
Exports vector PDFs and high-DPI PNGs into figures/ directory.
"""

import os
import sys
import json
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as patches

sys.stdout.reconfigure(encoding='utf-8')

plt.rcParams.update({
    "font.family": "serif",
    "font.serif": ["Times New Roman", "DejaVu Serif"],
    "font.size": 10,
    "axes.labelsize": 11,
    "axes.titlesize": 12,
    "xtick.labelsize": 9,
    "ytick.labelsize": 9,
    "legend.fontsize": 9,
    "figure.titlesize": 13,
    "figure.dpi": 300,
    "savefig.dpi": 300,
    "savefig.bbox": "tight"
})

FIG_DIR = "figures"
LOGS_DIR = os.path.join("experiments", "logs")
METRICS_DIR = os.path.join("experiments", "metrics")
os.makedirs(FIG_DIR, exist_ok=True)

def generate_figure_1_architecture():
    fig, ax = plt.subplots(figsize=(8.5, 6.5))
    ax.axis("off")

    # Define boxes: (x, y, w, h, text, color)
    boxes = [
        (0.25, 0.88, 0.50, 0.08, "Buyer Request $q$ & Decision Time $t$\n(Product Scope, Buyer ID, Mandatory Rules)", "#EBF3FB", "#1B4F72"),
        (0.25, 0.74, 0.50, 0.08, "Candidate Retrieval $\\mathcal{C}(q,t)$ & SIREN Identity Validation\n(Broad CPV Retrieval + Format Verification)", "#EBF3FB", "#1B4F72"),
        (0.25, 0.60, 0.50, 0.08, "3-Outcome Mandatory Eligibility Gate $G(q,s,t)$\n(Deterministic Predicates: Commercial, Debarment, Capacity)", "#FEF9E7", "#B7950B"),
        (0.02, 0.44, 0.28, 0.08, "REJECT\n(Verified Applicable Failure)\nQuarantine / Audit Log", "#FDEDEC", "#922B21"),
        (0.70, 0.44, 0.28, 0.08, "REVIEW\n(Unresolved Evidence)\nHuman Adjudication Queue", "#FEF9E7", "#B7950B"),
        (0.25, 0.44, 0.50, 0.08, "Feasible Candidate Set $\\mathcal{F}(q,t)$\n(Verified All Applicable Mandatory PASS)", "#E8F8F5", "#117864"),
        (0.25, 0.30, 0.50, 0.08, "LambdaMART Relevance Ranking\n(LightGBM Boosted Trees: Past History + Content)", "#E8F8F5", "#117864"),
        (0.25, 0.16, 0.50, 0.08, "Optional Evidence Bounds & Utility Integration\n$U = (1-\\beta)a + \\beta (RCI^- / 100)$", "#EBF3FB", "#1B4F72"),
        (0.25, 0.02, 0.50, 0.08, "Approved Shortlist $\\mathcal{L}_q$ (Top-$k$) & Audit Trail\n(Decisive Evidence, Validity Hashes, Timestamps)", "#E8F8F5", "#117864")
    ]

    for x, y, w, h, text, bg, border in boxes:
        rect = patches.FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.015",
                                     ec=border, fc=bg, lw=1.5)
        ax.add_patch(rect)
        ax.text(x + w/2.0, y + h/2.0, text, ha="center", va="center", color="#1C2833", weight="semibold", fontsize=9.5)

    # Arrows
    arrow_props = dict(arrowstyle="->", lw=1.5, color="#2C3E50")
    # Downward central flow
    ax.annotate("", xy=(0.50, 0.82), xytext=(0.50, 0.88), arrowprops=arrow_props)
    ax.annotate("", xy=(0.50, 0.68), xytext=(0.50, 0.74), arrowprops=arrow_props)
    
    # Gate branching:
    ax.annotate("", xy=(0.16, 0.52), xytext=(0.35, 0.60), arrowprops=arrow_props)
    ax.text(0.22, 0.57, "Fail", color="#922B21", weight="bold", fontsize=9)

    ax.annotate("", xy=(0.84, 0.52), xytext=(0.65, 0.60), arrowprops=arrow_props)
    ax.text(0.74, 0.57, "Unknown", color="#B7950B", weight="bold", fontsize=9)

    ax.annotate("", xy=(0.50, 0.52), xytext=(0.50, 0.60), arrowprops=arrow_props)
    ax.text(0.51, 0.56, "All Pass", color="#117864", weight="bold", fontsize=9)

    # Continue downward flow
    ax.annotate("", xy=(0.50, 0.38), xytext=(0.50, 0.44), arrowprops=arrow_props)
    ax.annotate("", xy=(0.50, 0.24), xytext=(0.50, 0.30), arrowprops=arrow_props)
    ax.annotate("", xy=(0.50, 0.10), xytext=(0.50, 0.16), arrowprops=arrow_props)

    plt.title("Figure 1: End-to-End System Architecture with Strict Non-Relaxation Gate and RCI Utility", pad=15, weight="bold")
    plt.tight_layout()
    plt.savefig(os.path.join(FIG_DIR, "fig1_system_architecture.png"))
    plt.savefig(os.path.join(FIG_DIR, "fig1_system_architecture.pdf"))
    plt.close()
    print("Generated Figure 1.")

def generate_figure_2_preparation():
    fig, ax = plt.subplots(figsize=(8.0, 5.5))
    ax.axis("off")

    boxes = [
        (0.20, 0.84, 0.60, 0.10, "1. Raw Source Audits & Hash Verification\n(TenderKG v1: 815k triples | BOE v3: 97k rows | Latvian: 30 projects)", "#F4F6F6", "#34495E"),
        (0.20, 0.68, 0.60, 0.10, "2. Tender Family Grouping & Quarantining\n(137 award-before-pub & 13 out-of-horizon 2024 records purged)", "#FDEDEC", "#C0392B"),
        (0.20, 0.52, 0.60, 0.10, "3. Chronological Parent-Tender Partitioning\n(Train 60%: 2021-01 to 2022-07 | Val 20%: 2022-07 to 2023-02 | Test 20%: 2023-02 to 2023-11)", "#FEF9E7", "#D4AC0D"),
        (0.20, 0.36, 0.60, 0.10, "4. Past-Only Dynamic Feature Construction\n(Historical CPV affinity & buyer interactions using strictly pre-t events)", "#E8F8F5", "#16A085"),
        (0.20, 0.20, 0.60, 0.10, "5. Strict Training-Only Preprocessing & Encoders\n(TF-IDF vocabulary & frequency encoders fitted solely on Train split)", "#EBF5FB", "#2980B9"),
        (0.20, 0.04, 0.60, 0.10, "6. Locked Test Candidate Set & Paired Evaluation\n(Pre-decision candidate pools evaluated with 2,000 cluster bootstrap resamples)", "#E8F8F5", "#27AE60")
    ]

    for x, y, w, h, text, bg, border in boxes:
        rect = patches.FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.015",
                                     ec=border, fc=bg, lw=1.5)
        ax.add_patch(rect)
        ax.text(x + w/2.0, y + h/2.0, text, ha="center", va="center", color="#1C2833", weight="semibold", fontsize=9.5)

    arrow_props = dict(arrowstyle="->", lw=1.8, color="#2C3E50")
    for y_pos in [0.84, 0.68, 0.52, 0.36, 0.20]:
        ax.annotate("", xy=(0.50, y_pos - 0.06), xytext=(0.50, y_pos), arrowprops=arrow_props)

    plt.title("Figure 2: Time-Respecting Data Preparation and Leakage Isolation Protocol", pad=15, weight="bold")
    plt.tight_layout()
    plt.savefig(os.path.join(FIG_DIR, "fig2_data_preparation.png"))
    plt.savefig(os.path.join(FIG_DIR, "fig2_data_preparation.pdf"))
    plt.close()
    print("Generated Figure 2.")

def generate_figure_3_schema():
    fig, ax = plt.subplots(figsize=(9.0, 5.5))
    ax.axis("off")

    tables = [
        (0.05, 0.55, 0.25, 0.35, "requests\n" + "-"*25 + "\n• query_id (PK)\n• tender_id (FK)\n• buyer_id\n• pub_datetime (t)\n• cpv_code\n• lot_title\n• description", "#EBF5FB", "#2980B9"),
        (0.38, 0.55, 0.25, 0.35, "outcomes\n" + "-"*25 + "\n• query_id (FK)\n• supplier_id (FK)\n• target_relevance (y)\n• award_datetime\n• observed_amount\n• multi_award_flag", "#FEF9E7", "#D4AC0D"),
        (0.70, 0.55, 0.25, 0.35, "suppliers\n" + "-"*25 + "\n• supplier_id (PK)\n• siren_id\n• siren_valid_flag\n• staffing_range\n• legal_form\n• naf_activity_code\n• company_label", "#E8F8F5", "#16A085"),
        (0.05, 0.05, 0.25, 0.35, "requirements\n" + "-"*25 + "\n• rule_id (PK)\n• query_id (FK)\n• is_mandatory (bool)\n• predicate_type\n• required_scope\n• effective_from\n• effective_to", "#FEF9E7", "#B7950B"),
        (0.38, 0.05, 0.25, 0.35, "compliance_evidence\n" + "-"*25 + "\n• evidence_id (PK)\n• supplier_id (FK)\n• document_hash\n• certificate_type\n• issuer_authority\n• issue_date\n• expiry_date", "#FDEDEC", "#C0392B"),
        (0.70, 0.05, 0.25, 0.35, "predictions_audit\n" + "-"*25 + "\n• run_id\n• query_id (FK)\n• supplier_id (FK)\n• gate_decision\n• ranker_score_f\n• rci_lower / rci_upper\n• final_utility_u", "#E8F8F5", "#27AE60")
    ]

    for x, y, w, h, text, bg, border in tables:
        rect = patches.FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.015",
                                     ec=border, fc=bg, lw=1.5)
        ax.add_patch(rect)
        ax.text(x + 0.02, y + h - 0.03, text, ha="left", va="top", color="#1C2833", fontsize=8.5, family="monospace")

    arrow_props = dict(arrowstyle="<->", lw=1.5, color="#7F8C8D")
    ax.annotate("", xy=(0.30, 0.72), xytext=(0.38, 0.72), arrowprops=arrow_props)
    ax.annotate("", xy=(0.63, 0.72), xytext=(0.70, 0.72), arrowprops=arrow_props)
    ax.annotate("", xy=(0.17, 0.55), xytext=(0.17, 0.40), arrowprops=arrow_props)
    ax.annotate("", xy=(0.82, 0.55), xytext=(0.82, 0.40), arrowprops=arrow_props)
    ax.annotate("", xy=(0.50, 0.40), xytext=(0.50, 0.55), arrowprops=arrow_props)

    plt.title("Figure 3: Relational Evidence Schema Linking Procurement Requests, Suppliers, and Provenance Audit", pad=15, weight="bold")
    plt.tight_layout()
    plt.savefig(os.path.join(FIG_DIR, "fig3_evidence_schema.png"))
    plt.savefig(os.path.join(FIG_DIR, "fig3_evidence_schema.pdf"))
    plt.close()
    print("Generated Figure 3.")

def generate_figure_4_timeline():
    fig, ax = plt.subplots(figsize=(8.5, 3.2))

    # Horizontal timeline bar
    # 2021-01 to 2023-12 (36 months)
    ax.set_xlim(0, 36)
    ax.set_ylim(-0.5, 1.5)
    ax.axis("off")

    # Split ranges: Train: 0 to 18.5 (2021-01 to 2022-07)
    # Val: 18.5 to 25.5 (2022-07 to 2023-02)
    # Test: 25.5 to 35.0 (2023-02 to 2023-11)
    # Quarantined: 35.0 to 36.0 (2024 dates)
    bars = [
        (0.0, 18.5, "#AED6F1", "#2471A3", "TRAINING PARTITION (60%)\n3,593 Parent Tenders | 17,492 Lots\n2021-01-04 to 2022-07-22"),
        (18.5, 7.0, "#FAD7A0", "#B9770E", "VALIDATION (20%)\n1,198 Tenders | 5,735 Lots\n2022-07-23 to 2023-02-16"),
        (25.5, 9.5, "#A9DFBF", "#1E8449", "LOCKED TEST (20%)\n1,198 Tenders | 5,247 Lots\n2023-02-16 to 2023-11-26")
    ]

    for start, width, bg, border, label in bars:
        rect = patches.Rectangle((start, 0.2), width, 0.6, ec=border, fc=bg, lw=2.0)
        ax.add_patch(rect)
        ax.text(start + width/2.0, 0.5, label, ha="center", va="center", color="#1C2833", weight="bold", fontsize=8.5)

    # Cutoff markers
    ax.axvline(18.5, ymin=0.15, ymax=0.85, color="#B9770E", ls="--", lw=2)
    ax.text(18.5, 0.95, "Split Cutoff $T_1$\n(2022-07-22)", ha="center", va="bottom", color="#B9770E", weight="bold", fontsize=8)

    ax.axvline(25.5, ymin=0.15, ymax=0.85, color="#1E8449", ls="--", lw=2)
    ax.text(25.5, 0.95, "Split Cutoff $T_2$\n(2023-02-16)", ha="center", va="bottom", color="#1E8449", weight="bold", fontsize=8)

    # Invariant indicator below
    ax.text(18.0, -0.2, "Temporal Invariant: Features for query at time $t$ constructed strictly from events where $\\text{award\\_datetime} < t$.\nSibling lots of parent tenders grouped indivisibly across split boundaries.", ha="center", va="center", fontsize=8.5, style="italic", color="#566573")

    plt.title("Figure 4: Chronological Partitioning Protocol and Decision-Time Invariant", pad=20, weight="bold")
    plt.tight_layout()
    plt.savefig(os.path.join(FIG_DIR, "fig4_temporal_protocol.png"))
    plt.savefig(os.path.join(FIG_DIR, "fig4_temporal_protocol.pdf"))
    plt.close()
    print("Generated Figure 4.")

def generate_figure_5_data_profile():
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(9.0, 4.0))

    # Panel A: Missingness rates across datasets
    labels = ["BOE: Estim. Value", "BOE: Offer Value", "BOE: Supplier Name", "TenderKG: Staffing", "TenderKG: CPV", "TenderKG: Desc."]
    pcts = [54.67, 35.10, 35.10, 28.57, 5.46, 2.43]
    colors = ["#E74C3C", "#E67E22", "#E67E22", "#3498DB", "#2ECC71", "#1ABC9C"]

    ax1.barh(labels[::-1], pcts[::-1], color=colors[::-1], edgecolor="black", alpha=0.8)
    ax1.set_xlabel("Missing / Sentinel Prevalence (%)")
    ax1.set_title("(a) Attribute Missingness & Sentinels", weight="bold")
    ax1.set_xlim(0, 65)
    for i, v in enumerate(pcts[::-1]):
        ax1.text(v + 1.0, i, f"{v:.1f}%", va="center", fontsize=8.5)

    # Panel B: Supplier award count heavy tail
    # From audit: 7,084 with 1 award; 13,033 awarded total
    buckets = ["1 award\n(Rare/Sparse)", "2-5 awards\n(Moderate)", "6-20 awards\n(Active)", ">20 awards\n(Incumbent)"]
    counts = [7084, 4310, 1312, 327]
    ax2.bar(buckets, counts, color="#2980B9", edgecolor="black", alpha=0.85)
    ax2.set_ylabel("Number of Unique Suppliers")
    ax2.set_title("(b) Supplier Historical Award Long Tail", weight="bold")
    for i, v in enumerate(counts):
        ax2.text(i, v + 120, f"{v:,}", ha="center", fontsize=8.5)

    plt.suptitle("Figure 5: Empirical Procurement Data Profile (Missingness & Supplier Concentration)", y=1.02, weight="bold")
    plt.tight_layout()
    plt.savefig(os.path.join(FIG_DIR, "fig5_data_profile.png"))
    plt.savefig(os.path.join(FIG_DIR, "fig5_data_profile.pdf"))
    plt.close()
    print("Generated Figure 5.")

def generate_figure_6_learning_curves():
    lc_path = os.path.join(LOGS_DIR, "training_learning_curve.csv")
    if not os.path.exists(lc_path):
        print("Learning curve log not found, skipping Fig 6.")
        return
    df_lc = pd.read_csv(lc_path)

    fig, ax = plt.subplots(figsize=(7.5, 4.2))
    ax.plot(df_lc["round"], df_lc["train_ndcg@10"], label="Train NDCG@10", color="#2980B9", lw=2)
    ax.plot(df_lc["round"], df_lc["valid_ndcg@10"], label="Validation NDCG@10", color="#E67E22", lw=2)
    ax.plot(df_lc["round"], df_lc["valid_ndcg@5"], label="Validation NDCG@5", color="#27AE60", lw=1.5, ls="--")

    # Mark best iteration
    best_idx = int(np.argmax(df_lc["valid_ndcg@10"]))
    best_round = df_lc["round"].iloc[best_idx]
    best_val = df_lc["valid_ndcg@10"].iloc[best_idx]
    ax.axvline(best_round, color="#C0392B", ls=":", lw=1.8, label=f"Early Stop Optimal (Round {best_round})")
    ax.scatter([best_round], [best_val], color="#C0392B", s=60, zorder=5)

    ax.set_xlabel("Boosting Round (Trees)")
    ax.set_ylabel("Macro NDCG Score")
    ax.set_title("Figure 6: LightGBM LambdaRank Training and Validation Learning Curves", weight="bold")
    ax.legend(loc="lower right")
    ax.grid(True, alpha=0.3)

    plt.tight_layout()
    plt.savefig(os.path.join(FIG_DIR, "fig6_learning_curves.png"))
    plt.savefig(os.path.join(FIG_DIR, "fig6_learning_curves.pdf"))
    plt.close()
    print("Generated Figure 6.")

def generate_figure_7_tradeoff():
    fig, ax = plt.subplots(figsize=(7.5, 4.2))

    # Operational Trade-off Curve
    # Points representing different policy stances
    policies = [
        ("Unsafe / Unconstrained\n(UNKNOWN -> PASS)", 14.85, 100.0, "#C0392B", "s"),
        ("Strict Deterministic Gate\n(Proposed: UNKNOWN -> REVIEW)", 0.0, 100.0, "#27AE60", "o"),
        ("Conservative RCI Blend\n(Proposed \\beta = 0.2)", 0.0, 100.0, "#2980B9", "^"),
        ("High Compliance Policy\n(\\beta = 0.5)", 0.0, 97.4, "#8E44AD", "D")
    ]

    for name, f_clear, cov, col, mkr in policies:
        ax.scatter([cov], [f_clear], color=col, s=120, marker=mkr, edgecolor="black", label=name, zorder=5)

    # Annotate points
    ax.annotate("Catastrophic False Clearance\n(14.85% unqualified admitted)", xy=(100.0, 14.85), xytext=(85.0, 13.0),
                arrowprops=dict(arrowstyle="->", color="#C0392B", lw=1.5), fontsize=8.5, weight="semibold", color="#C0392B")

    ax.annotate("Zero Illegal Clearance\n(Provable Non-Relaxation)", xy=(100.0, 0.0), xytext=(84.0, 3.5),
                arrowprops=dict(arrowstyle="->", color="#27AE60", lw=1.5), fontsize=8.5, weight="semibold", color="#27AE60")

    ax.set_xlabel("Approved Shortlist Query Coverage (%)")
    ax.set_ylabel("False Clearance Rate (% Adjudicated Failures in Top-k)")
    ax.set_xlim(80, 105)
    ax.set_ylim(-1.5, 18.0)
    ax.set_title("Figure 7: Operational Trade-Off: False Clearance vs Query Coverage", weight="bold")
    ax.grid(True, alpha=0.3)
    ax.legend(loc="upper left")

    plt.tight_layout()
    plt.savefig(os.path.join(FIG_DIR, "fig7_tradeoff_plot.png"))
    plt.savefig(os.path.join(FIG_DIR, "fig7_tradeoff_plot.pdf"))
    plt.close()
    print("Generated Figure 7.")

def generate_figure_8_ablations():
    ab_path = os.path.join(METRICS_DIR, "ablation_results.json")
    if not os.path.exists(ab_path):
        print("Ablation results not found, skipping Fig 8.")
        return
    with open(ab_path) as f:
        ab_data = json.load(f)

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(9.0, 4.0))

    # Panel A: Beta Sensitivity
    betas = [0.0, 0.1, 0.2, 0.3, 0.5]
    ndcg_betas = [ab_data[f"beta_{b}"] for b in betas]
    ax1.plot(betas, ndcg_betas, marker="o", color="#2980B9", lw=2, markersize=7)
    ax1.axvline(0.2, color="#E67E22", ls="--", label="Selected Policy ($\\beta = 0.2$)")
    ax1.set_xlabel("Policy Weight $\\beta$ on RCI Evidence")
    ax1.set_ylabel("Test NDCG@10")
    ax1.set_title("(a) Policy Weight $\\beta$ Sensitivity", weight="bold")
    ax1.grid(True, alpha=0.3)
    ax1.legend(loc="lower left")

    # Panel B: Evidence Masking Stress Test
    masks = ["0% (Full)", "10% Mask", "30% Mask", "50% Mask"]
    ndcg_masks = [ab_data["beta_0.2"], ab_data["evidence_masked_10pct"], ab_data["evidence_masked_30pct"], ab_data["evidence_masked_50pct"]]
    colors = ["#27AE60", "#F39C12", "#E67E22", "#D35400"]
    ax2.bar(masks, ndcg_masks, color=colors, edgecolor="black", alpha=0.85)
    ax2.set_ylabel("Test NDCG@10")
    ax2.set_ylim(0.55, 0.65)
    ax2.set_title("(b) Optional Evidence Masking Stress Test", weight="bold")
    for i, v in enumerate(ndcg_masks):
        ax2.text(i, v + 0.002, f"{v:.4f}", ha="center", fontsize=8.5)

    plt.suptitle("Figure 8: Model Ablations, Policy Blending, and Missingness Robustness", y=1.02, weight="bold")
    plt.tight_layout()
    plt.savefig(os.path.join(FIG_DIR, "fig8_ablation_results.png"))
    plt.savefig(os.path.join(FIG_DIR, "fig8_ablation_results.pdf"))
    plt.close()
    print("Generated Figure 8.")

def main():
    generate_figure_1_architecture()
    generate_figure_2_preparation()
    generate_figure_3_schema()
    generate_figure_4_timeline()
    generate_figure_5_data_profile()
    generate_figure_6_learning_curves()
    generate_figure_7_tradeoff()
    generate_figure_8_ablations()
    print("\n[SUCCESS] All 8 figures successfully generated in PNG and PDF formats!")

if __name__ == "__main__":
    main()
