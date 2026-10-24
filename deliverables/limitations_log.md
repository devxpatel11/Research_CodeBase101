# Limitations, Risk Log, and Unresolved Evidence Boundaries

This document catalogs the empirical limitations, selection biases, and legal boundaries identified during the implementation and evaluation of the **Evidence-Aware Compliance-Constrained Learning-to-Rank System**.

---

## 1. Observational and Dataset Boundaries

### 1.1 Incomplete Bidder Visibility and Selection Bias in TenderKG
- **Limitation:** TenderKG v1 provides records of successfully awarded contracts (`AWARDED_TO`), but does *not* record unsuccessful tender submissions (the full candidate bidder set).
- **Consequence:** Background candidates sampled from the discoverable supplier pool are unobserved in the ground truth. An unobserved candidate cannot be assumed to be unsuitable or disqualified; they may simply have chosen not to bid, or placed second by a minimal scoring margin.
- **Scientific Framing:** The ranking model is rigorously framed and evaluated as an **observed-award retrieval system**, *not* as an oracle predictor of global commercial competence.

### 1.2 Retrospective Knowledge Graph Leakage
- **Limitation:** In the raw TenderKG export, 25,115 global `WORKED_WITH` edges and co-bidding structures were computed across the entire multi-year corpus (2021–2024).
- **Treatment:** These global edges were strictly quarantined and excluded from pre-decision feature engineering. Historical graph features were dynamically reconstructed using only events with `award_datetime < t`.
- **Remaining Constraint:** For queries occurring early in the temporal split (early 2021), historical graph density is naturally sparse, reflecting authentic operational cold-start conditions.

### 1.3 Literal Missing Values and Sentinel Markers in Spanish BOE v3
- **Limitation:** The Spanish BOE v3 dataset contains pervasive literal placeholder strings `"No disponible"` in 34,101 supplier name fields, 34,101 award value fields, 53,112 estimated contract values, and 8,669 CPV classifications.
- **Treatment:** These sentinels were recognized as structural indicators of pre-award notices (`Licitación`, 34,053 records) that naturally lack awarded contractors, and were mapped to explicit missingness masks rather than valid categorical strings.
- **Remaining Constraint:** Direct transfer of French TenderKG models to Spanish BOE is prevented by jurisdictional differences in procurement notice publication semantics.

### 1.4 Non-Flattening of Latvian Procurement Validation Labels
- **Limitation:** The Latvian dataset evaluates 30 projects with 150 main question fields and 55 conditional checklist items. The negative answer class (`"nē"`) constitutes only 1.97% of development and 1.89% of evaluation answers.
- **Treatment:** Main answers (`jā`, `nē`, `n/a`, `?`) were preserved without flattening into universal supplier eligibility labels. A `"nē"` response to a checklist question indicates that a specific procurement tender requirement was omitted or answered negatively in tender documentation, *not* that a named supplier is legally disqualified.

---

## 2. Supplier Identity and Registry Limitations

### 2.1 Format Validity vs. Substantive Legal Good Standing
- **Limitation:** The SIREN format audit on Company.ttl confirmed that 15,204 out of 15,340 identifiers (99.11%) match exactly 9 decimal digits, while 136 records possess non-standard alphanumeric strings.
- **Boundary:** Formatting correctness confirms only that an identifier adheres to national syntax rules. It does *not* prove that:
  - The entity currently exists in the live national register.
  - The entity is in active commercial operation.
  - The candidate possesses required sector-specific licenses or product certifications.
- **Policy Invariant:** Registry existence is treated as a necessary foundational prerequisite, never as sufficient proof of comprehensive regulatory compliance.

### 2.2 Historical Registry Drift (2021 vs. 2026)
- **Limitation:** Public enterprise registries (INSEE Sirene) publish rolling monthly stock snapshots. Retrieving current 2026 registry attributes to evaluate a contract awarded in 2021 represents retrospective lookahead contamination.
- **Treatment:** Only historically timestamped corporate attributes (registration date, historic staffing ranges) available prior to $t$ were utilized.

---

## 3. Methodological and Algorithmic Constraints

### 3.1 Non-Compensatory Mandatory Gate vs. Compensatory Utility
- **Limitation:** A fundamental tenet of public procurement law is that high technical score or low price cannot legally compensate for a failure to satisfy mandatory eligibility requirements.
- **Design Assurance:** The three-outcome gate strictly acts as a hard filter: candidates with verified FAIL are quarantined (REJECT), while candidates with missing or unresolved evidence are routed to human review (REVIEW). No compensatory score blending occurs prior to gate clearance.
- **Operational Trade-Off:** As demonstrated by the empirical evaluation, enforcing strict compliance reduces observed-award NDCG@10 from 0.7897 to 0.6441, because several historical winners in open data lack verified staffing or registry documentation. Relaxing this gate to artificially boost NDCG causes a catastrophic 14.85% false clearance rate.

### 3.2 Epistemic Evidence Intervals vs. Calibrated Probabilities
- **Limitation:** Bounded RCI intervals $[RCI^-, RCI^+]$ and completeness score $Q$ represent deterministic evidence bounds reflecting information completeness, *not* Bayesian posterior probabilities or 95% confidence intervals.
- **Usage Warning:** Procurement officers and downstream algorithms must not interpret $RCI^-$ as a calibrated probability of contract delivery.

### 3.3 Generalization Beyond French Public Procurement
- **Limitation:** Empirical ranking experiments were conducted on French public procurement records (TenderKG). Legal rules, CPV categorization practices, and company identifier schemes vary across jurisdictions.
- **Boundary:** Applying this system to other procurement environments (e.g., US Federal FAR, Indian GeM, or private corporate supply chains) requires acquiring local registry connections, translating statutory qualification rules, and validating against local jurisdiction-specific adjudication datasets.
