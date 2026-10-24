# Independent Compliance Adjudication Rubric and Sampling Protocol

## 1. Purpose and Scope
This document specifies the standardized adjudication protocol and rubric for evaluating the **3-outcome Mandatory Eligibility Gate** and **Optional Regulatory Compatibility Index (RCI)**. The ground truth established through this rubric serves as an independent benchmark to validate system predictions against human expert legal determinations.

Under no circumstances may automated system outputs or rule engine heuristics be substituted for independent human adjudication.

---

## 2. Sampling Design and Strata
The validation pool comprises **100 procurement requests** drawn from the locked chronological test partition (February 2023 – November 2023), paired with **10 candidate suppliers** each, yielding **1,000 linked request–supplier–rule decision pairs**.

### Stratification Dimensions:
1. **Industry / Product Category (CPV Division)**:
   - Manufacturing & Industrial Components (CPV 42xxx, 44xxx): 25 requests.
   - Information Technology & Software (CPV 48xxx, 72xxx): 25 requests.
   - Works & Construction (CPV 45xxx): 25 requests.
   - Specialized Services & Medical Supplies (CPV 33xxx, 79xxx): 25 requests.
2. **Supplier Historical Maturity**:
   - Zero-history / pure cold-start bidders: 30% of candidate pairs.
   - Low-to-moderate historical activity (1–5 prior awards): 40% of candidate pairs.
   - High-activity incumbents (>5 prior awards): 30% of candidate pairs.
3. **Evidence Completeness Regime**:
   - Natural prevalence sample (unmodified documentation state).
   - Enriched challenge sample: Deliberately injected edge cases including expired certificates, missing scopes, non-standard identifiers, and partial name collisions.

---

## 3. Mandatory Predicate Rubrics

Each applicable mandatory predicate $r_j(q, s, t)$ must be adjudicated into strictly one of three outcomes: **PASS**, **FAIL**, or **UNKNOWN**.

### Rule M1: Commercial Legal Existence and Registry Validity
*Predicate definition:* The candidate supplier must correspond to a verified, actively registered legal unit in the official national or cross-border commercial registry (e.g., INSEE Sirene) at decision time $t$.
- **PASS**:
  - The supplier identifier (SIREN/SIRET) matches an active enterprise in the official registry.
  - The registration effective date is strictly prior to request publication time $t$.
  - The enterprise status is active (non-cessée) as of $t$.
- **FAIL**:
  - The identifier format is invalid and cannot be resolved to any legal unit.
  - The enterprise is officially recorded as dissolved, liquidated, or struck off prior to $t$.
  - The enterprise was created strictly *after* $t$ (post-decision incorporation).
- **UNKNOWN**:
  - The registry records are temporarily unreachable or ambiguous due to unresolved mergers/successions.
  - Identification cannot be definitively confirmed without manual registry inquiry.

### Rule M2: Non-Debarment and Sanctions Exclusions
*Predicate definition:* The candidate supplier and its identified controlling beneficial owners must not be subject to active debarment, financial sanctions, or public procurement exclusions within the applicable jurisdiction at time $t$.
- **PASS**:
  - Exact match search against official dated sanctions lists (e.g., OFAC SDN, EU Consolidated Financial Sanctions List) returns no applicable hits as of time $t$.
  - Certified self-declaration of clean criminal record / non-bankruptcy is verified.
- **FAIL**:
  - Positive, verified match against an applicable debarment list with an active penalty horizon covering time $t$.
  - Formal bankruptcy or liquidation proceedings verified in official legal gazettes.
- **UNKNOWN**:
  - Partial or fuzzy name match with an alert that lacks confirming identifier metadata (requires secondary reviewer clearance).
  - Ambiguous beneficial ownership chain preventing definitive clearance.

### Rule M3: Technical and Workforce Capacity Floor
*Predicate definition:* The candidate supplier must demonstrate minimum staffing and operational capacity commensurate with the contract scope.
- **PASS**:
  - Officially reported staffing headcount bracket (INSEE tranche d'effectif) is non-zero and meets the lot minimum threshold.
  - Documented physical establishment within the required delivery territory.
- **FAIL**:
  - Formally documented zero workforce (0 salariés) for a contract requiring substantial on-site execution, with no valid subcontracting or consortium declaration.
- **UNKNOWN**:
  - Missing staffing metadata in commercial registers (structural absence).
  - Unverified capacity claims pending audited payroll documentation.

### Rule M4: Tax and Social Security Compliance
*Predicate definition:* The supplier must possess valid certificates of good standing for social security contributions (URSSAF) and fiscal obligations.
- **PASS**:
  - Formal certificate of tax compliance issued within the 6 months preceding $t$ and unexpired at $t$.
- **FAIL**:
  - Verified administrative notice of tax evasion or unpaid final judicial tax assessments.
- **UNKNOWN**:
  - Certificate missing from supplier profile; administrative validity cannot be confirmed from public open data alone.

---

## 4. Optional Regulatory Compatibility Index (RCI) Rubric
For eligible suppliers entering the approved shortlist, optional evidence is adjudicated as deterministic bounding intervals $[l_j, u_j] \subseteq [0, 1]$:
1. **Verified Specific Certification (e.g., ISO 9001, ISO 14001, ISO 27001)**:
   - Valid certificate covering requested facility and scope: $l_j = 1.0, u_j = 1.0$ ($z_j = 1$).
   - Documented partial coverage (e.g., corporate group certified but target subsidiary unlisted): $l_j = 0.5, u_j = 0.8$ ($z_j = 1$).
   - Completely missing documentation: $l_j = 0.0, u_j = 1.0$ ($z_j = 0$).
2. **Environmental & ESG Compliance Declarations**:
   - Audited third-party verification: $l_j = 1.0, u_j = 1.0$.
   - Unverified self-declaration: $l_j = 0.3, u_j = 0.7$.
   - Absent declaration: $l_j = 0.0, u_j = 1.0$.

---

## 5. Dual-Reviewer Adjudication Protocol
1. **Blinded First-Pass Review**:
   - Reviewer 1 and Reviewer 2 independently inspect request requirements, candidate corporate dossier, registry extracts, and sanctions databases as of decision date $t$.
   - All automated model predictions, ranking scores, and gate determinations are strictly masked.
2. **Disagreement Resolution**:
   - For any item where Reviewer 1 and Reviewer 2 assign differing outcomes (e.g., PASS vs UNKNOWN), an expert third adjudicator (Senior Procurement Legal Specialist) conducts independent review and issues a final binding adjudication with documented rationale.
3. **Inter-Rater Reliability**:
   - Raw percentage agreement $P_o$ and Cohen's Kappa $\kappa = \frac{P_o - P_e}{1 - P_e}$ are calculated separately for each predicate.
   - Adjudications with $\kappa < 0.75$ trigger rubric refinement and re-annotation of the ambiguous class.
