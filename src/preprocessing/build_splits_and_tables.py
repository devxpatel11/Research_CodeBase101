"""
Data Extraction, Chronological Splitting, and Feature Lineage Generation Script:
Parses TenderKG into clean relational tables:
  - requests (lots, tenders, buyers, CPV, text, dates)
  - suppliers (companies, SIREN, staffing, NAF, legal form)
  - outcomes (lot-supplier award edges)
Applies chronological partitioning grouped by parent tender (60% Train / 20% Val / 20% Test).
Exports split manifest, data dictionary, and feature lineage tables.
"""

import os
import sys
import json
import csv
import re
import pandas as pd
import numpy as np
import rdflib
from datetime import datetime

sys.stdout.reconfigure(encoding='utf-8')

TENDERKG_DIR = os.path.join("data", "raw", "tenderkg")
PROCESSED_DIR = os.path.join("data", "processed")
SPLITS_DIR = os.path.join("data", "splits")
DELIVERABLES_DIR = "deliverables"

TKG = rdflib.Namespace("http://intescia_kg.com/")
RDF = rdflib.RDF

def clean_literal(lit):
    if lit is None:
        return ""
    s = str(lit).strip()
    return s.split("^^")[0].strip('"').strip("'")

def extract_id(uri):
    if uri is None:
        return ""
    s = str(uri).strip()
    if "#" in s:
        return s.split("#")[-1]
    return s.split("/")[-1]

def build_relational_tables():
    print("=" * 60)
    print("EXTRACTING RELATIONAL TABLES FROM TENDERKG RDF")
    print("=" * 60)

    # 1. Parse Tenders
    print("Loading Tender.ttl...")
    g_tender = rdflib.Graph()
    g_tender.parse(os.path.join(TENDERKG_DIR, "Tender.ttl"), format="turtle")

    tender_records = {}
    lot_to_tender = {}

    for t_node, _, _ in g_tender.triples((None, RDF.type, TKG.Tender)):
        t_id = extract_id(t_node)
        pub_date = None
        awd_date = None
        buyer_id = None
        t_title = None

        for _, p, o in g_tender.triples((t_node, None, None)):
            p_str = str(p)
            if "HAS_PUBLICATION_DATE" in p_str:
                pub_date = clean_literal(o)
            elif "HAS_AWARDED_DATE" in p_str:
                awd_date = clean_literal(o)
            elif "PUBLISHED_BY" in p_str:
                buyer_id = extract_id(o)
            elif "HAS_TITLE" in p_str:
                t_title = clean_literal(o)
            elif "COMPOSED_OF" in p_str:
                lot_id = extract_id(o)
                lot_to_tender[lot_id] = t_id

        tender_records[t_id] = {
            "tender_id": t_id,
            "publication_date": pub_date,
            "award_date": awd_date,
            "buyer_id": buyer_id,
            "title": t_title
        }
    print(f"Extracted {len(tender_records):,} tenders, {len(lot_to_tender):,} lot-to-tender associations.")

    # 2. Parse Lots
    print("Loading Lot.ttl...")
    g_lot = rdflib.Graph()
    g_lot.parse(os.path.join(TENDERKG_DIR, "Lot.ttl"), format="turtle")

    lot_records = []
    outcomes = []

    for l_node, _, _ in g_lot.triples((None, RDF.type, TKG.LOT)):
        lot_id = extract_id(l_node)
        t_id = lot_to_tender.get(lot_id, "")
        
        cpv_code = ""
        lot_title = ""
        lot_desc = ""
        lot_amount = None

        for _, p, o in g_lot.triples((l_node, None, None)):
            p_str = str(p)
            if "CLASSIFIED_BY" in p_str:
                cpv_code = extract_id(o)
            elif "HAS_TITLE" in p_str:
                lot_title = clean_literal(o)
            elif "HAS_DESCRIPTION" in p_str:
                lot_desc = clean_literal(o)
            elif "HAS_AMOUNT" in p_str:
                try:
                    lot_amount = float(clean_literal(o))
                except Exception:
                    pass
            elif "AWARDED_TO" in p_str:
                sup_id = extract_id(o)
                outcomes.append({
                    "lot_id": lot_id,
                    "tender_id": t_id,
                    "supplier_id": sup_id
                })

        t_info = tender_records.get(t_id, {})
        lot_records.append({
            "lot_id": lot_id,
            "tender_id": t_id,
            "buyer_id": t_info.get("buyer_id", ""),
            "publication_date": t_info.get("publication_date", ""),
            "award_date": t_info.get("award_date", ""),
            "cpv_code": cpv_code,
            "title": lot_title,
            "description": lot_desc,
            "amount": lot_amount
        })
    print(f"Extracted {len(lot_records):,} lot records, {len(outcomes):,} award links.")

    # 3. Parse Companies
    print("Loading Company.ttl...")
    g_comp = rdflib.Graph()
    g_comp.parse(os.path.join(TENDERKG_DIR, "Company.ttl"), format="turtle")

    company_records = []
    nine_digit_regex = re.compile(r"^\d{9}$")

    for c_node, _, _ in g_comp.triples((None, RDF.type, TKG.Company)):
        c_id = extract_id(c_node)
        siren = ""
        label = ""
        staff_range = ""
        legal_form = ""
        naf = ""

        for _, p, o in g_comp.triples((c_node, None, None)):
            p_str = str(p)
            if "HAS_SIREN" in p_str:
                siren = clean_literal(o)
            elif "HAS_LABEL" in p_str:
                label = clean_literal(o)
            elif "HAS_SR" in p_str:
                staff_range = extract_id(o)
            elif "HAS_LF" in p_str:
                legal_form = extract_id(o)
            elif "HAS_NAF" in p_str:
                naf = extract_id(o)

        company_records.append({
            "company_id": c_id,
            "siren": siren,
            "siren_valid_format": bool(nine_digit_regex.match(siren)),
            "label": label,
            "staffing_range": staff_range,
            "legal_form": legal_form,
            "naf_code": naf
        })
    print(f"Extracted {len(company_records):,} company records.")

    # Convert to DataFrames
    df_lots = pd.DataFrame(lot_records)
    df_companies = pd.DataFrame(company_records)
    df_outcomes = pd.DataFrame(outcomes)
    df_tenders = pd.DataFrame(list(tender_records.values()))

    # Save processed Parquet / CSV
    df_lots.to_parquet(os.path.join(PROCESSED_DIR, "requests.parquet"), index=False)
    df_companies.to_parquet(os.path.join(PROCESSED_DIR, "suppliers.parquet"), index=False)
    df_outcomes.to_parquet(os.path.join(PROCESSED_DIR, "outcomes.parquet"), index=False)
    df_tenders.to_parquet(os.path.join(PROCESSED_DIR, "tenders.parquet"), index=False)
    print("Successfully saved processed Parquet tables.")

    return df_lots, df_companies, df_outcomes, df_tenders

def create_chronological_splits(df_lots, df_tenders):
    print("\n" + "=" * 60)
    print("CREATING TIME-RESPECTING PARENT-TENDER CHRONOLOGICAL SPLITS")
    print("=" * 60)

    # Filter out anomalous tenders
    # Parse dates
    df_tenders["pub_dt"] = pd.to_datetime(df_tenders["publication_date"], errors="coerce")
    df_tenders["awd_dt"] = pd.to_datetime(df_tenders["award_date"], errors="coerce")

    # Anomaly conditions:
    # 1. Missing publication date
    # 2. Publication year == 2024 (out of horizon)
    # 3. Award date < Publication date (negative duration anomaly)
    valid_mask = (
        df_tenders["pub_dt"].notnull() &
        (df_tenders["pub_dt"].dt.year < 2024) &
        (df_tenders["awd_dt"].isnull() | (df_tenders["awd_dt"] >= df_tenders["pub_dt"]))
    )

    anomalous_tenders = df_tenders[~valid_mask]
    usable_tenders = df_tenders[valid_mask].copy()

    print(f"Total parent tenders: {len(df_tenders):,}")
    print(f"Anomalous / quarantined tenders: {len(anomalous_tenders):,}")
    print(f"Usable tenders: {len(usable_tenders):,}")

    # Sort strictly by publication date
    usable_tenders = usable_tenders.sort_values("pub_dt").reset_index(drop=True)

    n_usable = len(usable_tenders)
    train_end = int(0.60 * n_usable)
    val_end = int(0.80 * n_usable)

    train_tenders = usable_tenders.iloc[:train_end].copy()
    val_tenders = usable_tenders.iloc[train_end:val_end].copy()
    test_tenders = usable_tenders.iloc[val_end:].copy()

    train_tenders["split"] = "train"
    val_tenders["split"] = "validation"
    test_tenders["split"] = "test"

    split_tenders = pd.concat([train_tenders, val_tenders, test_tenders])

    # Map split to lots
    lot_split_df = df_lots.merge(
        split_tenders[["tender_id", "split", "pub_dt"]],
        on="tender_id",
        how="inner"
    )

    print("\n--- Split Statistics (Parent Tenders & Lots) ---")
    for s in ["train", "validation", "test"]:
        sub_t = split_tenders[split_tenders["split"] == s]
        sub_l = lot_split_df[lot_split_df["split"] == s]
        min_date = sub_t["pub_dt"].min().strftime("%Y-%m-%d")
        max_date = sub_t["pub_dt"].max().strftime("%Y-%m-%d")
        print(f"  {s.upper():<12}: {len(sub_t):,} tenders ({len(sub_t)/n_usable*100:.1f}%), {len(sub_l):,} lots | Range: {min_date} to {max_date}")

    # Save split manifest
    split_manifest_path = os.path.join(DELIVERABLES_DIR, "split_manifest.csv")
    split_tenders[["tender_id", "split", "publication_date", "award_date", "buyer_id"]].to_csv(
        split_manifest_path, index=False
    )
    print(f"Saved immutable split manifest to {split_manifest_path}")

    # Also save lot splits
    lot_split_df.to_parquet(os.path.join(SPLITS_DIR, "lots_with_splits.parquet"), index=False)
    return split_tenders, lot_split_df

def generate_data_dictionary_and_lineage():
    print("\n" + "=" * 60)
    print("GENERATING DATA DICTIONARY AND FEATURE LINEAGE")
    print("=" * 60)

    # Data Dictionary
    dict_rows = [
        {"table": "requests", "field_name": "lot_id", "data_type": "string", "description": "Unique identifier for procurement lot (primary ranking query unit)", "nullable": "False", "sentinels": "None", "source": "Lot.ttl"},
        {"table": "requests", "field_name": "tender_id", "data_type": "string", "description": "Parent tender identifier grouping multi-lot competitions", "nullable": "False", "sentinels": "None", "source": "Tender.ttl"},
        {"table": "requests", "field_name": "buyer_id", "data_type": "string", "description": "Public authority / purchasing entity identifier", "nullable": "True", "sentinels": "Empty string", "source": "Tender.ttl PUBLISHED_BY"},
        {"table": "requests", "field_name": "publication_date", "data_type": "datetime (ISO 8601)", "description": "Tender publication timestamp (pre-decision cutoff time t)", "nullable": "False", "sentinels": "None", "source": "Tender.ttl HAS_PUBLICATION_DATE"},
        {"table": "requests", "field_name": "award_date", "data_type": "datetime (ISO 8601)", "description": "Outcome award timestamp (label availability cutoff)", "nullable": "True", "sentinels": "None", "source": "Tender.ttl HAS_AWARDED_DATE"},
        {"table": "requests", "field_name": "cpv_code", "data_type": "string", "description": "Common Procurement Vocabulary 8-digit product/service code", "nullable": "True", "sentinels": "Empty string (1,597 missing)", "source": "Lot.ttl CLASSIFIED_BY"},
        {"table": "requests", "field_name": "title", "data_type": "string", "description": "Lot title text describing purchased goods/services", "nullable": "True", "sentinels": "Empty string", "source": "Lot.ttl HAS_TITLE"},
        {"table": "requests", "field_name": "description", "data_type": "string", "description": "Detailed lot requirement and scope text", "nullable": "True", "sentinels": "Empty string (710 missing)", "source": "Lot.ttl HAS_DESCRIPTION"},
        {"table": "requests", "field_name": "amount", "data_type": "float", "description": "Estimated contract lot value in EUR", "nullable": "True", "sentinels": "NaN / None", "source": "Lot.ttl HAS_AMOUNT"},
        {"table": "suppliers", "field_name": "company_id", "data_type": "string", "description": "Unique company node identifier in knowledge graph", "nullable": "False", "sentinels": "None", "source": "Company.ttl"},
        {"table": "suppliers", "field_name": "siren", "data_type": "string", "description": "French national business register 9-digit enterprise ID", "nullable": "False", "sentinels": "Non-standard (136 cases)", "source": "Company.ttl HAS_SIREN"},
        {"table": "suppliers", "field_name": "siren_valid_format", "data_type": "boolean", "description": "Format verification flag: True if strictly 9 decimal digits", "nullable": "False", "sentinels": "None", "source": "Derived by regex ^[0-9]{9}$"},
        {"table": "suppliers", "field_name": "staffing_range", "data_type": "string", "description": "Company employee headcount bracket identifier", "nullable": "True", "sentinels": "Empty string (4,383 missing)", "source": "Company.ttl HAS_SR"},
        {"table": "suppliers", "field_name": "legal_form", "data_type": "string", "description": "Categorical legal unit structure code (SA, SAS, SARL, etc.)", "nullable": "True", "sentinels": "Empty string", "source": "Company.ttl HAS_LF"},
        {"table": "suppliers", "field_name": "naf_code", "data_type": "string", "description": "Nomenclature d'Activites Francaise economic activity code", "nullable": "True", "sentinels": "Empty string", "source": "Company.ttl HAS_NAF"},
        {"table": "suppliers", "field_name": "label", "data_type": "string", "description": "Registered trade or legal name of supplier", "nullable": "True", "sentinels": "Empty string", "source": "Company.ttl HAS_LABEL"},
        {"table": "outcomes", "field_name": "lot_id", "data_type": "string", "description": "Awarded procurement lot identifier", "nullable": "False", "sentinels": "None", "source": "Lot.ttl AWARDED_TO"},
        {"table": "outcomes", "field_name": "supplier_id", "data_type": "string", "description": "Awarded supplier company identifier (relevance label y=1)", "nullable": "False", "sentinels": "None", "source": "Lot.ttl AWARDED_TO"}
    ]
    dict_path = os.path.join(DELIVERABLES_DIR, "data_dictionary.csv")
    pd.DataFrame(dict_rows).to_csv(dict_path, index=False)
    print(f"Saved data dictionary to {dict_path}")

    # Feature Lineage Table
    lineage_rows = [
        {"feature_name": "cpv_prefix_match_len", "source_field": "requests.cpv_code, past_awards.cpv_code", "transformation": "Longest common prefix length (0-8) between request CPV and supplier historical awarded CPVs", "availability_condition": "Events strictly before request publication_date t", "missing_treatment": "0 for cold-start or missing CPV", "feature_type": "Integer [0, 8]", "interpretation": "Measures domain/product category compatibility."},
        {"feature_name": "past_awards_category_freq", "source_field": "outcomes.supplier_id, requests.cpv_code", "transformation": "log1p(count of supplier historical contract awards sharing top-2 CPV division before t)", "availability_condition": "Award dates strictly < t", "missing_treatment": "0.0 for zero prior category awards", "feature_type": "Continuous >= 0", "interpretation": "Captures supplier experience in the specific commodity area."},
        {"feature_name": "prior_buyer_interactions", "source_field": "outcomes.supplier_id, requests.buyer_id", "transformation": "Count of prior awarded lots between candidate supplier and buyer before t", "availability_condition": "Award dates strictly < t", "missing_treatment": "0 for new buyer-supplier pairs", "feature_type": "Integer >= 0", "interpretation": "Measures incumbent familiarity and past buyer engagement."},
        {"feature_name": "text_similarity_tfidf", "source_field": "requests.title/description, suppliers.label/past_titles", "transformation": "Cosine similarity of TF-IDF vectors (unigrams + bigrams) between request scope and supplier profile", "availability_condition": "Request text and past historical award text prior to t", "missing_treatment": "0.0 if text absent", "feature_type": "Continuous [0, 1]", "interpretation": "Semantic match between buyer procurement description and supplier domain."},
        {"feature_name": "supplier_total_prior_awards", "source_field": "outcomes.supplier_id", "transformation": "log1p(total awards won by supplier across all categories before t)", "availability_condition": "Award dates strictly < t", "missing_treatment": "0.0 for first-time / cold-start bidders", "feature_type": "Continuous >= 0", "interpretation": "General supplier scale and public procurement bidding maturity."},
        {"feature_name": "staffing_range_code", "source_field": "suppliers.staffing_range", "transformation": "Ordinal rank encoding of employee brackets (0-9)", "availability_condition": "Static supplier attribute available at t", "missing_treatment": "Dedicated sentinel code -1 with missingness flag", "feature_type": "Categorical / Ordinal", "interpretation": "Proxy for delivery capacity and organizational scale."},
        {"feature_name": "siren_valid_format_flag", "source_field": "suppliers.siren", "transformation": "Binary indicator 1{siren matches 9 digits}", "availability_condition": "Available at registration / decision time t", "missing_treatment": "0 if invalid format or absent", "feature_type": "Binary {0, 1}", "interpretation": "Basic registration integrity check before external registry lookup."},
        {"feature_name": "legal_form_encoded", "source_field": "suppliers.legal_form", "transformation": "Frequency encoding on training partition only; unseen values mapped to 0", "availability_condition": "Fitted strictly on train split", "missing_treatment": "0 for missing legal form", "feature_type": "Continuous [0, 1]", "interpretation": "Corporate legal structure category."}
    ]
    lineage_path = os.path.join(DELIVERABLES_DIR, "feature_lineage.csv")
    pd.DataFrame(lineage_rows).to_csv(lineage_path, index=False)
    print(f"Saved feature lineage table to {lineage_path}")

def main():
    df_lots, df_companies, df_outcomes, df_tenders = build_relational_tables()
    create_chronological_splits(df_lots, df_tenders)
    generate_data_dictionary_and_lineage()
    print("\n[SUCCESS] Relational extraction, splitting, and lineage complete!")

if __name__ == "__main__":
    main()
