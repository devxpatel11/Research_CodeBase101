"""
Spanish BOE (2014-2024) v3 Dataset Audit:
Reproduces Section 3.1 audit from the Research Guide:
  - Validates total rows (97,154) and columns (15).
  - Measures exact duplicate rows (expected 124) and unique source URLs (90,140).
  - Categorizes notice types: Contratación vs Licitación (63,101 vs 34,053).
  - Detects 'No disponible' sentinels in supplier-name, award-value, estimated-value, CPV.
  - Audits suspicious extraction patterns (e.g. tribunal names in supplier fields).
"""

import os
import sys
import pandas as pd
import numpy as np
import json

sys.stdout.reconfigure(encoding='utf-8')

BOE_CSV = os.path.join("data", "raw", "boe", "licitaciones_contrataciones_BOE_2014_2024.csv")

def audit_boe():
    print("=" * 60)
    print("SPANISH BOE v3 DATASET AUDIT")
    print("=" * 60)
    print(f"Loading {BOE_CSV} ({os.path.getsize(BOE_CSV):,} bytes)...")

    # Read as strings to preserve all formatting and sentinels
    df = pd.read_csv(BOE_CSV, dtype=str, keep_default_na=False)
    
    total_rows, total_cols = df.shape
    print(f"Total rows: {total_rows:,}, Total columns: {total_cols}")
    print(f"Columns: {list(df.columns)}")

    # 1. Exact duplicates
    duplicate_mask = df.duplicated()
    exact_duplicates_count = int(duplicate_mask.sum())
    print(f"Exact duplicate rows: {exact_duplicates_count:,} (expected 124)")

    # 2. Unique source URLs
    url_col = None
    for col in df.columns:
        if 'url' in col.lower() or 'enlace' in col.lower() or 'link' in col.lower():
            url_col = col
            break
    if not url_col:
        # Check first column or search by content
        for col in df.columns:
            if df[col].str.startswith('http').any():
                url_col = col
                break
    
    unique_urls = int(df[url_col].nunique()) if url_col else None
    print(f"URL column: '{url_col}' -> Unique source URLs: {unique_urls:,} (expected 90,140)")

    # 3. Notice types: Contratacion vs Licitacion
    # Check type or section column
    type_col = None
    for col in ['tipo', 'tipo_anuncio', 'seccion', 'categoria', 'tipo_aviso']:
        if col in df.columns:
            type_col = col
            break
    if not type_col:
        # Search column containing 'Contratación' or 'Licitación'
        for col in df.columns:
            val_counts = df[col].value_counts()
            if any('contrataci' in str(k).lower() for k in val_counts.index):
                type_col = col
                break

    notice_counts = {}
    if type_col:
        print(f"Notice type column: '{type_col}'")
        raw_types = df[type_col].value_counts()
        for k, v in raw_types.items():
            notice_counts[k] = int(v)
            print(f"  - {k}: {v:,}")

    # 4. Sentinel 'No disponible' detection
    sentinel = "No disponible"
    sentinel_counts = {}
    for col in df.columns:
        cnt = int((df[col].str.strip() == sentinel).sum())
        if cnt > 0:
            sentinel_counts[col] = cnt
            print(f"Sentinel '{sentinel}' in '{col}': {cnt:,}")

    # 5. Suspicious supplier extraction: Tribunal names in supplier column
    # Look for supplier name column
    sup_col = None
    for col in ['adjudicatario', 'contratista', 'empresa', 'proveedor', 'nombre_adjudicatario']:
        if col in df.columns:
            sup_col = col
            break
    if not sup_col:
        for col in df.columns:
            if 'adjudicat' in col.lower():
                sup_col = col
                break

    tribunal_count = 0
    if sup_col:
        print(f"Supplier column: '{sup_col}'")
        tribunal_mask = df[sup_col].str.contains('tribunal', case=False, na=False)
        tribunal_count = int(tribunal_mask.sum())
        print(f"Records with 'tribunal' in supplier column: {tribunal_count:,} (expected ~1,244)")

    report = {
        "dataset": "licitaciones_contrataciones_BOE_2014_2024.csv",
        "file_bytes": os.path.getsize(BOE_CSV),
        "total_rows": total_rows,
        "total_columns": total_cols,
        "column_names": list(df.columns),
        "exact_duplicates": exact_duplicates_count,
        "url_column": url_col,
        "unique_source_urls": unique_urls,
        "type_column": type_col,
        "notice_type_breakdown": notice_counts,
        "sentinel_value": sentinel,
        "sentinel_occurrences_by_column": sentinel_counts,
        "supplier_column": sup_col,
        "tribunal_name_in_supplier_field": tribunal_count,
        "audit_resolution_policy": [
            "Exact duplicates (124) must be deduplicated before external retrieval evaluation.",
            "Literal sentinels ('No disponible') must be encoded as explicit missing indicators rather than valid string categories.",
            "Licitación notices represent tenders prior to award and naturally lack awarded suppliers (hence 34,101 'No disponible' values).",
            "Tribunal extractions (~1,244) reflect institutional dispute/resolution notices where official public bodies were misidentified as commercial contractors by heuristic parsers; must be filtered from commercial supplier benchmarks."
        ]
    }

    out_file = os.path.join("experiments", "metrics", "boe_audit_report.json")
    with open(out_file, "w", encoding="utf-8") as f:
        json.dump(report, f, indent=2, ensure_ascii=False)
    print(f"\n[SUCCESS] Saved comprehensive BOE audit report to {out_file}")

if __name__ == "__main__":
    audit_boe()
