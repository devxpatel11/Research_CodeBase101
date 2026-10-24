"""
High-Performance Chronological Ranking Dataset Builder:
Optimized vectorised feature extraction:
  - div_to_suppliers: defaultdict(set) for O(1) set operations.
  - Vectorised TF-IDF candidate score computation via sparse dot product.
  - Strict chronological information isolation (past awards only).
  - Explicit candidate retrieval recall and query group export.
"""

import os
import sys
import math
import random
import time
import json
import numpy as np
import pandas as pd
from datetime import datetime
from collections import defaultdict, Counter
from sklearn.feature_extraction.text import TfidfVectorizer

sys.stdout.reconfigure(encoding='utf-8')

PROCESSED_DIR = os.path.join("data", "processed")
SPLITS_DIR = os.path.join("data", "splits")

def extract_cpv_division(cpv):
    if not cpv or pd.isna(cpv):
        return "00"
    s = str(cpv).strip()
    return s[:2] if len(s) >= 2 else "00"

def get_common_prefix_length(s1, s2):
    if not s1 or not s2:
        return 0
    l = min(len(s1), len(s2))
    for i in range(l):
        if s1[i] != s2[i]:
            return i
    return l

def build_ranking_dataset(
    max_train_queries=1000,
    max_val_queries=300,
    max_test_queries=300,
    candidates_per_query=25,
    random_seed=42
):
    random.seed(random_seed)
    np.random.seed(random_seed)
    t0 = time.time()

    print("=" * 60)
    print("HIGH-PERFORMANCE RANKING DATASET BUILDER")
    print("=" * 60)

    # 1. Load data
    lots_df = pd.read_parquet(os.path.join(SPLITS_DIR, "lots_with_splits.parquet"))
    suppliers_df = pd.read_parquet(os.path.join(PROCESSED_DIR, "suppliers.parquet"))
    outcomes_df = pd.read_parquet(os.path.join(PROCESSED_DIR, "outcomes.parquet"))

    print(f"Loaded {len(lots_df):,} split lots, {len(suppliers_df):,} suppliers, {len(outcomes_df):,} outcome links.")

    # Map lot_id to outcomes
    lot_winners = defaultdict(set)
    for _, row in outcomes_df.iterrows():
        lot_winners[row["lot_id"]].add(row["supplier_id"])

    lots_df["awarded_suppliers"] = lots_df["lot_id"].map(lambda lid: list(lot_winners.get(lid, [])))
    lots_df["winner_count"] = lots_df["awarded_suppliers"].map(len)
    lots_df["pub_dt"] = pd.to_datetime(lots_df["pub_dt"])

    # Filter to lots with at least 1 observed winner
    ltr_lots = lots_df[lots_df["winner_count"] > 0].copy()
    print(f"Lots with >= 1 observed winner: {len(ltr_lots):,} out of {len(lots_df):,}")

    # Supplier metadata dictionary
    sup_lookup = {}
    for _, row in suppliers_df.iterrows():
        sup_lookup[row["company_id"]] = {
            "company_id": row["company_id"],
            "siren_valid_format": 1 if row["siren_valid_format"] else 0,
            "has_staffing": 0 if (pd.isna(row["staffing_range"]) or row["staffing_range"] == "") else 1,
            "label": str(row["label"]) if not pd.isna(row["label"]) else ""
        }
    all_supplier_ids = list(sup_lookup.keys())

    # Build TF-IDF on training queries + suppliers
    print("Fitting TF-IDF on training text...")
    train_corpus = []
    for _, row in ltr_lots[ltr_lots["split"] == "train"].iterrows():
        txt = (str(row["title"]) + " " + str(row["description"])).strip()
        if txt:
            train_corpus.append(txt)
    for sid in all_supplier_ids[:4000]:
        lbl = sup_lookup[sid]["label"]
        if lbl:
            train_corpus.append(lbl)

    tfidf = TfidfVectorizer(max_features=2500, stop_words="english", ngram_range=(1, 2))
    tfidf.fit(train_corpus[:6000])

    sup_labels = [sup_lookup[sid]["label"] for sid in all_supplier_ids]
    sup_tfidf_matrix = tfidf.transform(sup_labels)
    sup_id_to_idx = {sid: idx for idx, sid in enumerate(all_supplier_ids)}

    # Subsample queries preserving parent tenders
    selected_lot_ids_by_split = {}
    all_selected_lot_ids = set()

    sorted_lots = ltr_lots.sort_values("pub_dt").reset_index(drop=True)

    for split_name, max_q in [("train", max_train_queries), ("validation", max_val_queries), ("test", max_test_queries)]:
        sub = sorted_lots[sorted_lots["split"] == split_name]
        tenders = list(sub["tender_id"].unique())
        if len(tenders) > max_q:
            rng = random.Random(random_seed)
            selected_tenders = set(rng.sample(tenders, max_q))
            selected_sub = sub[sub["tender_id"].isin(selected_tenders)]
        else:
            selected_sub = sub
        s_ids = set(selected_sub["lot_id"].values)
        selected_lot_ids_by_split[split_name] = s_ids
        all_selected_lot_ids.update(s_ids)
        print(f"Selected {len(s_ids):,} queries ({selected_sub['tender_id'].nunique():,} tenders) for {split_name.upper()}.")

    # Inverted indices for O(1) temporal features
    past_sup_award_count = defaultdict(int)
    past_sup_cpv_div_count = defaultdict(int)
    past_buyer_sup_count = defaultdict(int)
    div_to_suppliers = defaultdict(set)
    past_sup_cpvs = defaultdict(set)
    past_sup_first_seen = {}

    dataset_records = {"train": [], "validation": [], "test": []}
    query_groups = {"train": [], "validation": [], "test": []}
    retrieval_stats = {"train": [], "validation": [], "test": []}

    print("\nProcessing chronological timeline with vectorized sparse matching...")
    lot_records = sorted_lots.to_dict("records")

    for q_idx, q_row in enumerate(lot_records):
        q_id = q_row["lot_id"]
        q_split = q_row["split"]
        q_dt = q_row["pub_dt"]
        q_cpv = str(q_row["cpv_code"]) if q_row["cpv_code"] else ""
        q_div = extract_cpv_division(q_cpv)
        q_buyer = str(q_row["buyer_id"]) if q_row["buyer_id"] else ""
        q_winners = set(q_row["awarded_suppliers"])

        if q_id in all_selected_lot_ids:
            # Candidate generation:
            cat_cands = list(div_to_suppliers.get(q_div, set()))
            n_cat = min(12, len(cat_cands))
            sampled_cat = random.sample(cat_cands, n_cat) if len(cat_cands) > n_cat else cat_cands

            n_bg = max(10, candidates_per_query - len(q_winners))
            bg_cands = random.sample(all_supplier_ids, min(n_bg, len(all_supplier_ids)))

            # Candidate retrieval recall on operational candidate set
            operational_set = set(sampled_cat + bg_cands)
            retrieved_count = len(q_winners.intersection(operational_set))
            rec_recall = retrieved_count / len(q_winners) if len(q_winners) > 0 else 1.0
            retrieval_stats[q_split].append(rec_recall)

            # Assemble training/ranking candidate pool
            pool_set = set(list(q_winners) + list(operational_set)[:candidates_per_query])
            candidate_pool = list(pool_set)
            random.shuffle(candidate_pool)

            # Vectorized text similarity via single sparse matrix multiply
            q_text = (str(q_row["title"]) + " " + str(q_row["description"])).strip()
            if q_text:
                q_vec = tfidf.transform([q_text])
                cand_row_indices = [sup_id_to_idx[sid] for sid in candidate_pool]
                cand_matrix = sup_tfidf_matrix[cand_row_indices]
                text_sims = (cand_matrix.dot(q_vec.T)).toarray().flatten()
            else:
                text_sims = np.zeros(len(candidate_pool))

            q_records = []
            for i, sup_id in enumerate(candidate_pool):
                sup_data = sup_lookup[sup_id]
                is_pos = 1 if sup_id in q_winners else 0

                tot_awards = past_sup_award_count[sup_id]
                cat_awards = past_sup_cpv_div_count[(sup_id, q_div)]
                buyer_awards = past_buyer_sup_count[(q_buyer, sup_id)]

                # Prefix match
                sup_cpv_set = past_sup_cpvs[sup_id]
                max_pfx = 0
                if sup_cpv_set:
                    for p_cpv in sup_cpv_set:
                        pfx = get_common_prefix_length(q_cpv, p_cpv)
                        if pfx > max_pfx:
                            max_pfx = pfx
                            if max_pfx >= 8:
                                break

                sim = float(text_sims[i])
                days_active = max(0, (q_dt - past_sup_first_seen[sup_id]).days) if sup_id in past_sup_first_seen else 0

                q_records.append({
                    "query_id": q_id,
                    "tender_id": q_row["tender_id"],
                    "supplier_id": sup_id,
                    "target_relevance": is_pos,
                    "feat_cpv_prefix_len": max_pfx,
                    "feat_category_awards_log1p": math.log1p(cat_awards),
                    "feat_total_prior_awards_log1p": math.log1p(tot_awards),
                    "feat_prior_buyer_interactions": buyer_awards,
                    "feat_text_similarity": round(sim, 4),
                    "feat_days_active_log1p": math.log1p(days_active),
                    "feat_staffing_flag": sup_data["has_staffing"],
                    "feat_siren_valid_format": sup_data["siren_valid_format"],
                    "feat_is_cold_start": 1 if tot_awards == 0 else 0
                })

            dataset_records[q_split].extend(q_records)
            query_groups[q_split].append(len(q_records))

        # UPDATE PAST LEDGER (strictly after t)
        for w_id in q_winners:
            past_sup_award_count[w_id] += 1
            past_sup_cpv_div_count[(w_id, q_div)] += 1
            past_buyer_sup_count[(q_buyer, w_id)] += 1
            if q_cpv:
                past_sup_cpvs[w_id].add(q_cpv)
            if w_id not in past_sup_first_seen:
                past_sup_first_seen[w_id] = q_dt
            div_to_suppliers[q_div].add(w_id)

    elapsed = time.time() - t0
    print(f"\nProcessing complete in {elapsed:.2f} seconds!")

    # Save outputs
    print("\n--- Summary of Prepared Datasets ---")
    summary = {}
    for s in ["train", "validation", "test"]:
        df_s = pd.DataFrame(dataset_records[s])
        groups_s = query_groups[s]
        mean_retrieval = np.mean(retrieval_stats[s]) if retrieval_stats[s] else 0.0

        print(f"Partition {s.upper()}:")
        print(f"  Total pairs: {len(df_s):,}, Queries: {len(groups_s):,}")
        print(f"  Positive awards: {df_s['target_relevance'].sum():,} ({df_s['target_relevance'].mean()*100:.2f}%)")
        print(f"  Mean Candidate Retrieval Recall: {mean_retrieval*100:.2f}%")

        out_pq = os.path.join(PROCESSED_DIR, f"dataset_{s}.parquet")
        df_s.to_parquet(out_pq, index=False)
        np.save(os.path.join(PROCESSED_DIR, f"groups_{s}.npy"), np.array(groups_s))

        summary[s] = {
            "total_pairs": len(df_s),
            "total_queries": len(groups_s),
            "positive_awards": int(df_s['target_relevance'].sum()),
            "positive_prevalence_pct": round(float(df_s['target_relevance'].mean() * 100.0), 2),
            "mean_candidate_retrieval_recall_pct": round(float(mean_retrieval * 100.0), 2)
        }

    with open(os.path.join("experiments", "metrics", "candidate_retrieval_report.json"), "w", encoding="utf-8") as f:
        json.dump(summary, f, indent=2)

    print("\n[SUCCESS] Datasets, query groups, and retrieval report saved successfully!")

if __name__ == "__main__":
    build_ranking_dataset(
        max_train_queries=1000,
        max_val_queries=300,
        max_test_queries=300,
        candidates_per_query=25
    )
