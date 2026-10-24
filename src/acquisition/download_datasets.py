import os
import sys
import json
import time
import shutil
import hashlib
import urllib.request
from datetime import datetime, timezone

HEADERS = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}

TENDERKG_FILES = [
    {"name": "LF.ttl", "url": "https://ndownloader.figshare.com/files/64193494", "expected_md5": "bdcfc133db4e323b6c3bac158fb92c0c"},
    {"name": "NAF.ttl", "url": "https://ndownloader.figshare.com/files/64193497", "expected_md5": "64455dcfb2489a4a9582d1c24783de72"},
    {"name": "CPV.ttl", "url": "https://ndownloader.figshare.com/files/64193500", "expected_md5": "aaee1fe61fe8608146f64cf47ad93f80"},
    {"name": "Amount.ttl", "url": "https://ndownloader.figshare.com/files/64193503", "expected_md5": "6e5088cfccc052ad7b2339387b57e3f4"},
    {"name": "Lot.ttl", "url": "https://ndownloader.figshare.com/files/64193506", "expected_md5": "b911704844a7975985f0c2860a7b51eb"},
    {"name": "Tender.ttl", "url": "https://ndownloader.figshare.com/files/64193509", "expected_md5": "cea362d8acf674813732fe70f91496c4"},
    {"name": "Company.ttl", "url": "https://ndownloader.figshare.com/files/64193512", "expected_md5": "447d77a6ad64915acd63d4415d1cca37"},
    {"name": "GEO.ttl", "url": "https://ndownloader.figshare.com/files/64193518", "expected_md5": "18e7ac79510eeead15880b51e669c3e0"},
    {"name": "ContractType.ttl", "url": "https://ndownloader.figshare.com/files/64193515", "expected_md5": "216a10e5005236f27e7b29955afa8bb2"},
    {"name": "Staff.ttl", "url": "https://ndownloader.figshare.com/files/64193521", "expected_md5": "9ec1b01d6b20e3705089924d98aef9dc"},
    {"name": "DPT.ttl", "url": "https://ndownloader.figshare.com/files/64193524", "expected_md5": "f844013a98fe4edb72df09b60b246b9d"},
]

BOE_FILES = [
    {
        "name": "licitaciones_contrataciones_BOE_2014_2024.csv",
        "url": "https://zenodo.org/api/records/18712463/files/licitaciones_contrataciones_BOE_2014_2024.csv/content",
        "expected_md5": "32629d15d711fdd304e681431a5fe147",
        "expected_bytes": 67396873
    }
]

LATVIAN_FILES = [
    {
        "name": "dev.zip",
        "url": "https://repository.clarin.lv/repository/xmlui/bitstream/handle/20.500.12574/135/dev.zip?isAllowed=y&sequence=3",
        "expected_md5": "4ddc6cbbcec1c452a86e01482dab186d",
        "expected_bytes": 1929183
    },
    {
        "name": "test.zip",
        "url": "https://repository.clarin.lv/repository/xmlui/bitstream/handle/20.500.12574/135/test.zip?isAllowed=y&sequence=4",
        "expected_md5": "0b23ae8bccd2c161d51c2f1f02145876",
        "expected_bytes": 8753304
    }
]

def compute_hashes(path):
    md5 = hashlib.md5()
    sha256 = hashlib.sha256()
    size = os.path.getsize(path)
    with open(path, 'rb') as f:
        for chunk in iter(lambda: f.read(65536), b''):
            md5.update(chunk)
            sha256.update(chunk)
    return size, md5.hexdigest(), sha256.hexdigest()

def download_file(url, target_path, expected_md5=None):
    os.makedirs(os.path.dirname(target_path), exist_ok=True)
    if os.path.exists(target_path) and os.path.getsize(target_path) > 0:
        size, md5_val, _ = compute_hashes(target_path)
        if expected_md5 and md5_val == expected_md5:
            print(f"File {os.path.basename(target_path)} already exists and verified (MD5 {expected_md5}). Skipping download.")
            return

    print(f"Downloading {os.path.basename(target_path)} from {url}...")
    t0 = time.time()
    req = urllib.request.Request(url, headers=HEADERS)
    with urllib.request.urlopen(req, timeout=180) as resp, open(target_path, 'wb') as out_f:
        shutil.copyfileobj(resp, out_f)
    elapsed = time.time() - t0
    size = os.path.getsize(target_path)
    print(f"Finished {os.path.basename(target_path)}: {size:,} bytes in {elapsed:.2f}s ({size/(1024*1024*max(elapsed, 0.01)):.2f} MB/s)")

def main():
    print("=" * 60)
    print("DATASET ACQUISITION SCRIPT (Reproducible & Verifiable)")
    print("=" * 60)
    manifest = {
        "manifest_version": "1.0",
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "datasets": {}
    }

    # 1. TenderKG
    print("\n--- Acquiring TenderKG v1 (Figshare 32149909) ---")
    tenderkg_manifest = {
        "title": "TenderKG: A Knowledge Graph Dataset for Recommendation in Public Procurement",
        "provider": "Figshare / Mokhtari, Smits, Pukhkoy",
        "version": "v1",
        "doi": "https://doi.org/10.6084/m9.figshare.32149909.v1",
        "landing_page": "https://figshare.com/articles/dataset/TenderKG_A_Knowledge_Graph_Dataset_for_Recommendation_in_Public_Procurement/32149909",
        "license": "CC BY 4.0",
        "download_timestamp": datetime.now(timezone.utc).isoformat(),
        "files": []
    }
    for item in TENDERKG_FILES:
        target = os.path.join("data", "raw", "tenderkg", item["name"])
        download_file(item["url"], target, item.get("expected_md5"))
        size, md5_val, sha256_val = compute_hashes(target)
        if item.get("expected_md5"):
            assert md5_val == item["expected_md5"], f"MD5 mismatch for {item['name']}: {md5_val} != {item['expected_md5']}"
        tenderkg_manifest["files"].append({
            "name": item["name"],
            "url": item["url"],
            "bytes": size,
            "md5": md5_val,
            "sha256": sha256_val,
            "verified": True
        })
    manifest["datasets"]["tenderkg"] = tenderkg_manifest

    # 2. BOE v3
    print("\n--- Acquiring BOE v3 (Zenodo 18712463) ---")
    boe_manifest = {
        "title": "Licitaciones y Contrataciones del BOE (2014-2024)",
        "provider": "Zenodo / Manuel Munoz Pla",
        "version": "v3",
        "doi": "https://doi.org/10.5281/zenodo.18712463",
        "landing_page": "https://zenodo.org/records/18712463",
        "license": "CC0 1.0 Universal",
        "download_timestamp": datetime.now(timezone.utc).isoformat(),
        "files": []
    }
    for item in BOE_FILES:
        target = os.path.join("data", "raw", "boe", item["name"])
        download_file(item["url"], target, item.get("expected_md5"))
        size, md5_val, sha256_val = compute_hashes(target)
        if item.get("expected_md5"):
            assert md5_val == item["expected_md5"], f"MD5 mismatch for {item['name']}: {md5_val} != {item['expected_md5']}"
        boe_manifest["files"].append({
            "name": item["name"],
            "url": item["url"],
            "bytes": size,
            "md5": md5_val,
            "sha256": sha256_val,
            "verified": True
        })
    manifest["datasets"]["boe"] = boe_manifest

    # 3. Latvian
    print("\n--- Acquiring Latvian Procurement Validation Dataset ---")
    latvian_manifest = {
        "title": "Procurement Validation Dataset",
        "provider": "CLARIN-LV / University of Latvia (Deksne et al.)",
        "version": "1.0",
        "pid": "https://repository.clarin.lv/repository/xmlui/handle/20.500.12574/135",
        "landing_page": "https://repository.clarin.lv/repository/xmlui/handle/20.500.12574/135",
        "license": "CC BY-NC-SA 4.0",
        "download_timestamp": datetime.now(timezone.utc).isoformat(),
        "files": []
    }
    for item in LATVIAN_FILES:
        target = os.path.join("data", "raw", "latvian", item["name"])
        download_file(item["url"], target, item.get("expected_md5"))
        size, md5_val, sha256_val = compute_hashes(target)
        if item.get("expected_md5"):
            assert md5_val == item["expected_md5"], f"MD5 mismatch for {item['name']}: {md5_val} != {item['expected_md5']}"
        latvian_manifest["files"].append({
            "name": item["name"],
            "url": item["url"],
            "bytes": size,
            "md5": md5_val,
            "sha256": sha256_val,
            "verified": True
        })
    manifest["datasets"]["latvian"] = latvian_manifest

    # 4. INSEE Sirene documentation metadata
    print("\n--- Documenting INSEE Sirene Registry & Metadata ---")
    sirene_manifest = {
        "title": "Base Sirene des entreprises et de leurs etablissements (SIREN / SIRET)",
        "provider": "INSEE (Institut national de la statistique et des etudes economiques)",
        "landing_page": "https://www.data.gouv.fr/datasets/base-sirene-des-entreprises-et-de-leurs-etablissements-siren-siret",
        "license": "Open Licence 2.0 (Licence Ouverte v2.0)",
        "access_method": "Targeted SIREN registry verification and historical format audit",
        "status": "Documented API endpoints and registry schema; tested SIREN format coverage against TenderKG",
        "endpoints": {
            "StockUniteLegale": "https://www.data.gouv.fr/api/1/datasets/r/350182c9-148a-46e0-8389-76c2ec1374a3",
            "StockEtablissement": "https://www.data.gouv.fr/api/1/datasets/r/a29c1297-1f92-4e2a-8f6b-8c902ce96c5f",
            "StockUniteLegaleHistorique": "https://www.data.gouv.fr/api/1/datasets/r/1b9290ed-d0bc-461f-ba31-0250a99cc140"
        }
    }
    manifest["datasets"]["sirene"] = sirene_manifest

    with open(os.path.join("deliverables", "dataset_manifest.json"), "w", encoding="utf-8") as f:
        json.dump(manifest, f, indent=2)
    print("\nSuccessfully generated deliverables/dataset_manifest.json!")

if __name__ == "__main__":
    main()
