#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PO Extractor - Regex Only
Reads from:  json/<COMPANY>/<date>/
Saves Excel to: results/<COMPANY>/<date>.xlsx
"""

import json
import re
import sys
from pathlib import Path
from typing import Dict, Optional, List
from datetime import datetime

import pandas as pd

# Fix Windows console encoding
if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except AttributeError:
        import io
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# =========================
# Companies to process
# =========================

COMPANIES = ["GAP", "GPI", "GNP"]
BASE_DIR = Path(__file__).parent

# =========================
# Document type mapping
# =========================

TYPE_MAPPING = {
    # Single-letter / short API codes
    'i':                  'Invoice',
    'c':                  'Credit Note',
    'd':                  'Debit Note',
    'ii':                 'Import Invoice',
    'ei':                 'Export Invoice',
    'ec':                 'Export Credit Note',
    'ed':                 'Export Debit Note',
    # Full-text fallbacks
    'invoice':            'Invoice',
    'credit note':        'Credit Note',
    'debit note':         'Debit Note',
    'import invoice':     'Import Invoice',
    'export invoice':     'Export Invoice',
    'export credit note': 'Export Credit Note',
    'export debit note':  'Export Debit Note',
}

def map_document_type(type_value):
    if not type_value:
        return "Invoice"
    return TYPE_MAPPING.get(str(type_value).lower().strip(), str(type_value))

# =========================
# PO Regex & Cleaning
# =========================

# Matches purchaseOrderReference value
PO_PATTERN = re.compile(r'"?purchaseOrderReference"?\s*:\s*"([^"]{1,50})"')

# Strip common PO prefixes: PO#, PO-, P.O., Purchase Order:, etc.
PO_PREFIX = re.compile(r'^(?:P\.?O\.?|Purchase\s*Order)\s*[#:\-\s]*', re.IGNORECASE)

# Strip company code prefix like "GNP / " or "GAP - "
COMPANY_PREFIX = re.compile(r'^[A-Z]{2,5}\s*[/\-]\s*', re.IGNORECASE)

# Strip trailing garbage like "(File: 400)" or "(file: something)"
TRAILING_JUNK = re.compile(r'\s*\(.*?\)\s*$')

EXCLUSIONS = [
    re.compile(r'^20[2-3][0-9]$'),
    re.compile(r'^\d{1,4}\.\d{2,4}$'),
    re.compile(r'^\d{10,}$'),
    re.compile(r'^[0-9]{1,3}$'),
    re.compile(r'^\d{4}-\d{2}-\d{2}$'),
]

def clean_po_number(po: str) -> str:
    """Remove all known garbage from PO value."""
    po = TRAILING_JUNK.sub('', po).strip()     # "(File: 400)" etc.
    po = COMPANY_PREFIX.sub('', po).strip()    # "GNP / " or "GAP - "
    po = PO_PREFIX.sub('', po).strip()         # "PO#", "P.O.", "Purchase Order:"
    return po

def is_excluded(candidate: str) -> bool:
    return any(p.match(candidate) for p in EXCLUSIONS)

# =========================
# PO Extraction
# =========================

def extract_po(json_data: Dict) -> str:
    """Extract purchaseOrderReference from outer or inner (nested) JSON."""
    # Prefer the inner document string if present
    inner = json_data.get("document")
    if isinstance(inner, str):
        try:
            search_data = json.loads(inner)
        except (json.JSONDecodeError, ValueError):
            search_data = json_data
    else:
        search_data = json_data

    json_str = json.dumps(search_data)

    for match in PO_PATTERN.finditer(json_str, re.IGNORECASE):
        candidate = clean_po_number(match.group(1).strip())
        if candidate and not is_excluded(candidate):
            return candidate

    return ""

# =========================
# Full Invoice Data Extraction
# =========================

def extract_invoice_fields(json_data: Dict) -> Dict:
    """
    Extract all fields matching json_extractor.py output columns:
    INTERNAL ID -1, INTERNAL ID -2, DATE, TYPE, version,
    TOTAL VALUE EGP, FROM, REGESTRAION NUMBER, STATUS, REGESTRAION, PO number
    """
    # Parse inner document if nested
    inner = json_data.get("document")
    if isinstance(inner, str):
        try:
            doc = json.loads(inner)
        except (json.JSONDecodeError, ValueError):
            doc = {}
    else:
        doc = json_data  # flat format fallback

    # --- IDs ---
    uuid       = json_data.get("uuid", "") or doc.get("uuid", "")
    internal_id = json_data.get("internalId", "") or doc.get("internalID", "")

    # --- Date ---
    date_raw = json_data.get("dateTimeReceived", "") or json_data.get("dateTimeIssued", "")
    try:
        date_val = datetime.strptime(date_raw.split("T")[0], "%Y-%m-%d").strftime("%Y-%m-%d") if date_raw else ""
    except ValueError:
        date_val = ""

    # --- Type ---
    doc_type = map_document_type(json_data.get("typeName", "") or doc.get("documentType", ""))

    # --- Version ---
    version = json_data.get("typeVersionName", "") or doc.get("documentTypeVersion", "")

    # --- Total ---
    total = json_data.get("total", "") or doc.get("totalAmount", "")

    # --- Issuer ---
    issuer_name = json_data.get("issuerName", "") or (
        doc.get("issuer", {}).get("name", "") if isinstance(doc.get("issuer"), dict) else ""
    )
    issuer_id = json_data.get("issuerId", "") or (
        doc.get("issuer", {}).get("id", "") if isinstance(doc.get("issuer"), dict) else ""
    )

    # --- Status ---
    status = json_data.get("status", "") or doc.get("status", "")

    # --- Receiver ---
    receiver_id = json_data.get("receiverId", "") or (
        doc.get("receiver", {}).get("id", "") if isinstance(doc.get("receiver"), dict) else ""
    )

    # --- PO ---
    po_number = extract_po(json_data)

    return {
        "INTERNAL ID -1":     uuid,
        "INTERNAL ID -2":     internal_id,
        "DATE":               date_val,
        "TYPE":               doc_type,
        "version":            version,
        "TOTAL VALUE EGP":    total,
        "FROM":               issuer_name,
        "REGESTRAION NUMBER": issuer_id,
        "STATUS":             status,
        "REGESTRAION":        receiver_id,
        "PO number":          po_number,
    }

# =========================
# File Processing
# =========================

def process_json_file(json_file: Path) -> Optional[Dict]:
    try:
        with open(json_file, 'r', encoding='utf-8') as f:
            json_data = json.load(f)
        return extract_invoice_fields(json_data)
    except json.JSONDecodeError:
        print(f"  ⚠️  Invalid JSON: {json_file.name}")
        return None
    except KeyboardInterrupt:
        raise
    except Exception as e:
        print(f"  ⚠️  Error in {json_file.name}: {e}")
        return None

# =========================
# Process one company/date folder
# =========================

def process_folder(json_folder: Path, company: str, date_label: str) -> List[Dict]:
    json_files = sorted(json_folder.glob("*.json"))

    if not json_files:
        print(f"  ⚠️  No JSON files found")
        return []

    print(f"  📊 Files: {len(json_files)}")

    results = []
    po_found = errors = 0

    try:
        for i, jf in enumerate(json_files, 1):
            row = process_json_file(jf)
            if row is None:
                errors += 1
                continue

            row['company']    = company
            row['date_label'] = date_label

            if row.get('PO number'):
                po_found += 1

            results.append(row)

            if i % 10 == 0 or i == len(json_files):
                print(f"  [{i}/{len(json_files)}] POs found: {po_found}  Errors: {errors}")

    except KeyboardInterrupt:
        print(f"\n  ⚠️  Interrupted at {len(results)}/{len(json_files)}")

    return results

# =========================
# Excel Export
# =========================

COLUMN_ORDER = [
    "INTERNAL ID -1", "INTERNAL ID -2", "DATE", "TYPE",
    "version", "TOTAL VALUE EGP", "FROM", "REGESTRAION NUMBER",
    "STATUS", "REGESTRAION", "PO number"
]

def save_excel(results: List[Dict], output_path: Path):
    output_path.parent.mkdir(parents=True, exist_ok=True)

    df = pd.DataFrame(results)

    # Keep only the standard columns (drop internal company/date_label)
    cols = [c for c in COLUMN_ORDER if c in df.columns]
    df = df[cols]

    # Drop full duplicates on UUID
    if "INTERNAL ID -1" in df.columns:
        df = df.drop_duplicates(subset=["INTERNAL ID -1"], keep='first')

    df.to_excel(output_path, index=False, engine='openpyxl')
    print(f"  💾 Saved → {output_path}  ({len(df)} rows)")


# =========================
# General CSV Export
# =========================

CSV_COLUMN_MAP = {
    "INTERNAL ID -2": "INVOICE NUMBER",
    "DATE":           "INVOICE DATE",
    "TYPE":           "INVOICE TYPE",
    "TOTAL VALUE EGP":"AMOUNT",
    "PO number":      "PO NUMBER",
    "INTERNAL ID -1": "REFERENCE1",
}

def save_general_csv(all_results: List[Dict], base_dir: Path):
    """
    Export a single general CSV/Excel file for all companies combined.
    Columns: COMPANY, INVOICE NUMBER, INVOICE DATE, INVOICE TYPE,
             AMOUNT, PO NUMBER, REFERENCE1
    Filename: invoices_YYYY-MM-DD.xlsx  (today's date)
    Saved to: results/invoices_YYYY-MM-DD.xlsx
    """
    if not all_results:
        print("  ⚠️  No data for general CSV export")
        return

    rows = []
    for r in all_results:
        company_raw = r.get("company", "")
        # Ensure company ends with _OU
        company_ou = company_raw if company_raw.endswith("_OU") else f"{company_raw}_OU"

        # Convert date from YYYY-MM-DD → DD-Mon-YY  (e.g. 30-Dec-25)
        raw_date = r.get("DATE", "")
        try:
            formatted_date = datetime.strptime(raw_date, "%Y-%m-%d").strftime("%d-%b-%y")
        except (ValueError, TypeError):
            formatted_date = raw_date

        row = {
            "COMPANY":        company_ou,
            "INVOICE NUMBER": r.get("INTERNAL ID -2", ""),
            "INVOICE DATE":   formatted_date,
            "INVOICE TYPE":   r.get("TYPE", ""),
            "AMOUNT":         r.get("TOTAL VALUE EGP", ""),
            "PO NUMBER":      r.get("PO number", ""),
            "TAX ID":         r.get("REGESTRAION", ""),
            "REFERENCE1":     r.get("INTERNAL ID -1", ""),
        }
        rows.append(row)

    df = pd.DataFrame(rows)

    # Keep only rows that have a PO NUMBER
    df = df[df["PO NUMBER"].notna() & (df["PO NUMBER"].astype(str).str.strip() != "")]

    # Drop duplicates on REFERENCE1 (UUID)
    df = df.drop_duplicates(subset=["REFERENCE1"], keep="first")

    # Use the actual invoice date from the data, not today's date
    invoice_dates = [r.get("DATE", "") for r in all_results if r.get("DATE", "")]
    if invoice_dates:
        # DATE is stored as YYYY-MM-DD internally
        date_str = sorted(invoice_dates, reverse=True)[0]
    else:
        date_str = datetime.now().strftime("%Y-%m-%d")
    out_path  = base_dir / "results" / f"invoices_{date_str}.csv"
    out_path.parent.mkdir(parents=True, exist_ok=True)

    df.to_csv(out_path, index=False, encoding="utf-8-sig")
    print(f"\n  📋 General CSV saved → {out_path}  ({len(df)} rows)")

# =========================
# Summary
# =========================

def print_summary(all_results: List[Dict]):
    if not all_results:
        return

    print("\n" + "=" * 60)
    print("📊 FINAL SUMMARY")
    print("=" * 60)

    by_company: Dict[str, List] = {}
    for r in all_results:
        c = r.get('company', '?')
        by_company.setdefault(c, []).append(r)

    total_po = sum(1 for r in all_results if r.get('PO number'))
    total    = len(all_results)

    for company, rows in by_company.items():
        found = sum(1 for r in rows if r.get('PO number'))
        print(f"  {company}: {found}/{len(rows)} POs found")

    print(f"\n  Total: {total_po}/{total} ({total_po/total*100:.1f}%)" if total else "")
    print("=" * 60)

# =========================
# Main
# =========================

def main():
    json_root = BASE_DIR / "json"

    if not json_root.exists():
        print(f"❌ 'json/' folder not found: {json_root}")
        sys.exit(1)

    print("\n" + "=" * 60)
    print("🚀 PO EXTRACTOR - REGEX ONLY")
    print("=" * 60)

    all_results = []

    for company in COMPANIES:
        company_dir = json_root / company

        if not company_dir.exists():
            print(f"\n[{company}] ⚠️  Not found: {company_dir}")
            continue

        date_folders = sorted(
            [d for d in company_dir.iterdir() if d.is_dir()],
            reverse=True
        )

        if not date_folders:
            print(f"\n[{company}] ⚠️  No date folders")
            continue

        # Only process the latest folder (most recent date)
        latest_dir = date_folders[0]
        date_label = latest_dir.name
        print(f"\n[{company}] 📅 {date_label} (latest)")

        results = process_folder(latest_dir, company, date_label)

        if results:
            # results/<company>/<date>.xlsx
            xlsx_path = BASE_DIR / "results" / company / f"{date_label}.xlsx"
            save_excel(results, xlsx_path)
            all_results.extend(results)

    print_summary(all_results)

    # ── General combined CSV (all companies) ────────────────────────────────
    save_general_csv(all_results, BASE_DIR)

    print("\n✅ Done!")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n⚠️  Interrupted. Exiting...")
        sys.exit(0)