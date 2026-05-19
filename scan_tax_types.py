#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Scan all JSON files and report every taxType found.
Run from the same folder as your json/ directory.
"""

import json
from pathlib import Path
from collections import defaultdict

BASE_DIR = Path(__file__).parent
json_root = BASE_DIR / "json"

tax_summary = defaultdict(list)  # taxType -> list of filenames

files_scanned = 0
files_with_t4  = []

for json_file in json_root.rglob("*.json"):
    try:
        with open(json_file, 'r', encoding='utf-8') as f:
            raw = json.load(f)

        # Parse inner document if nested
        inner_str = raw.get("document")
        if isinstance(inner_str, str):
            try:
                doc = json.loads(inner_str)
            except Exception:
                doc = {}
        else:
            doc = raw

        files_scanned += 1

        for line in doc.get("invoiceLines", []):
            for tax in line.get("taxableItems", []):
                t = tax.get("taxType", "?")
                tax_summary[t].append(json_file.name)
                if t == "T4":
                    files_with_t4.append(json_file.name)

        for t in doc.get("taxTotals", []):
            tt = t.get("taxType", "?")
            if tt not in [x for x in tax_summary]:
                tax_summary[tt].append(json_file.name)

    except Exception as e:
        print(f"  ⚠️  Error in {json_file.name}: {e}")

print(f"\nFiles scanned: {files_scanned}")
print("\n=== Tax Types Found ===")
for tax_type, files in sorted(tax_summary.items()):
    print(f"  {tax_type}: {len(files)} occurrences")

if files_with_t4:
    print(f"\n✅ T4 (Withholding Tax) found in {len(files_with_t4)} files:")
    for f in files_with_t4[:20]:  # show first 20
        print(f"    {f}")
else:
    print("\n❌ T4 (Withholding Tax) not found in any file")
