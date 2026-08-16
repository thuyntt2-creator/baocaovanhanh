import sys
import io
import os
import re
import json

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

SCRATCH_DIR = r"C:\Users\lap4all\.gemini\antigravity-ide\scratch"

def main():
    mappings = {} # PO name -> (AM, Province)
    
    # 1. Parse docx_content.txt
    docx_path = os.path.join(SCRATCH_DIR, "docx_content.txt")
    if os.path.exists(docx_path):
        print(f"Reading {docx_path}...")
        with open(docx_path, "r", encoding="utf-8") as f:
            for line in f:
                # Look for lists representing rows, e.g. ['Lâm Đồng', 'Trầm Hữu Tiến', 'Bưu Cục 468 Quốc Lộ 20...', ...]
                match = re.search(r"Row \d+:\s*(\[.*\])", line)
                if match:
                    try:
                        row = eval(match.group(1))
                        # Check if row looks like: [Province, AM, PO, ...] or similar
                        if len(row) >= 3:
                            # Let's see if we can identify elements
                            # Usually, one of the elements is the PO name (starts with Bưu Cục / BC)
                            # One is AM, one is Province
                            # Let's inspect the elements
                            po = None
                            am = None
                            prov = None
                            for el in row:
                                el_str = str(el).strip()
                                if el_str.startswith("Bưu Cục") or el_str.startswith("BC ") or el_str.startswith("Bưu cục"):
                                    po = el_str
                            
                            # Let's find AM and Province
                            # Province is one of: Đắk Nông, Lâm Đồng, Ninh Thuận, Khánh Hòa, Bình Thuận, Khánh Hoà
                            provinces = ["Đắk Nông", "Lâm Đồng", "Ninh Thuận", "Khánh Hòa", "Bình Thuận", "Khánh Hoà"]
                            for el in row:
                                el_str = str(el).strip()
                                if el_str in provinces:
                                    prov = el_str
                                    
                            # AM is not Province, not PO, and is a person name
                            # We can check names in the row
                            for el in row:
                                el_str = str(el).strip()
                                if el_str != po and el_str != prov and len(el_str) > 5 and not el_str.startswith("http") and not el_str.replace(".","").replace(",","").replace("-","").isdigit():
                                    # Basic check for name (e.g. Capitalized words)
                                    if all(w[0].isupper() for w in el_str.split() if w and w[0].isalpha()):
                                        am = el_str
                                        
                            if po and am and prov:
                                mappings[po] = (am, prov)
                    except Exception as e:
                        pass
                        
    # 2. Parse comparison_results.txt
    comp_path = os.path.join(SCRATCH_DIR, "comparison_results.txt")
    if os.path.exists(comp_path):
        print(f"Reading {comp_path}...")
        with open(comp_path, "r", encoding="utf-8") as f:
            for line in f:
                parts = [p.strip() for p in line.split("|")]
                if len(parts) >= 2:
                    am = parts[0]
                    po = parts[1]
                    if po.startswith("Bưu Cục") or po.startswith("Bưu cục") or po.startswith("BC "):
                        # Find province from name
                        prov = None
                        for p in ["Đắk Nông", "Lâm Đồng", "Ninh Thuận", "Khánh Hòa", "Bình Thuận", "Khánh Hoà"]:
                            if p in po:
                                prov = p
                                break
                        if am and po and prov:
                            mappings[po] = (am, prov)
                            
    # Let's print out the size and details
    print(f"Found {len(mappings)} unique mappings in files:")
    for po, (am, prov) in sorted(mappings.items()):
        print(f"  '{po}': ('{am}', '{prov}')")
        
    # Write to a JSON file for reuse
    out_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "extracted_mappings.json")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(mappings, f, ensure_ascii=False, indent=2)
    print(f"Saved mappings to {out_path}")

if __name__ == "__main__":
    main()
