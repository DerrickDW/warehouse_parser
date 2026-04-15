import json
import csv
from pathlib import Path

from parser import normalize_part_for_validation
print("RUNNING CONVERTER")

def jsonl_to_csv(input_path, out_path):

    print("json_to_csv REACHED")
    seen={}

    with input_path.open('r', encoding='utf8') as f:
        print(F"[INFO] Reading {input_path}")
        for line in f:
            row = json.loads(line)

            part = normalize_part_for_validation(row.get("Part Number"))
            desc = str(row.get("Part Description")or "").strip()

            if not part:
                continue

            #keep first good desc, ignore dupes
            if part not in seen and desc:
                seen[part] = desc

    with out_path.open("w", newline="", encoding="utf-8") as f:
        print(f"Outputting to {out_path}")
        writer = csv.writer(f)
        writer.writerow(["Part Number", "Part Description"])

        for part, desc in seen.items():
            writer.writerow([part, desc])

    print(f"[OK] Build scraped CSV with {len(seen)} unique parts")

if __name__ == "__main__":
    input_path = Path("Rules") / "YOUR_JSONL_HERE"
    out_path = Path("Rules") / "YOUR_CSV_HERE"
    print("ABOUT TO CONVERT")
    jsonl_to_csv(input_path, out_path)