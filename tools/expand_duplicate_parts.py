import pandas as pd
from pathlib import Path

input_file = Path("duplicate_parts.csv")
output_file = Path("duplicate_parts_expanded.csv")

df = pd.read_csv(input_file)

rows = []

for _, r in df.iterrows():
    part = str(r["part_number"]).strip()
    types = str(r["types"]).split("+")

    for t in types:
        t = t.strip()
        if t:
            rows.append({
                "part_number": part,
                "type": t
            })

expanded = pd.DataFrame(rows)

expanded = expanded.sort_values(["part_number", "type"])

expanded.to_csv(output_file, index=False)

print(f"Wrote {output_file}")