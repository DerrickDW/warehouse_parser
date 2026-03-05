#!/usr/bin/env python3
"""
mine_part_rules.py

Mines historical Excel sheets to generate:
  1) valid_part_numbers.csv
  2) part_number_frequency.csv
  3) duplicate_parts.csv  (parts that frequently appear with multiple types)

Assumptions:
- Part number is in a "Part Number" / "Item" / "Part" type column (names vary)
- Type is in its own column called "Type" (names may vary slightly)
- Quantity is ignored
- Anything in parentheses after the part number should be ignored, e.g. "A-12345 (Part)" -> "A-12345"

Run:
  pip install pandas openpyxl
  python mine_part_rules.py --input-dir "C:\path\to\excels" --out-dir "C:\path\to\out" --recursive
"""

import argparse
import re
from pathlib import Path
from collections import defaultdict

import pandas as pd


# -----------------------------
# Regex + normalization helpers
# -----------------------------

CELL_RE = re.compile(r"^\s*(?P<part>[^()]+?)\s*(?:\((?P<ptype>[^()]*)\))?\s*$")
PART_LIKE_RE = re.compile(r"[A-Z0-9]+(?:[-/][A-Z0-9]+)+|[A-Z]{1,6}\d{2,}", re.IGNORECASE)
HYPHENS_RE = re.compile(r"[\u2010\u2011\u2012\u2013\u2014\u2212]")


def normalize_part_number(s: str) -> str:
    if s is None:
        return ""
    s = str(s).strip()
    if not s:
        return ""
    s = HYPHENS_RE.sub("-", s)
    s = re.sub(r"\s+", "", s)
    s = s.upper()
    s = re.sub(r"-{2,}", "-", s)
    return s


def normalize_type(s: str) -> str:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    s = str(s).strip()
    if not s or s.lower() in {"nan", "none"}:
        return ""
    s = HYPHENS_RE.sub("-", s)
    s = re.sub(r"\s+", "", s)
    return s.upper()


def parse_part_cell(cell) -> str:
    """
    Returns the normalized part number portion only.
    Strips anything in parentheses.
    """
    if cell is None or (isinstance(cell, float) and pd.isna(cell)):
        return ""
    text = str(cell).strip()
    if not text:
        return ""

    m = CELL_RE.match(text)
    if m:
        part = (m.group("part") or "").strip()
    else:
        part = text.split("(", 1)[0].strip()

    return normalize_part_number(part)


# -----------------------------
# Column detection
# -----------------------------

def score_part_column(series: pd.Series, sample_n: int = 40) -> float:
    vals = series.dropna().astype(str).head(sample_n).tolist()
    if not vals:
        return 0.0

    hits = 0
    total = 0
    for v in vals:
        v = v.strip()
        if not v:
            continue
        total += 1
        head = v.split("(", 1)[0].strip()
        if PART_LIKE_RE.search(head):
            hits += 1

    return (hits / total) if total else 0.0


def find_part_column(df: pd.DataFrame) -> str | None:
    if df is None or df.empty:
        return None

    cols = list(df.columns)

    preferred = []
    for c in cols:
        name = str(c).strip().lower()
        if any(k in name for k in ["part", "item", "p/n", "pn", "part number", "part#", "item#"]):
            if not any(bad in name for bad in ["qty", "quantity", "description", "desc", "type", "uom", "price"]):
                preferred.append(c)

    if len(preferred) == 1:
        return preferred[0]
    if len(preferred) > 1:
        scored = [(c, score_part_column(df[c])) for c in preferred]
        scored.sort(key=lambda x: x[1], reverse=True)
        return scored[0][0] if scored and scored[0][1] > 0 else preferred[0]

    scored_all = []
    for c in cols:
        name = str(c).strip().lower()
        if any(bad in name for bad in ["qty", "quantity", "description", "desc", "type", "notes", "note", "uom", "price"]):
            continue
        scored_all.append((c, score_part_column(df[c])))

    if not scored_all:
        return None

    scored_all.sort(key=lambda x: x[1], reverse=True)
    best_col, best_score = scored_all[0]
    return best_col if best_score >= 0.25 else None


def find_type_column(df: pd.DataFrame) -> str | None:
    if df is None or df.empty:
        return None

    cols = list(df.columns)

    candidates = []
    for c in cols:
        name = str(c).strip().lower()
        if name in {"type", "ty", "t"}:
            candidates.append(c)
            continue
        if any(k in name for k in ["type", "line type", "item type", "label type"]):
            candidates.append(c)

    filtered = []
    for c in candidates:
        name = str(c).strip().lower()
        if any(bad in name for bad in ["qty", "quantity", "part", "item", "desc", "description", "price", "uom"]):
            continue
        filtered.append(c)

    if len(filtered) == 1:
        return filtered[0]

    if len(filtered) > 1:
        scored = []
        for c in filtered:
            s = df[c].dropna().astype(str).str.strip()
            s = s[s != ""]
            if s.empty:
                continue
            scored.append((c, s.nunique()))
        scored.sort(key=lambda x: x[1])
        return scored[0][0] if scored else filtered[0]

    return None


# -----------------------------
# File reading
# -----------------------------

def iter_excel_files(root: Path, recursive: bool) -> list[Path]:
    patterns = ["*.xlsx", "*.xls", "*.xlsm", "*.xlsb"]
    files = []
    if recursive:
        for pat in patterns:
            files.extend(root.rglob(pat))
    else:
        for pat in patterns:
            files.extend(root.glob(pat))
    files = [f for f in files if not f.name.startswith("~$")]
    return sorted(set(files))


def read_workbook_safely(path: Path) -> dict[str, pd.DataFrame]:
    try:
        xls = pd.ExcelFile(path)
        out = {}
        for sheet in xls.sheet_names:
            try:
                out[sheet] = pd.read_excel(xls, sheet_name=sheet)
            except Exception:
                continue
        return out
    except Exception:
        return {}


# -----------------------------
# Mining
# -----------------------------

def mine_rules(
    input_dir: Path,
    out_dir: Path,
    recursive: bool,
    min_part_len: int,
    min_dupe_docs: int,
    min_dupe_rate: float,
    require_type_for_dupes: bool,
    write_mapping_report: bool,
    append_mode: bool,
):
    files = iter_excel_files(input_dir, recursive=recursive)

    # Overall row frequency:
    part_row_counts = defaultdict(int)

    # Duplicate detection (doc = file+sheet):
    docs_with_part = defaultdict(int)
    docs_with_multitype = defaultdict(int)
    overall_types = defaultdict(set)

    total_docs = 0
    processed_files = 0
    mapping_rows = []

    for fp in files:
        wb = read_workbook_safely(fp)
        if not wb:
            continue
        processed_files += 1

        for sheet_name, df in wb.items():
            if df is None or df.empty:
                continue

            part_col = find_part_column(df)
            type_col = find_type_column(df)

            if write_mapping_report:
                mapping_rows.append({
                    "file": fp.name,
                    "sheet": sheet_name,
                    "part_col": str(part_col) if part_col else "",
                    "type_col": str(type_col) if type_col else "",
                    "rows": int(len(df)),
                })

            if part_col is None:
                continue

            total_docs += 1
            doc_part_types = defaultdict(set)

            for _, row in df.iterrows():
                part = parse_part_cell(row.get(part_col))
                if not part or len(part) < min_part_len:
                    continue

                part_row_counts[part] += 1

                # ensure presence in this doc
                _ = doc_part_types[part]

                if type_col is not None:
                    t = normalize_type(row.get(type_col))
                    if t:
                        doc_part_types[part].add(t)
                        overall_types[part].add(t)

            for part, typeset in doc_part_types.items():
                if require_type_for_dupes and type_col is None:
                    continue
                docs_with_part[part] += 1
                if len(typeset) >= 2:
                    docs_with_multitype[part] += 1

    out_dir.mkdir(parents=True, exist_ok=True)

    # -----------------------------
    # BUILD duplicate_parts "rows"
    # -----------------------------
    rows = []
    for part, total in docs_with_part.items():
        multi = docs_with_multitype.get(part, 0)
        rate = (multi / total) if total else 0.0

        if (multi >= min_dupe_docs) or (rate >= min_dupe_rate):
            types = sorted(overall_types.get(part, set()))
            if len(types) >= 2:
                rows.append({"part_number": part, "types": "+".join(types)})

    # -----------------------------
    # VALID PARTS (append-aware)
    # -----------------------------
    valid_parts = set(part_row_counts.keys())
    valid_path = out_dir / "valid_part_numbers.csv"

    if append_mode and valid_path.exists():
        old = pd.read_csv(valid_path)
        if "part_number" in old.columns:
            old_parts = set(old["part_number"].astype(str).str.strip())
            valid_parts |= old_parts

    pd.DataFrame({"part_number": sorted(valid_parts)}).to_csv(valid_path, index=False)

    # -----------------------------
    # PART FREQUENCY (append-aware)
    # -----------------------------
    freq_new = pd.DataFrame(
        [{"part_number": p, "count": c} for p, c in part_row_counts.items()]
    )
    freq_path = out_dir / "part_number_frequency.csv"

    if append_mode and freq_path.exists():
        old = pd.read_csv(freq_path)
        combined = pd.concat([old, freq_new], ignore_index=True)
        combined["part_number"] = combined["part_number"].astype(str).str.strip()
        combined["count"] = pd.to_numeric(combined["count"], errors="coerce").fillna(0).astype(int)

        freq_df = (
            combined.groupby("part_number", as_index=False)["count"]
            .sum()
            .sort_values(["count", "part_number"], ascending=[False, True])
        )
    else:
        freq_df = freq_new.sort_values(["count", "part_number"], ascending=[False, True])

    freq_df.to_csv(freq_path, index=False)

    # -----------------------------
    # DUPLICATE PARTS (append-aware)
    # -----------------------------
    dupe_new = pd.DataFrame(rows)
    dupe_path = out_dir / "duplicate_parts.csv"

    if append_mode and dupe_path.exists():
        old = pd.read_csv(dupe_path)
        combined = pd.concat([old, dupe_new], ignore_index=True)

        # merge type sets per part_number
        merged: dict[str, set[str]] = {}
        for _, r in combined.iterrows():
            p = normalize_part_number(r.get("part_number"))
            if not p:
                continue
            types_raw = str(r.get("types") or "")
            tset = {normalize_type(t) for t in types_raw.split("+") if normalize_type(t)}
            if not tset:
                continue
            merged.setdefault(p, set()).update(tset)

        final_rows = [{"part_number": p, "types": "+".join(sorted(ts))} for p, ts in merged.items()]
        dupe_df = pd.DataFrame(final_rows).sort_values("part_number")
    else:
        dupe_df = dupe_new.sort_values("part_number") if not dupe_new.empty else pd.DataFrame(columns=["part_number", "types"])

    dupe_df.to_csv(dupe_path, index=False)
    
    # -----------------------------
# EXPANDED DUPLICATE RULES
# -----------------------------

    expanded_rows = []

    for _, r in dupe_df.iterrows():
        part = str(r["part_number"]).strip()

        for t in str(r["types"]).split("+"):
            t = t.strip()
            if t:
                expanded_rows.append({
                    "part_number": part,
                    "type": t
                })

    expanded_df = pd.DataFrame(expanded_rows)

    expanded_df = expanded_df.sort_values(["part_number", "type"])

    expanded_df.to_csv(out_dir / "duplicate_parts_expanded.csv", index=False)

    # Optional mapping report
    if write_mapping_report:
        pd.DataFrame(mapping_rows).to_csv(out_dir / "detected_columns_report.csv", index=False)

    print("Done.")
    print(f"Files scanned: {len(files)} | Files processed: {processed_files} | Docs (file+sheet) analyzed: {total_docs}")
    print(f"Unique parts found (this run + append): {len(valid_parts)}")
    print(f"Wrote: {valid_path}")
    print(f"Wrote: {freq_path}")
    print(f"Wrote: {dupe_path}")
    if write_mapping_report:
        print(f"Wrote: {out_dir / 'detected_columns_report.csv'}")


def main():
    ap = argparse.ArgumentParser(description="Mine historical Excel sheets to build rule datasets for a warehouse parser.")
    ap.add_argument("--input-dir", required=True, help="Directory containing historical Excel files")
    ap.add_argument("--out-dir", default="out_rules", help="Output directory for CSV files")
    ap.add_argument("--recursive", action="store_true", help="Scan subdirectories")
    ap.add_argument("--min-part-len", type=int, default=4, help="Ignore very short tokens (default: 4)")

    ap.add_argument("--min-dupe-docs", type=int, default=3,
                    help="Min docs with >=2 types to flag duplicates (default: 3)")
    ap.add_argument("--min-dupe-rate", type=float, default=0.30,
                    help="Min fraction of docs with >=2 types to flag duplicates (default: 0.30)")

    ap.add_argument("--require-type-for-dupes", action="store_true",
                    help="If set, sheets without a detected Type column won't count toward duplicate stats.")
    ap.add_argument("--mapping-report", action="store_true",
                    help="Write detected_columns_report.csv to debug templates / header variations.")
    ap.add_argument("--append-mode", action="store_true",
                    help="Append/merge results with existing CSVs instead of overwriting.")

    args = ap.parse_args()

    mine_rules(
        input_dir=Path(args.input_dir),
        out_dir=Path(args.out_dir),
        recursive=args.recursive,
        min_part_len=args.min_part_len,
        min_dupe_docs=args.min_dupe_docs,
        min_dupe_rate=args.min_dupe_rate,
        require_type_for_dupes=args.require_type_for_dupes,
        write_mapping_report=args.mapping_report,
        append_mode=args.append_mode,
    )


if __name__ == "__main__":
    main()