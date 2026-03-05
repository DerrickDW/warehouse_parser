import os
print("RUNNING:", __file__)
print("CWD:", os.getcwd())
import re
import pdfplumber
from openpyxl import Workbook
import sys
from pathlib import Path
import csv
from datetime import datetime
import pandas as pd

DEBUG = False
if DEBUG:
    print("🐒 systems stable")

# --- PO ---
PO_RE = re.compile(r"\bPO\s*#?\s*(\d+)\b", re.IGNORECASE)

def extract_po(page_text: str) -> str:
    if not page_text:
        return ""
    m = PO_RE.search(page_text)
    return m.group(1) if m else ""


# --- Items (core accuracy: Qty + A- + Line/Item only) ---
ITEM_LINE_RE = re.compile(
    r"^\s*(?P<qty>\d{1,6})\s+A\s*-\s*(?P<item>\S+)\s*(?P<desc>.*\S)?\s*$",
    re.IGNORECASE,
)

DESC_TRAILING_PACK_RE = re.compile(
    r"""
    (?:\s+
        (?:
            BAG\s*\d+\s*[Xx]\s*\d+ |
            BAG\s*&\s*LABEL |
            LABEL\s*(?:\d+|[A-Z]{1,3}) |
            TAG\s*\d+(?:\s*[A-Z])? |
            BX\s*[A-Z0-9]+ |
            BOX\s*\d+ |
            PACK\s+OF\s*\d+ |
            CASE\s*QTY\s*:\s*\d+ |
            W/\s*BEARING |
            [CF]\d{1,2}\s*-\s*[A-Z]\s*-\s*[A-Z0-9]{1,2} |     # C37- F- 04, F06- D- O1
            \d+\.\d+(?:\.[A-Z0-9]+)+                    # 6.05.C38.1K3
        )
    )\s*$
    """,
    re.IGNORECASE | re.VERBOSE,
)

DESC_TRAILING_MFGNUM_RE = re.compile(
    r"\s+\d{4,6}(?:\s+LABEL\s*[A-Z]{1,3})?\s*$",
    re.IGNORECASE,
)

HYPHENS_RE = re.compile(r"[\u2010\u2011\u2012\u2013\u2014\u2212]")

def normalize_part_for_validation(s: str) -> str:
    """Strip parentheses text and normalize spacing/case for dictionary checks only."""
    if s is None:
        return ""
    s = str(s).strip()
    if not s:
        return ""
    s = HYPHENS_RE.sub("-", s)
    s = s.split("(", 1)[0].strip()     # strip (DESC) etc
    s = re.sub(r"\s+", "", s)          # remove whitespace
    s = s.upper()
    s = re.sub(r"-{2,}", "-", s)
    return s

def load_part_corrections(csv_path: Path) -> dict[str, str]:
    if not csv_path.exists():
        print(f"[WARN] corrections file not found: {csv_path} (corrections disabled)")
        return {}

    df = pd.read_csv(csv_path)
    if not {"bad_part", "good_part"}.issubset(df.columns):
        print(f"[WARN] {csv_path} must have columns bad_part,good_part (corrections disabled)")
        return {}

    mapping = {}
    for _, r in df.iterrows():
        bad = normalize_part_for_validation(r.get("bad_part"))
        good = normalize_part_for_validation(r.get("good_part"))
        if bad and good:
            mapping[bad] = good

    print(f"[OK] Loaded {len(mapping)} corrections from {csv_path}")
    return mapping

def load_valid_parts(csv_path: Path) -> set[str]:
    if not csv_path.exists():
        print(f"[WARN] valid parts file not found: {csv_path} (validation disabled)")
        return set()

    df = pd.read_csv(csv_path)
    if "part_number" not in df.columns:
        print(f"[WARN] {csv_path} missing 'part_number' column (validation disabled)")
        return set()

    parts = set()
    for raw in df["part_number"].astype(str).tolist():
        p = normalize_part_for_validation(raw)
        if not p:
            continue
        parts.add(p)

        # ALSO add a variant without leading "A-" if present
        if p.startswith("A-") and len(p) > 2:
            parts.add(p[2:])

        # ALSO add a variant with leading "A-" if missing
        if not p.startswith("A-"):
            parts.add("A-" + p)

    parts.discard("")
    print(f"[OK] Loaded {len(parts)} valid part variants from {csv_path}")
    return parts

def write_unknown_parts_csv(unknown_rows: list[dict], out_path: Path):
    if not unknown_rows:
        print("[OK] No unknown parts found.")
        return

    # de-dupe by normalized part
    seen = set()
    deduped = []
    for r in unknown_rows:
        key = r.get("part_number_norm", "")
        if key and key not in seen:
            seen.add(key)
            deduped.append(r)

    headers = [
    "part_number_norm",
    "part_number_before_correction",
    "part_number_display",
    "po",
    "source_pdf",
    "raw_line",
]
    with out_path.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=headers)
        w.writeheader()
        for r in deduped:
            w.writerow({h: r.get(h, "") for h in headers})

    print(f"[WARN] Unknown parts exported: {out_path} ({len(deduped)} unique)")

def clean_desc(desc: str) -> str:
    d = (desc or "").strip()

    # Normalize whitespace
    d = re.sub(r"\s{2,}", " ", d).strip()

    # OCR cleanup inside description text
    d = d.replace("$", "S")

    # Strip trailing pack/junk tokens (bounded loop prevents hang)
    for _ in range(3):
        new_d = DESC_TRAILING_PACK_RE.sub("", d).strip()
        new_d = re.sub(r"\s{2,}", " ", new_d).strip()
        if new_d == d:
            break
        d = new_d

    # --- NEW: strip trailing manufacturer numbers like " 123456 LABEL A"
    for _ in range(2):
        new_d = DESC_TRAILING_MFGNUM_RE.sub("", d).strip()
        if new_d == d:
            break
        d = new_d

    # Strip up to 2 trailing mfg-style bleed tokens (must contain digit)
    for _ in range(2):
        new_d = re.sub(
            r"\s+(?=[A-Z0-9-]{4,15}\s*$)(?=[A-Z0-9-]*\d)[A-Z0-9-]{4,15}\s*$",
            "",
            d,
            flags=re.IGNORECASE,
        ).strip()
        if new_d == d:
            break
        d = new_d

    # Strip trailing standalone numeric codes (4–8 digits)
    d = re.sub(r"\s+\d{1,2}\.\d{1,2}\s*$", "", d).strip()
    d = d.rstrip(" ,.;:-")

    return d

def normalize_item_token(tok: str) -> str:
    t = (tok or "").strip()

    # strip stray punctuation around token early
    t = t.strip(".,;:()[]{}")

    # 1) Explicit leading glyph fixes FIRST (so they don't get caught by generic rules)
    # £77502 -> L77502
    if t.startswith("£"):
        t = "L" + t[1:]
    # $BA... -> SBA...
    if t.startswith("$"):
        t = "S" + t[1:]
    # €211017 -> C211017
    if t.startswith("€"):
        t = "C" + t[1:]

    # 2) Fix internal OCR glyphs (seen in your output: C€S54816)
    # Treat € inside token as S (common OCR swap you’re seeing)
    t = t.replace("€", "S")
    t = t.replace("CSS", "CS")

    # 3) Fix leading '8' misread as 'B' when it clearly starts a part prefix
    # 8P638... -> BP638...
    if t.startswith("8P"):
        t = "BP" + t[2:]

    # 4) Generic fallback: if it starts with non-alnum glyph(s) then digits only -> C + digits
    # (kept, but now it won't steal £/$/€ cases)
    m = re.match(r"^[^A-Za-z0-9]+(\d{4,10})$", t)
    if m:
        t = "C" + m.group(1)

    return t


def extract_items(page_text: str, valid_parts: set[str], corrections: dict[str, str], unknown_parts: list, po: str, source_pdf: str):
    rows = []
    if not page_text:
        return rows

    for line in page_text.splitlines():
        # skip header line
        if "qty" in line.lower() and "line/item" in line.lower():
            continue

        m = ITEM_LINE_RE.match(line)
        if not m:
            continue

        qty = int(m.group("qty"))

        item_raw = m.group("item")
        item = normalize_item_token(item_raw)

        desc = clean_desc((m.group("desc") or "").strip())

        part_raw = f"A-{item}"  # what OCR gave you (normalized token but not corrected)
        original_part_norm = normalize_part_for_validation(part_raw)

        # Apply correction (if any)
        corrected_part_norm = original_part_norm
        if corrections and corrected_part_norm in corrections:
            corrected_part_norm = corrections[corrected_part_norm]

        # IMPORTANT: use corrected part for Excel output too (keep description!)
        part_for_output = corrected_part_norm  # this is like "A-H135423"
        part_display = f"{part_for_output} ({desc})" if desc else part_for_output

        # try both forms (A-XXX and XXX)
        part_candidates = {corrected_part_norm}
        if corrected_part_norm.startswith("A-") and len(corrected_part_norm) > 2:
            part_candidates.add(corrected_part_norm[2:])
        else:
            part_candidates.add("A-" + corrected_part_norm)

        # log unknowns (but still export normal output row)
        if valid_parts and not any(p in valid_parts for p in part_candidates):
            unknown_parts.append({
                "part_number_norm": corrected_part_norm,
                "part_number_before_correction": original_part_norm,
                "part_number_display": part_display,
                "po": po or "",
                "source_pdf": source_pdf or "",
                "raw_line": line.strip(),
            })

        if DEBUG:
            print("ITEM:", qty, part_for_output, "|", desc)

        rows.append(
            {
                "Amount": qty,
                "Type": "",
                "Part #": part_display,   # KEEP (DESC) in output
                "P.O. Number": "",
                "Notes": "",
                "Boxes/PC": "",
            }
        )

    return rows

def write_output(rows, filename):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    headers = ["Amount", "Type", "Part #", "P.O. Number", "Notes", "Boxes/PC"]
    ws.append(headers)

    for r in rows:
        ws.append([r.get(h, "") for h in headers])

    wb.save(filename)
    print(f"Wrote {len(rows)} rows to {filename}")


def main():
    if len(sys.argv) < 2:
        print("Usage: python parser.py <pdf_path>")
        return

    pdf_path = Path(sys.argv[1])
    if not pdf_path.exists():
        print("File not found:", pdf_path)
        return

    all_rows = []
    
    rules_dir = Path(__file__).resolve().parent / "Rules"  # change to wherever your mined CSVs live
    valid_parts = load_valid_parts(rules_dir / "valid_part_numbers.csv")
    corrections = load_part_corrections(rules_dir / "part_corrections.csv")
    unknown_parts = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            po = extract_po(text)

            items = extract_items(
                text,
                valid_parts=valid_parts,
                corrections=corrections,
                unknown_parts=unknown_parts,
                po=po,
                source_pdf=pdf_path.name,
)
            for r in items:
                r["P.O. Number"] = po
            all_rows.extend(items)

    write_output(all_rows, pdf_path.stem + "_output.xlsx")
    
    unknown_out = pdf_path.with_name(pdf_path.stem + "_unknown_parts.csv")
    write_unknown_parts_csv(unknown_parts, unknown_out)


if __name__ == "__main__":
    main()