import os
print("RUNNING:", __file__)
print("CWD:", os.getcwd())
import re
import pdfplumber
from openpyxl import Workbook
import sys
from pathlib import Path

DEBUG = True
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
            [CF]\d{1,2}\s*-\s*[A-Z]\s*-\s*[A-Z0-9]{1,2} |} |     # C37- F- 04, F06- D- O1
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


def extract_items(page_text: str):
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

        part = f"A-{item}"
        part_display = f"{part} ({desc})" if desc else part

        if DEBUG:
            print("ITEM:", qty, part, "|", desc)

        rows.append(
            {
                "Amount": qty,
                "Type": "",
                "Part #": part_display,   # A-XXXX (DESC)
                "P.O. Number": "",
                "Notes": "",              # blank
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

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            po = extract_po(text)

            items = extract_items(text)
            for r in items:
                r["P.O. Number"] = po
            all_rows.extend(items)

    write_output(all_rows, pdf_path.stem + "_output.xlsx")


if __name__ == "__main__":
    main()