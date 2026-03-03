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


# --- Items ---
# Matches lines like:
# 43 A- BP212066766-A IMP. YOKE, ROUND BORE 6.05.C38.1K3
ITEM_RE = re.compile(
    r"^\s*(\d+)\s+([A-Z]-)\s+([A-Z0-9][A-Z0-9-]*)\s+(.+?)\s*$",
    re.IGNORECASE | re.MULTILINE,
)

SKIP_CONTAINS = (
    "confirmed dates",
    "updated/confirmed",
    "updated dates",
    "receiving purchase order",
    "ordered from",
    "authorized by",
    "ref #",
    "final receipt",
    "total weight",
    "shipped to",
    "qty line/item",
    "floor qty",
    "case qty",
    "pack of",
    "pallet",
    "max height",
)

# Things that frequently appear on the RIGHT side and get glued onto description
TRAILING_JUNK_TOKEN_RE = re.compile(
    r"""
    (?:\s+
        (?:
            BAG\s*\d+\s*X\s*\d+ |
            LABEL\s*\d+ |
            BX\s*[A-Z0-9]+ |
            TAG\s*\d+(?:\s*[A-Z])? |
            BOX\s*\d+ |
            W/\s*BEARING |
            C\d{1,2}\s*[-/]\s*[A-Z]\s*[-/]\s*\d{1,2} |   # e.g. C37- F- 04
            \d+\.\d+\.[A-Z0-9.]+ |                       # e.g. 6.05.C38.1K3
            [A-Z]{1,4}\d{3,}[A-Z0-9-]*                  # e.g. SW0618, SEA53287-A, HQC32820
        )
    )\s*$
    """,
    re.IGNORECASE | re.VERBOSE,
)

def clean_desc(desc: str) -> str:
    d = (desc or "").strip()
    d = re.sub(r"\s{2,}", " ", d).strip()

    # Repeatedly strip trailing junk tokens until it stops changing
    while True:
        new_d = re.sub(TRAILING_JUNK_TOKEN_RE, "", d).strip()
        new_d = re.sub(r"\s{2,}", " ", new_d).strip()
        if new_d == d:
            break
        d = new_d

    return d


def extract_items(page_text: str):
    rows = []
    if not page_text:
        return rows

    for m in ITEM_RE.finditer(page_text):
        qty = int(m.group(1))
        prefix = m.group(2).upper()   # "A-"
        code = m.group(3).strip()
        desc = m.group(4).strip()
        part = f"{prefix}{code}"
        low = desc.lower()
        if any(k in low for k in SKIP_CONTAINS):
            continue

        desc_clean = clean_desc(desc)
        #remove duplicate part number
        if desc_clean.upper().endswith(part.upper()):
            desc_clean = desc_clean[: -len(part)].strip()

        part = f"{prefix} {code}"  # keep A- visible
        desc_clean = clean_desc(desc)
        #kill duplicates
        for tail in (part, code, f"{prefix} {code}", f"{prefix}{code}".replace("-","")):
            if desc_clean.upper().endswith(tail.upper()):
                desc_clean = desc_clean[:-len(tail)].strip()
        part_display = f"{part} ({desc_clean})" if desc_clean else part

        if DEBUG:
            print("MATCH:", qty, "|", part_display)

        rows.append(
            {
                "Amount": qty,
                "Type": "",
                "Part #": part_display,
                "P.O. Number": "",  # filled in main()
                "Notes": "",
                "Boxes/PC": "",
            }
        )

    return rows
def write_output (rows, filename):
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
    #allow file finder
    if len(sys.argv) < 2:
        print("Usage: python parser.py <pdf_path>")
        return
    pdf_path = Path(sys.argv[1])
    
    if not pdf_path.exists():
        print("File not found:", pdf_path)
        return
        #path = sys.argb[1]
    #else:
        #path = "China 141.26.pdf"
        #path = str(Path(path).expanduser())
    all_rows = []
    
    #TEMP LOOOP
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            print(f"\n--- PAGE {i+1} RAW TEXT ---\n")
            print(text)
            print("\n--- END PAGE ---\n")
            break  # just first page for now

    #with pdfplumber.open(pdf_path) as pdf:
        #for page in pdf.pages:
            #text = page.extract_text() or ""
            #po = extract_po(text) or ""

            #items = extract_items(text)
            #for r in items:
                #r["P.O. Number"] = po
            #all_rows.extend(items)
    #write_output(all_rows, pdf_path.stem + "_output.xlsx")

   


if __name__ == "__main__":
    main()