import re
import pdfplumber
from openpyxl import Workbook

DASHES = "-\u2010\u2011\u2013\u2014\u2212"  # -, -,

PO_RE = re.compile(r"PO\s*#\s*(\d+)", re.IGNORECASE)
# ITEM_RE = re.compile(r"\bQty\s+(\d+)\s+([A-Z0-9\-]+)\b", re.IGNORECASE)
# ITEM_RE = re.compile(r"^\s*(\d+)\s+[A-Z]-\s+([A-Z0-9]+)", re.MULTILINE)
# ITEM_RE = re.compile(r"^\s*(\d+)\s+[A-Z]-\s+([A-Z0-9]+)")
ITEM_RE = re.compile(r"^\s*(\d+)\s+[A-Z]?\s*([A-Z0-9]+)\s+(.+?)\s*$", re.MULTILINE)


def extract_po(page_text: str) -> str | None:
    if not page_text:
        return None
    m = PO_RE.search(page_text)
    return m.group(1) if m else None
    # extract items

    # def extract_items_debug(page_text):
    lines = page_text.splitlines()
    for line in lines:
        s = line.strip()
        if not s:
            continue
        # print only lines that look like they contain a part number
        if any(ch.isdigit() for ch in s) and len(s) > 10:
            print("CANDIDATE:", s)
    # return []

    # def extract_items(page_text: str):
    rows = []
    text = page_text or ""
    # lines = page_text.splitlines()
    # for m in ITEM_RE.finditer(page_text or ""):
    # print("MATCH:", m.group(1), m.group(2))

    # return []
    # temp check
    for ln in (l.strip() for l in (page_text or "").splitlines()):
        if ln and any(ch.isdigit() for ch in ln):
            print("LINE:", ln)
    print("extract_items CALLED, chars =", len(page_text))
    """
    Returns list of dict rows: Amount, Part #, PO
    """
    rows = []


def extract_items(page_text):
    # lines = page_text.splitlines()
    # for s in (ln.strip() for ln in lines):
    # if not s:
    # continue
    # if any(k in s.lower() for k in ("shipped to", "ordered from", "authorized by", "ref #", "page", "po #")):
    # continue
    # if re.fullmatch(r"[xX\-\.\s]{10,}\d{2,}[xX\-\.\s]{5.}", s):
    # continue
    # if re.match(r"^\d+\s", s):
    # print("CANIDATE:", s)
    rows = []
    # text = page_text or ""
    # for d in DASHES[1:]:
    # text = text.replace(d, "-")
    # print("chars:", len(text))
    # show any lines that look like they contain an item code fragment
    # hits = []
    # for ln in text.splitlines():
    # s = ln.strip()
    # if not s:
    # continue
    # if "A" in s and "-" in s:  # crude filter
    # hits.append(s)

    # print("lines with A and - :", len(hits))
    # for s in hits[:20]:
    # print("HIT:", s)
    # matches = list(ITEM_RE.finditer(text))
    # print(f"ITEM matches: {len(matches)}")
    # for m in matches[:5]:
    # print("MATCH", m.group(1), m.group(2))
    # for m in matches:
    # qty = int(m.group(1))
    # part = m.group(2).strip()
    # desc = m.group(3) or "".strip
    for m in ITEM_RE.finditer(page_text or ""):
        qty = int(m.group(1))
        part = m.group(2).strip()
        desc = m.group(3).strip()
        # optional keep a little description if it exists (can remove if noisy)
        # part_display = part  # f"{part} ({desc})" if desc else part
        rows.append(
            {
                "Amount": qty,
                # "Type": "",  # manual fill
                "Part #": part,
                "Description": desc,
                # "P.O. Number": "",  # filled later per page
                # "Notes": "",
                # "Boxes/PC": "",
            }
        )
    return rows


# build rows and write output.xlsx
def main():
    path = "China 141.26.pdf"

    all_rows = []

    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            po = extract_items(text)  # like 61418

            items = extract_items(text)
            for r in items:
                r["P.O. Number"] = po or ""
                all_rows.append(r)

    # write excel
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    headers = ["Amount", "Type", "Part #", "P.O. Number", "Notes", "Boxes/PC"]
    ws.append(headers)
    for r in all_rows:
        ws.append([r[h] for h in headers])
    wb.save("output.xlsx")
    print(f"Wrote {len(all_rows)} rows to output.xlsx")


def main() -> None:
    path = "China 141.26.pdf"
    all_rows = []

    with pdfplumber.open(path) as pdf:
        # print(f"Total rows captured: {len(all_rows)}")
        for idx, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ""
            po = extract_po(text) or ""
            # print(f"PAGE {idx} chars={len(text)}")
            # print(f"\n===== PAGE {idx} RAW TEXT PREVIEW =====")
            # print(text[:1000])
            # print("======================================\n")
            items = extract_items(text)
            for r in items:
                all_rows.append(r)
            po = extract_po(text)
            for r in items:
                r["P.O. Number"] = po or ""
                all_rows.append(r)

    print(f"Total rows captured: {len(all_rows)}")
    # print(f"Page {idx}: PO={po}")


if __name__ == "__main__":
    main()
