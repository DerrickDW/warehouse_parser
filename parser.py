import re
import pdfplumber
from openpyxl import Workbook

PO_RE = re.compile(r"PO\s*#\s*(\d+)", re.IGNORECASE)
ITEM_RE = re.compile(r"\bQty\s+(\d+)\s+([A-Z0-9\-]+)\b", re.IGNORECASE)


def extract_po(page_text: str) -> str | None:
    if not page_text:
        return None
    m = PO_RE.search(page_text)
    return m.group(1) if m else None
    # extract items


def extract_items_debug(page_text):
    lines = page_text.splitlines()
    for line in lines:
        s = line.strip()
        if not s:
            continue
        print(s)


def extract_items(page_text: str):
    """
    Returns list of dict rows: Amount, Part #, PO
    """
    rows = []
    for m in ITEM_RE.finditer(page_text or ""):
        qty = int(m.group(1))
        part = m.group(2).strip()
        desc = (m.group(3) or "").strip()
        # optional keep a little description if it exists (can remove if noisy)
        part_display = f"{part} ({desc})" if desc else part
        rows.append(
            {
                "Amount": qty,
                "Type": "",  # manual fill
                "Part #": part_display,
                "P.O. Number": "",  # filled later per page
                "Notes": "",
                "Boxes/PC": "",
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
        print(f"Total rows captured: {len(all_rows)}")
        for idx, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ""
            print("\n--- FIRST 80 LINES---\n")
            for i, line in enumerate(text.splitlines()[:80], start=1):
                print(f"{i:02d}: {line}")
            break
            items = extract_items(text)
            po = extract_po(text)
        for r in items:
            r["P.O. Number"] = po or ""
            all_rows.append(r)
            print(f"Page {idx}: PO={po}")


if __name__ == "__main__":
    main()
