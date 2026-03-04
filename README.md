📦 Warehouse PDF → Excel Parser

A local-first Python tool that extracts structured item data from warehouse Receiving Purchase Order PDFs and exports clean, audit-ready Excel files.

Designed for printed warehouse documents with consistent formatting.

✅ Features

Extracts:

Quantity

Part Number (A- format only)

Description (cleaned)

PO Number

Removes:

Pack/label/tag junk

Right-column Mfg# bleed

OCR glyph noise ($ → S, £ → L, etc.)

Outputs:

Clean .xlsx file using openpyxl

CLI usage:

python parser.py <pdf_path>
📂 Example

Input line (raw PDF):

220 A- R26608 WASHER, PULL ARM FRONT PAW806 BAG 4X4 C37- F- 04

Output row:

Amount	Part #	Description	PO
220	A-R26608	WASHER, PULL ARM FRONT	613938
🚀 Installation
1. Clone repo
git clone <repo-url>
cd warehouse_parser
2. Create virtual environment
python -m venv .venv

Activate:

Windows:

.venv\Scripts\activate

Mac/Linux:

source .venv/bin/activate
3. Install dependencies
pip install -r requirements.txt
📋 Usage
python parser.py "path/to/ReceivingPO.pdf"

Output:

ReceivingPO_output.xlsx

Generated in the same directory as the PDF.

🧠 How It Works

Uses pdfplumber to extract page text.

Detects PO number using regex.

Extracts item lines matching:

Qty A- PartNumber Description

Cleans descriptions by removing:

BAG 4X4

LABEL 36

TAG 48

BX 14

W/ BEARING

Right-column Mfg codes

Writes structured rows using openpyxl.

⚠️ Assumptions

Only parts starting with A- are valid items.

PDFs must contain a readable text layer.

Handwritten scans are not guaranteed to extract correctly.

Best results come from scanning before documents are written on.

🛠️ Project Status

Phase 1 – Core Extraction: ✅ Complete
Stable extraction for printed warehouse documents.

Future ideas (optional):

CSV rules override layer

OCR fallback for image-only PDFs

Historical part duplication detection

Web UI wrapper

🔒 Git Policy

Ignored:

.venv/

*.pdf

*_output.xlsx

IDE folders

Only source code and config are versioned.

📦 Dependencies

pdfplumber

openpyxl

🏗️ Architecture Philosophy

Local-first

No cloud dependencies

Deterministic parsing

Regex bounded loops to prevent hangs

Clean separation between extraction and future rule logic