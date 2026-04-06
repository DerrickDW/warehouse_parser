# CHANGELOG

All notable changes to the Warehouse PDF Parser project will be documented in this file.

This project follows a pragmatic versioning model where minor versions represent meaningful improvements to parsing accuracy, rule generation, and validation layers.

---

## V2.0
  - Added GUI
  - Added More Rules (Need to add some examples yet.)
  - Added (Description) editing rules to parser and GUI
  - Fixed bug that made GUI not show Excel preview

## V1.6 Rule Engine Update
  - Added Rule Update that now adds what was done to the part and updated duplication logic as well as confidence scoring.

## v1.5 — Rule Engine + Validation Layer (Current)

### Added

- **Part validation system**
  - Parser now validates extracted parts against `valid_part_numbers.csv`
  - Prevents OCR garbage from silently entering output datasets

- **Correction mapping system**
  - Added `part_corrections.csv` for deterministic OCR correction rules
  - Automatically fixes known OCR misreads before validation

- **Correction audit logging**
  - Unknown export now records both:
    - `part_number_before_correction`
    - `part_number_norm`
  - Allows verification of correction rule effectiveness

- **Unknown parts export**
  - New `_unknown_parts.csv` output generated per parsed document
  - Includes:
    - normalized part
    - original OCR part
    - display part
    - PO number
    - source PDF
    - raw OCR line

- **Validation tolerant of A-prefix variations**
  - Validator now checks both:
    - `A-12345`
    - `12345`

- **Description cleanup improvements**
  - Added `DESC_TRAILING_MFGNUM_RE` to remove trailing manufacturer numbers
  - Improved removal of packaging artifacts (`BOX`, `LABEL`, `PACK`, etc.)

- **Correction-aware Excel output**
  - Corrected part numbers now propagate to final Excel output

---

## v1.4 — OCR Normalization Improvements

### Added

- OCR glyph correction logic:
  - `€ → S`
  - `$ → S`
  - `£ → L`

- Automatic fix for leading OCR error:
  - `8P → BP`

- Internal OCR cleanup handling cases like:
  - `C€S54816`
  - `CSS → CS`

### Improved

- `normalize_item_token()` expanded to handle common OCR artifacts
- Reduced incorrect part token extraction

---

## v1.3 — Description Parsing Improvements

### Added

- `clean_desc()` pipeline for description normalization

### Improved

Removal of trailing packaging artifacts:

- `BOX`
- `LABEL`
- `BAG`
- `PACK`
- `CASE QTY`

Also removes OCR bleed tokens and trailing numeric artifacts.

### Result

Significant reduction in garbage description text entering Excel output.

---

## v1.2 — Robust Part Extraction

### Added

Regex-based item extraction:
Qty + A- + Line Item

Pattern:
^\s*(qty)\s+A-\s*(item)\s*(description)

### Improvements

- Reliable parsing of OCR text from NAPS2 PDFs
- Reduced misidentification of description text as part numbers

---

## v1.1 — Initial OCR Token Normalization

### Added

- Basic token cleanup
- Whitespace normalization
- Hyphen normalization

### Improved

Handling of malformed OCR tokens.

---

## v1.0 — Initial Parser Release

### Features

Reads OCR text layer from NAPS2-generated PDFs.

Extracts:

- Quantity
- Part number
- Description
- PO number

Outputs structured Excel file with columns:
Amount
Type
Part #
P.O. Number
Notes
Boxes/PC

### Purpose

Initial prototype for automating warehouse logistics document transcription.

---

## Planned for v2.0

### Planned Features

- Intelligent fuzzy correction using valid part corpus
- Parser GUI wrapper for drag-and-drop operation ✅ Complete
- Rule mining integration pipeline ✅ Complete

---

## Project Goals

The parser aims to achieve:

- Near-zero manual transcription
- Deterministic OCR cleanup
- Self-improving rule datasets
- High reliability in warehouse logistics workflows