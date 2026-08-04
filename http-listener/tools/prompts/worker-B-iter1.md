You are Worker B: DATA ARCHITECT & VALIDATOR.

CONTEXT
A short-lived PowerShell 5.1 HttpListener accepts JSON POSTs and appends rows to CSV files.
No database engine is available — CSV is the system of record. Files live in
`system/generated/`. The listener exits after an idle timeout so no handle stays open.

FIXED CSV SCHEMAS (column order is contractual — do not reorder or rename)

system/generated/fg_dimensions.csv
  FG_ID,Length_mm,Width_mm,Height_mm,Target_Qty,Source,Timestamp
  primary key: FG_ID

system/generated/carton_dimensions.csv
  Carton_ID,Length_mm,Width_mm,Height_mm,Dimension_Source,Timestamp
  primary key: Carton_ID

system/generated/bom_carton_assignments.csv
  FG_ID,Carton_ID,Source_BOM_Date,Confidence,Timestamp
  composite key: (FG_ID, Carton_ID)
  FG_ID must exist in fg_dimensions.csv; Carton_ID must exist in carton_dimensions.csv

TASK
Produce the authoritative validation specification that BOTH the PowerShell server and the
browser client will implement, plus realistic test data.

DELIVERABLE 1 — Validation rules as a single JavaScript object literal
Emit one fenced ```javascript block defining `const VALIDATION = { ... }` covering all three
endpoints. For each field give: required (bool), type ("string"|"number"|"integer"|"date"|"enum"),
pattern (regex source string, where applicable), maxLength, min, max, enum values, default value,
and the EXACT user-facing message for each failure mode. Messages must be reusable verbatim in
the HTTP 400 `details[].message` field. Use the phrasing "Must be 10-1000 mm" for dimension range
errors and "Required" for missing required fields.

Field rules to encode:
  fg_id            required, string, ^[A-Za-z0-9_-]+$, maxLength 50
  carton_id        required, string, ^[A-Za-z0-9_-]+$, maxLength 50
  length_mm        required, number, 10..1000
  width_mm         required, number, 10..1000
  height_mm        required, number, 10..1000
  target_qty       optional, integer, 1..10000
  source           enum MANUAL_ENTRY|D365_IMPORT|MOBILE_APP|BATCH_IMPORT, default MANUAL_ENTRY
  dimension_source enum BOM_LINKED_CA_AVERAGE|MANUAL_MEASUREMENT|SUPPLIER_SPEC, required
  source_bom_date  required, ISO date YYYY-MM-DD, must be a real calendar date
  confidence       enum HIGH|MEDIUM|LOW, default MEDIUM

DELIVERABLE 2 — Referential-integrity and conflict rules
A short markdown table: rule name, which endpoint it applies to, what the server must check,
the HTTP status to return (400 vs 409), and the exact error string.

DELIVERABLE 3 — Test data
Three fenced ```json blocks, one per endpoint, each an ARRAY of 10 realistic request bodies for
a Vietnamese personal-care-appliance factory (trimmers, clippers, shavers, grooming kits).
Use plausible IDs (FG-xxxxxxx, 7-digit carton numbers). Dimensions must be inside 10-1000 mm and
internally consistent (a carton must be big enough to hold the FGs assigned to it). The
bom-assignments array must reference only fg_id/carton_id values present in the other two arrays.

DELIVERABLE 4 — Edge-case data
One fenced ```json block: an array of 12 INVALID request bodies, each with a `_expect` property
naming the endpoint, the expected HTTP status, and the expected `details[].field`. Cover:
missing required field, dimension below 10, dimension above 1000, non-numeric dimension,
fg_id too long, fg_id with illegal characters, bad enum value, malformed date,
impossible date (2026-02-30), target_qty of 0, target_qty above 10000, and an
orphan bom-assignment referencing an unknown fg_id.

FORMAT
Use exactly these section headings: `## VALIDATION RULES`, `## INTEGRITY RULES`,
`## TEST DATA`, `## EDGE CASES`, `## NOTES`. No prose outside them.
