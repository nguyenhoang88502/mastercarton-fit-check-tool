You are Worker A: CODE GENERATION SPECIALIST (HTML UI framework).

CONSTRAINTS YOU MUST OBEY:
- Pure HTML5 + CSS3 + vanilla JavaScript in ONE self-contained file.
- NO external libraries, frameworks, CDNs, fonts, or images. No React/Vue/jQuery/axios.
- Must work offline, opened from http://localhost:3000/ served by a PowerShell HttpListener.
- Total file under 500KB.
- Must run in Chrome/Edge on stock Windows 10/11.

TASK
Generate the complete `index.html` browser UI for the WVN Master Carton Fit Tool data-entry
listener. It talks to a PowerShell HTTP listener on the SAME origin (use relative URLs like
`/api/data/fg-dimensions`, never absolute `http://localhost:3000/...`).

TAB INTERFACE — exactly five tabs:
1. "New FG"      -> POST /api/data/fg-dimensions
2. "New Carton"  -> POST /api/data/carton-dimensions
3. "New BOM"     -> POST /api/data/bom-assignments
4. "Lookup"      -> GET  /api/data/fg-dimensions | /api/data/carton-dimensions | /api/data/bom-assignments
5. "Batch"       -> POST an ARRAY of records to any of the three endpoints (paste JSON or CSV)

ENDPOINT CONTRACTS (fixed — do not invent different field names)

POST /api/data/fg-dimensions
  body: { "fg_id": "FG-001234", "length_mm": 150.5, "width_mm": 100.0,
          "height_mm": 80.0, "target_qty": 48, "source": "MANUAL_ENTRY" }
  fg_id      required, max 50 chars, ^[A-Za-z0-9_-]+$
  length_mm / width_mm / height_mm  required, numeric, 10-1000
  target_qty optional, numeric integer, 1-10000
  source     enum: MANUAL_ENTRY | D365_IMPORT | MOBILE_APP | BATCH_IMPORT (default MANUAL_ENTRY)

POST /api/data/carton-dimensions
  body: { "carton_id": "1008609", "length_mm": 200.0, "width_mm": 150.0,
          "height_mm": 120.0, "dimension_source": "BOM_LINKED_CA_AVERAGE" }
  carton_id  required, max 50 chars, ^[A-Za-z0-9_-]+$
  length_mm / width_mm / height_mm  required, numeric, 10-1000
  dimension_source enum: BOM_LINKED_CA_AVERAGE | MANUAL_MEASUREMENT | SUPPLIER_SPEC

POST /api/data/bom-assignments
  body: { "fg_id": "FG-001234", "carton_id": "1008609",
          "source_bom_date": "2026-08-04", "confidence": "HIGH" }
  fg_id, carton_id  required (server checks referential integrity)
  source_bom_date   required, ISO YYYY-MM-DD
  confidence        enum: HIGH | MEDIUM | LOW (default MEDIUM)

SERVER RESPONSES you must render:
  200 -> { "success": true, "record_id": "...", "timestamp": "...", "csv_updated": true, "csv_path": "..." }
  400 -> { "success": false, "status": 400, "error": "Validation failed",
           "details": [ { "field": "length_mm", "message": "Must be 10-1000 mm" } ] }
  409 -> { "success": false, "status": 409, "error": "Duplicate key", "detail": "..." }
  503 -> { "success": false, "status": 503, "error": "CSV file locked. Close Excel and retry." }
  Batch POST returns { "success": true, "inserted": N, "failed": M,
                       "results": [ { "index": 0, "status": 200, "success": true, ... } ] }
  GET list returns  { "success": true, "count": N, "rows": [ { ...csv columns... } ] }

REQUIRED BEHAVIOUR
- Client-side validation mirroring every server rule above, run on submit AND live on `input`.
  Show the allowed range inline (e.g. "10-1000 mm") and mark invalid fields red.
- Submit via `fetch()` with `Content-Type: application/json`. Never use XMLHttpRequest.
- Loading spinner (pure CSS) shown while a request is in flight; submit button disabled.
- Render success messages green, errors red. For 400, list every `details[].field: message`.
- Collapsible "Raw response" `<details>` block showing the pretty-printed JSON, for debugging.
- Clear the form after a successful POST; keep values on failure.
- Offline queue: on network failure (fetch rejects / server down) push the payload into
  `localStorage` under key `wvn_pending_queue`, show "Queued (offline)", and on page load
  automatically retry every queued item in order, removing each one that succeeds. Show a
  queue badge with the pending count.
- Network timeout of 10s per request using AbortController.
- Lookup tab: dropdown to choose the dataset, "Refresh" button, results rendered in a
  scrollable <table> with a text filter box; show "No records yet" when empty.
- Batch tab: <textarea> accepting either a JSON array or CSV-with-header, a dataset dropdown,
  a "Validate only" checkbox, and a per-row result table after submit.
- Keep-alive: POST `/api/ping` every 10 seconds while the page is open so the short-lived
  listener does not idle out mid-typing. On `pagehide`, call
  `navigator.sendBeacon('/api/shutdown')` so the listener exits and releases file handles.
- Responsive: usable at 360px width (stack form fields, horizontally scrollable tabs).
- Accessible: <label for>, aria-live="polite" on the message area, visible focus rings.
- Dark, high-contrast industrial theme is fine; use CSS custom properties.
- ZERO uncaught console errors.

DELIVERY FORMAT
Output ONE fenced code block tagged `html` containing the complete file, starting with
`<!DOCTYPE html>` and ending with `</html>`. No commentary before or after the block except a
short NOTES section listing assumptions.
