You are Worker C: QA & TESTING SPECIALIST.

CONSTRAINTS YOU MUST OBEY:
- Windows PowerShell 5.1 ONLY. No Pester, no external modules, no PowerShell 7 syntax.
  Forbidden: ternary `? :`, null-coalescing `??`, `&&`/`||` chaining, `-AsHashtable`,
  `ForEach-Object -Parallel`, classes with typed generics.
- Runs on stock Windows 10/11 with nothing installed.
- No Python, Node, or .NET Core.

SYSTEM UNDER TEST
`listen.ps1` — a short-lived PowerShell HttpListener started as:
  powershell -NoProfile -ExecutionPolicy Bypass -File listen.ps1 -Port <port>
      -IdleTimeoutSeconds <n> -MaxRequests <n> -DataPath <dir> -LogPath <dir>
It serves `index.html` at `/` and exposes:
  GET  /api/health                    -> { "success": true, "status": "listening", ... }
  POST /api/ping                      -> keep-alive, does not count toward MaxRequests
  POST /api/shutdown                  -> graceful exit
  POST /api/data/fg-dimensions        -> append row to <DataPath>/fg_dimensions.csv
  POST /api/data/carton-dimensions    -> append row to <DataPath>/carton_dimensions.csv
  POST /api/data/bom-assignments      -> append row to <DataPath>/bom_carton_assignments.csv
  GET  /api/data/<same three>         -> { "success": true, "count": N, "rows": [...] }
POST bodies may be a single JSON object OR a JSON array (batch).
Transaction log: <LogPath>/listener.log

CSV SCHEMAS (exact column order)
  fg_dimensions.csv            FG_ID,Length_mm,Width_mm,Height_mm,Target_Qty,Source,Timestamp
  carton_dimensions.csv        Carton_ID,Length_mm,Width_mm,Height_mm,Dimension_Source,Timestamp
  bom_carton_assignments.csv   FG_ID,Carton_ID,Source_BOM_Date,Confidence,Timestamp

VALIDATION CONTRACT
  ids: ^[A-Za-z0-9_-]+$, max 50 chars | dimensions numeric 10-1000 | target_qty integer 1-10000
  source: MANUAL_ENTRY|D365_IMPORT|MOBILE_APP|BATCH_IMPORT
  dimension_source: BOM_LINKED_CA_AVERAGE|MANUAL_MEASUREMENT|SUPPLIER_SPEC
  confidence: HIGH|MEDIUM|LOW | source_bom_date: YYYY-MM-DD
  200 success | 400 validation (has details[]) | 409 duplicate key | 503 CSV locked

TASK
Generate `tests/test.ps1`: a complete, self-contained test suite.

REQUIRED STRUCTURE
1. param(): -Port (default 3097), -ListenerPath, -KeepArtifacts switch. Use a throwaway
   DataPath/LogPath under `tests/_run/` so production CSVs are never touched.
2. A tiny assertion helper set — Assert-Equal, Assert-True, Assert-Match, Assert-FileExists —
   that records results into a script-scope collection instead of throwing, so one failure does
   not abort the run.
3. Start listen.ps1 with Start-Process -PassThru -WindowStyle Hidden, then poll
   GET /api/health for up to 15s before running tests. Fail fast with a clear message if the
   listener never comes up.
4. A `Invoke-Api` helper wrapping Invoke-WebRequest that RETURNS the status code and parsed body
   even for 4xx/5xx. Note: PowerShell 5.1 throws on non-2xx, so you MUST catch
   [System.Net.WebException] and read $_.Exception.Response via a StreamReader.
   Also set -UseBasicParsing on every call.
5. AT LEAST 14 test cases, each printed as `[PASS]`/`[FAIL]` with a name:
   T01 health endpoint returns 200 and success:true
   T02 valid FG post returns 200 and success:true
   T03 fg_dimensions.csv created with the exact expected header line
   T04 fg_dimensions.csv contains the posted FG_ID
   T05 duplicate FG_ID returns 409 and the message contains "already exists"
   T06 dimension below 10 returns 400 with details[].field = "length_mm"
   T07 dimension above 1000 returns 400
   T08 missing required fg_id returns 400 with message "Required"
   T09 bad enum for source returns 400
   T10 malformed JSON body returns 400 (not 500)
   T11 valid carton post returns 200 and carton_dimensions.csv gets the row
   T12 bom-assignment referencing an unknown FG_ID returns 400 mentioning "not found"
   T13 valid bom-assignment (both keys exist) returns 200
   T14 batch post of a 3-element JSON array returns inserted=3
   T15 GET /api/data/fg-dimensions returns count >= 1 and rows[] shaped like the CSV
   T16 NO FILE LOCKING: while the listener is still running, open fg_dimensions.csv with
       [System.IO.File]::Open($p,'Open','ReadWrite','None') — exclusive — and assert it succeeds,
       proving the listener holds no handle. Close it again.
   T17 listener.log exists and contains a line for a POST to /api/data/fg-dimensions
   T18 p95 latency of 10 sequential health calls is under 200ms
6. Shutdown: POST /api/shutdown, wait up to 10s for the process to exit, assert it exited.
   Kill it in a finally block if it did not.
7. Write `tests/_run/test-results.json` with { total, passed, failed, durationMs, results[] }.
8. Print a summary table and `exit 1` when any test failed, `exit 0` otherwise.

DELIVERY FORMAT
Output ONE fenced code block tagged `powershell` containing the complete test.ps1.
Follow it with a short NOTES section listing assumptions and anything you could not verify.
