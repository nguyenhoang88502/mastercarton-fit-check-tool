# WVN Master Carton Fit Tool - HTTP Listener

Browser-based data entry for FG dimensions, carton dimensions and BOM assignments, backed by a
**short-lived** PowerShell HTTP listener that appends to CSV and then exits.

The point of the design is the exit. Nothing stays resident holding a file handle, so Excel and
the main Fit Tool can open the same CSVs at any moment. There is no "file in use" state to get
stuck in.

## Requirements

Stock Windows 10/11. Nothing to install.

- Windows PowerShell 5.1 (built in)
- `System.Net.HttpListener` (built into .NET Framework)
- Any modern browser

No Python, Node.js, .NET Core, IIS, pip, npm, or SQLite.

## Run it

Double-click **`launcher.cmd`**. It starts the listener, waits until it actually answers, and
opens the UI. To use a different port:

```bash
launcher.cmd 3100
```

Or start the listener by hand:

```bash
powershell -NoProfile -ExecutionPolicy Bypass -File listen.ps1 -Port 3000 -IdleTimeoutSeconds 60
```

Then browse to `http://localhost:3000/`.

## Lifetime

The listener exits when any of these happen:

| Trigger | Behaviour |
|---|---|
| No traffic for `-IdleTimeoutSeconds` (default 30) | Exits |
| `-MaxRequests` data requests handled (default 100) | Exits |
| Browser tab closed | `/api/release`, then exits after a 5s grace period |
| `POST /api/shutdown` | Exits immediately |
| Ctrl+C in the listener window | Exits |

The idle timeout is measured from the **last request**, not from startup. A fixed 30-second
wall-clock window would kill the server while someone was still typing. While the page is open
it pings every 10 seconds, so the listener lives exactly as long as it is being used.

The 5-second grace period on tab close exists because `pagehide` also fires on F5. Exiting
immediately would kill the listener out from under a plain refresh; the reloaded page cancels
the countdown with its first request.

## Endpoints

| Method | Path | Purpose |
|---|---|---|
| GET | `/` | The data entry UI |
| GET | `/api/health` | Liveness plus request count |
| POST | `/api/ping` | Keep-alive; does not count toward `-MaxRequests` |
| POST | `/api/release` | Page is closing; start the grace countdown |
| POST | `/api/shutdown` | Exit now |
| POST | `/api/data/fg-dimensions` | Append an FG row |
| POST | `/api/data/carton-dimensions` | Append a carton row |
| POST | `/api/data/bom-assignments` | Append a BOM assignment |
| GET | `/api/data/<any of the three>` | List rows as JSON |

A POST body may be a single JSON object or an **array** for batch entry.

### Status codes

| Code | Meaning |
|---|---|
| 200 | Inserted. Batch responses carry `inserted` / `failed` counts and a per-row `results[]` |
| 400 | Validation failed. `details[]` gives `field` and `message` per problem |
| 409 | Duplicate key |
| 503 | CSV locked by another program - close Excel and retry |

A batch response is `success: true` whenever the batch was *processed*. Partial success is
reported per row in `results[]`, not by failing the whole response.

## Data

Written to `system/generated/`, RFC-4180 quoted, matching the existing WVN generated CSVs:

```
fg_dimensions.csv           FG_ID,Length_mm,Width_mm,Height_mm,Target_Qty,Source,Timestamp
carton_dimensions.csv       Carton_ID,Length_mm,Width_mm,Height_mm,Dimension_Source,Timestamp
bom_carton_assignments.csv  FG_ID,Carton_ID,Source_BOM_Date,Confidence,Timestamp
```

Validation, referential-integrity rules and sample data: [`docs/data-contract.md`](docs/data-contract.md).

Transaction log: `system/logs/listener.log`.

> **This folder has its own `system/generated/`, deliberately.** The parent Fit Tool's
> `system/generated/fg_dimensions.csv` uses a completely different schema
> (`fg_id,fg_name,length_mm,...,conflict_detail`). Pointing this listener at that folder would
> append incompatible rows to a file the main tool parses. Keep them separate, or pass an
> explicit `-DataPath`.

## Tests

```bash
powershell -NoProfile -ExecutionPolicy Bypass -File tests\test.ps1
```

19 cases covering the happy path, every validation and conflict rule, malformed JSON, batch
insert, referential integrity, latency, graceful shutdown, and an explicit no-file-locking check
that takes an **exclusive** lock on a CSV while the listener is running. Runs against a throwaway
`tests/_run/` directory and exits non-zero on any failure. Results land in
`tests/_run/test-results.json`.

## Layout

```
http-listener/
├─ launcher.cmd                 Double-click entry point
├─ listen.ps1                   The listener
├─ index.html                   The UI (single file, no dependencies)
├─ README.md
├─ PROGRESS.md                  How this was built, and what QC caught
├─ docs/
│  └─ data-contract.md          Schemas, validation rules, test data
├─ system/
│  ├─ generated/                CSVs, created on first write
│  └─ logs/listener.log
├─ tests/
│  └─ test.ps1
└─ tools/                       Multi-agent build harness (not needed at runtime)
   ├─ dispatch-worker.ps1
   ├─ prompts/
   └─ worker-output/
```

## Troubleshooting

**"Port 3000 is already in use."** A listener is still running. Close its window, or use
`launcher.cmd 3100`.

**"CSV file locked. Close Excel and retry." (503)** Excel takes an exclusive lock on an open CSV.
The listener retries four times before giving up; close the file and resubmit.

**Nothing happens on double-click.** Group policy may block `.ps1` execution. `launcher.cmd`
already passes `-ExecutionPolicy Bypass`; if it still fails, check the listener window and
`system/logs/listener.log`.

**Browser says the server is unreachable.** The listener has exited (idle timeout). Run
`launcher.cmd` again. Anything you submitted while it was down is queued in the browser and
retried automatically on the next load.
