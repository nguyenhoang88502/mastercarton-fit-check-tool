# Data entry helper

This is the small background program that lets the Master Carton Fit Tool **save** new products
and cartons. It has no window and no screen of its own — everything you see is inside the tool
itself, on the **New Product** tab and the **Add Data** tab.

It starts automatically when you open the tool. Nothing to install.

## What it is for

The tool reads its product and carton list from the weekly workbooks. Until now there was no way
to add something that was not in those workbooks yet — a brand new product still being measured,
for example. This helper stores those entries on the PC and hands them back to the tool the next
time it opens, so a new product only has to be typed once.

Entries are kept **outside** the `generated` folder on purpose. The tool rebuilds that folder
from the workbooks every time it starts, so anything saved there would be wiped.

## How people use it

**One product** — fill in the New Product form as normal, then press **Save this product**.

**Many at once** — open the **Add Data** tab, download the blank sheet, fill it in, drag it back
onto the page, check the preview, then press **Save all**.

If an item is already on the list, the new numbers replace the old ones. Re-measuring is a
correction, so the newest entry always wins.

**To stop it** — press **Turn off** in the top-right corner of the tool. That ends the background
program. Closing and reopening the tool starts it again.

## Requirements

Stock Windows 10/11, nothing installed:

- Windows PowerShell 5.1 (built in)
- `System.Net.HttpListener` (built into .NET Framework)

No Python, Node.js, .NET Core, IIS, SQLite, pip or npm.

---

# Technical notes

## Why it exits on its own

The helper appends to CSV and then goes away, so no handle is left open and Excel or the tool can
open the same files at any moment. There is no "file in use" state to get stuck in.

It exits when any of these happen:

| Trigger | Behaviour |
|---|---|
| No traffic for `-IdleTimeoutSeconds` (launcher passes 900) | Exits |
| `-MaxRequests` data requests handled | Exits |
| Tool page closed | `/api/release`, then exits after a 5s grace period |
| `POST /api/shutdown` (the **Turn off** button) | Exits immediately |
| Ctrl+C | Exits |

The idle timeout runs from the **last request**, not from startup — a fixed window would kill it
while someone was still typing. The tool pings every 10 seconds while open, so it lives exactly
as long as it is being used.

The 5-second grace period exists because the browser fires the same "page closing" event on F5.
Exiting immediately would kill the helper out from under a plain refresh.

## Run it by hand

```bash
powershell -NoProfile -ExecutionPolicy Bypass -File listen.ps1 -Port 3000 -IdleTimeoutSeconds 900
```

## Endpoints

| Method | Path | Purpose |
|---|---|---|
| GET | `/api/health` | Liveness plus request count |
| POST | `/api/ping` | Keep-alive; does not count toward `-MaxRequests` |
| POST | `/api/release` | Page is closing; start the grace countdown |
| POST | `/api/shutdown` | Exit now |
| POST | `/api/data/fg-dimensions` | Save a product |
| POST | `/api/data/carton-dimensions` | Save a carton |
| POST | `/api/data/bom-assignments` | Save a product-to-carton assignment |
| GET | `/api/data/<any of the three>` | List saved rows as JSON |

A POST body may be a single JSON object or an **array** for batch entry.

| Code | Meaning |
|---|---|
| 200 | Saved. `replaced: true` means it overwrote an existing row |
| 400 | Validation failed; `details[]` gives `field` and `message` |
| 503 | CSV locked by another program — close Excel and retry |

There is deliberately **no 409**. Re-entering an existing key replaces it.

Batch responses are `success: true` whenever the batch was *processed*; per-row outcomes are in
`results[]`, and partial success is not a failure.

## Data

Written to `system/generated/` inside this folder, RFC-4180 quoted:

```
fg_dimensions.csv           FG_ID,Length_mm,Width_mm,Height_mm,Target_Qty,Source,Timestamp
carton_dimensions.csv       Carton_ID,Length_mm,Width_mm,Height_mm,Dimension_Source,Timestamp
bom_carton_assignments.csv  FG_ID,Carton_ID,Source_BOM_Date,Confidence,Timestamp
```

Validation rules and sample data: [`docs/data-contract.md`](docs/data-contract.md).
Log: `system/logs/listener.log`.

The tool maps these onto its own raw rows at load time, before any of its calculations run, so no
analysis code had to change.

## Tests

```bash
powershell -NoProfile -ExecutionPolicy Bypass -File tests\test.ps1
```

19 cases: happy path, every validation rule, overwrite-on-duplicate, malformed JSON, batch,
referential integrity, latency, graceful shutdown, and an explicit no-file-locking check that
takes an **exclusive** lock on a CSV while the helper is running. Uses a throwaway `tests/_run/`
directory and exits non-zero on any failure.

## Layout

```
system/http-listener/
├─ listen.ps1                   The helper
├─ README.md
├─ PROGRESS.md                  How this was built, and what QC caught
├─ docs/data-contract.md        Schemas, validation rules, test data
├─ system/generated/            Saved entries
├─ system/logs/listener.log
├─ tests/test.ps1
└─ tools/                       Multi-agent build harness (not needed at runtime)
```

It sits under `system/` so that the folder people actually open shows only the launcher and the
two text files. **Only `listen.ps1` is shipped** — this README, the data contract, the tests and
the build harness stay in the repository.

## If something goes wrong

**The Save controls are missing.** The helper is not running. Close and reopen the tool. Nothing
is lost — the tool works normally without it.

**"Could not save… CSV file locked."** Excel has one of the files open. Close it and press Save
again.

**Port 3000 is in use.** Another copy is already running; the launcher detects this and reuses it.
