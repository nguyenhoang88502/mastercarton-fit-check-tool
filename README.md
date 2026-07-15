# WVN Master Carton Fit Tool — Maintainer Notes

Single-file offline browser tool for WAHL Vietnam. Manual NPI carton check with new shared-carton generation, existing-FG fit evaluation against the loaded master-carton list, and FG/MasterCarton data-link cross-check. HTML + embedded JS (XLSX parser, packing engine, SVG 3D visualizer, EN/VI localization). No Node, Python, server, CDN, or external dependency.

## Files

| File | Role |
|------|------|
| `WVN_Master_Carton_Fit_Tool.html` | Complete application (HTML/CSS/JS) |
| `Open_WVN_Master_Carton_Fit_Tool.cmd` | Launcher — generates manifests, opens browser |
| `WVN Master Carton Workbook Manifest.json` | Generated at launch — points to dimension workbook |
| `Weekly BOM Manifest.json` | Generated at launch — points to Weekly BOM files |
| `carton_tool_input.xlsx` | **Legacy fallback only** — old combined three-sheet workbook |

## Launch sequence

Run `Open_WVN_Master_Carton_Fit_Tool.cmd`. What happens:

1. **Detects USER_ROOT** — parses username from `TOOL_DIR` when it starts with `C:\Users\`; otherwise falls back to `%USERPROFILE%`
2. **Builds fixed company paths** (username is the only variable):
   ```
   NPI_TOOLS_DIR = %USER_ROOT%\OneDrive - Wahl Clipper Corporation\NPI Information - Documents\General\Project Management\5. Tools for NPI
   WEEKLY_BOM_DIR = %NPI_TOOLS_DIR%\1. Weekly BOM
   MASTER_CARTON_TOOL_DIR = %NPI_TOOLS_DIR%\MasterCarton fit check tool
   INSPECTION_TOOL_DIR = %NPI_TOOLS_DIR%\Inspection Standard Tool
   ```
3. **Searches for dimension workbook** (priority): `WEEKLY_BOM_DIR` → `TOOL_DIR\1. Weekly BOM` → `MASTER_CARTON_TOOL_DIR` → `TOOL_DIR`. Prefers `Products Physical Dimension - FG 1.Apr.25.xlsx`. `carton_tool_input.xlsx` is legacy fallback only when no real dimension workbook exists.
4. **Searches for Weekly BOM** (priority): `WEEKLY_BOM_DIR` → `TOOL_DIR\1. Weekly BOM` → `TOOL_DIR`. Does not search Downloads or Desktop. Matches `BOM & FORMULA * D365.xlsx`, sorted by `LastWriteTimeUtc` descending (most recently modified file wins). Filename dates are not parsed or used for ordering.
5. **Writes two manifest files** beside the launcher. Each entry has `name`, `path` (Windows), `url` (file:/// URI), `modifiedUtc`, `sourceType`.
6. **Launches Chrome or Edge** with `--allow-file-access-from-files` and an isolated temp profile.

For launcher validation without opening a browser, run `Open_WVN_Master_Carton_Fit_Tool.cmd --manifests-only`.

## How the HTML loads workbooks

On startup the HTML fetches both manifest JSON files and auto-loads the first entry of each.

**`resolveWorkbookSource(entry)`** (global function) — accepts a manifest entry and returns `{value, url, type, name}`. Handles four source types:
- `fileUrl` — `file:///C:/...` passes through unchanged
- `absolutePath` — `C:\...` converted to `file:///C:/...` with percent-encoding
- `relativePath` — `1. Weekly BOM/file.xlsx` resolved against `location.href`
- `siblingName` — `carton_tool_input.xlsx` resolved as `./file.xlsx`

Prefers `entry.url` → `entry.href` → `entry.path` → `entry.name` → `entry.file`.

**`loadWorkbookFromManifestEntry(entry, label)`** — calls `resolveWorkbookSource()`, fetches via the resolved URL, returns `{file, sourceHref}` which feeds into the existing `load()` pipeline.

**`loadSibling(filename)`** — retained for backward compat. Only accepts simple filenames (no slashes).

**Manual file picker and drag/drop** — unchanged, bypass manifest entirely.

If the BOM manifest is unavailable or stale, the HTML also tries known BOM filenames in `./1. Weekly BOM` and the legacy sibling `../1. Weekly BOM` location. Candidates are sorted by `modifiedMs` descending only — no filename-date priority.

## Data architecture (split-source)

Two independent workbooks replace the old single `carton_tool_input.xlsx`.

### Dimension workbook — FG EA dimensions + carton dimensions

Key sheets (detected by header row, not sheet name):
- **Dimension** / **Dimension_FG_packaging** — `Item number`, `Unit`, `Width`, `Depth`, `Height`. Only `Unit = EA`. cm → ×10 → mm. Duplicate EA rows with different dims → flagged as conflict, FG blocked.
- **Mastercarton_Dimensions** — `carton_id`, `carton_name`, `carton_width_cm`, `carton_depth_cm`, `carton_height_cm`, `carton_lifecycle`. `fg_per_carton` is **optional** — cartons load without it.
- **FG_MasterCarton** (optional) — provisional assignment only. Supports column aliases: `FinishGood_ID`, `FinishGood_Name`, `QtyPerMasterCarton`. Header may be on row 1+ (row 0 can be a note).

### Weekly BOM — authoritative FG→Carton assignment + target quantity

Parsers in `WeeklyBomLoader`:
- `parseNormalized()` — header-driven (FG / Item number / Product name), scans first 60 rows
- `parseFormat2()` — FG / Item number / Product name / Quantity / QTY / Per series
- `parseD365()` — state-machine for D365 export blocks

BOM row shape: `{fg, fgName, itemNumber, productName, norm, quantity, perSeries, type, vendor, uom, buyerGroup, sourceSheet, sourceFormat, sourceName}`

### Target quantity derivation chain (in `BomTruthReconciler.deriveTargetFromBomRow()`)

1. BOM `perSeries` > 1 → `floor(perSeries)`, source = `weekly_bom_per_series`
2. BOM `norm` — if < 1 → `round(1/norm)`, else → `floor(norm)`, source = `weekly_bom_norm`
3. BOM `quantity` — if < 1 → `round(1/quantity)`, else → `floor(quantity)`, source = `weekly_bom_quantity`
4. If none yield a positive integer → `NaN`, source = `weekly_bom_no_target`

No fallback to `FG_MasterCarton QtyPerMasterCarton` or `carton.defaultQty`. Weekly BOM is the sole source of truth for target quantity during automatic reconciliation. Manual user override in the UI still works but is not part of automatic reconciliation.

### Master Carton candidate identification

`BomTruthReconciler.cartonCandidates()` — matches BOM rows to carton IDs (exact) or by product name regex. Excludes labels: `/carton\s*label|lith\s*label|fire\s*label|lithium\s*label|warning\s*label|rating\s*label|label\s*for/i`. Candidates are returned in **original Weekly BOM sheet order** — the first matching row wins as the authoritative assignment. No sorting by itemNumber, productName, date, quantity, or confidence.

## Internal unit contract (CRITICAL)

After `DataNormalizer`, **every dimension is in millimetres**. All downstream code receives normalized mm only:
- `CartonFitEngine`, `NewCartonGenerator`, `CartonScorer`, `PackingVisualizer3D`, `FitInterpretation`
- **Never add unit conversion inside these classes.**

## JavaScript class map (all in the HTML file)

| Class | Lines (approx) | Role |
|-------|---------------|------|
| `I18n` | ~85 | EN/VI string dictionaries |
| `WorkbookLoader` | ~125 | XLSX ZIP reading |
| `SimpleZip` | ~165 | Browser-native ZIP deflater |
| `WeeklyBomLoader` | ~174 | BOM parsing (3 format strategies) |
| `DataNormalizer` | ~184 | Schema detection, header scanning, unit conversion |
| `BomTruthReconciler` | ~234 | FG-to-carton assignment, target derivation |
| `OrthogonalPackingSolver` | ~256 | Bounded mixed-orientation backtracking |
| `CartonFitEngine` | ~267 | Six-orientation uniform grid |
| `NewCartonGenerator` | ~293 | Proposal generation (controlled patterns) |
| `CartonScorer` | ~314 | Existing-carton ranking |
| `SharedCartonLineGenerator` | ~335 | Manual NPI shared-carton line suggestions |
| `PackingVisualizer3D` | ~349 | SVG 3D renderer |
| `UiController` | ~388 | All UI logic, event handling, manifest loading |

Key helper functions: `resolveWorkbookSource()` (~388), `fmt()`, `pct()`, `dims()`, `escapeHtml()`, `idKey()`.

## What must NOT change

- Six-orientation uniform grid test in `CartonFitEngine`
- Physical-capacity gating (volume alone must never pass a carton)
- `NewCartonGenerator` patterns and clearance (5 mm)
- `OrthogonalPackingSolver` maxTarget (12), maxNodes (150,000)
- Downstream mm-only contract
- Manual NPI form logic and shared-carton generation
- `FitInterpretation` status codes
- Source workbook link infrastructure (manifest-driven, `file:///` URLs must NOT be filtered)

## Key behaviors for a new maintainer to verify

1. **Launch**: Run `.cmd` → both manifests generated → HTML auto-loads dimension and BOM → no fallback warning → source workbook links appear in "Source workbooks:" row
2. **Paths**: `TOOL_DIR` username parsing drives `USER_ROOT` and company OneDrive paths. Copy the folder to another machine under `C:\Users\<name>\...`, run the `.cmd`, and manifests resolve from that user's OneDrive Weekly BOM folder when present.
3. **Data completeness**: FG dropdown shows ✓ Data complete / ⚠ Missing with specific missing fields
4. **Weekly BOM is sole source of truth**: Assignment, carton ID, and target quantity come exclusively from the Weekly BOM. FG_MasterCarton and carton.defaultQty are never used as fallback for assignment or target during automatic reconciliation.
5. **Fallback mode**: When BOM fails, amber banner + provisional FG_MasterCarton data used
6. **Source workbook links**: Clickable `file:///` links for dimension workbook and Weekly BOM appear after manifest loading. Manual file selection shows a note that the browser hides the original path.
7. **JS syntax**: Extract `<script>` content, `node --check` — must pass
8. **Browser console**: Zero uncaught errors after startup and after selecting an FG
9. **EN/VI**: Toggle after results displayed, confirm all labels translate
10. **Dark mode**: Toggle, verify tables, inputs, cards, SVG wireframe all render
11. **3D rotation**: Drag and buttons must not mutate FG/carton/fit dimensions

## Historical BOM Existing-Carton Audit (removed)

The Historical BOM Existing-Carton Audit section was removed (commit `9a79b12`). It was a standalone audit module that evaluated one FG/current-carton pair against all existing cartons. It was removed because the implementation was tangled and unreliable. If re-added, it must be a self-contained module that never generates new carton dimensions, never calls `NewCartonGenerator` or `SharedCartonLineGenerator`, and never shows temporary/SHARED-LINE IDs.

## Known constraints

- Requires `.cmd` launcher with `--allow-file-access-from-files` for auto-loading
- `.xlsx` / `.xlsm` only, not legacy `.xls`
- Formula cells use cached values
- Proposal generation is synchronous; large FG pools (~1000) take seconds
- No diagonal/nesting/compression/non-box modelling — physical sample validation always required
- No result is production-approved — tool is advisory only
