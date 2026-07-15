# WVN Master Carton Fit Tool — Maintainer Notes

Single-file offline browser tool for WAHL Vietnam. Manual NPI carton check with new shared-carton generation, existing-FG fit evaluation against the loaded master-carton list, and FG/MasterCarton data-link cross-check. HTML + embedded JS (XLSX parser, packing engine, SVG 3D visualizer, EN/VI localization). No Node, Python, server, CDN, or external dependency.

## Files

| File | Role |
|------|------|
| `WVN_Master_Carton_Fit_Tool.html` | Complete application (HTML/CSS/JS) |
| `Open_WVN_Master_Carton_Fit_Tool.cmd` | Launcher — generates manifests, opens browser |
| `WVN Master Carton Workbook Manifest.json` | Generated at launch — points to dimension workbook |
| `Weekly BOM Manifest.json` | Generated at launch — points to Weekly BOM files |
| `WVN_Master_Carton_Fit_Tool_launcher.log` | Generated at launch — detailed log of all operations |
| `carton_tool_input.xlsx` | **Legacy fallback only** — old combined three-sheet workbook |

## Launch sequence

Run `Open_WVN_Master_Carton_Fit_Tool.cmd`. What happens:

1. **Detects username** — parses from `%~dp0` path (must start with `C:\Users\<username>\`).
2. **Dynamic OneDrive detection** — scans `C:\Users\<username>\OneDrive*` folders, picks the one containing `NPI Information - Documents`. Falls back to the first OneDrive folder.
3. **Builds company paths** (username and OneDrive are the only variables):
   ```
   NPI_TOOLS_DIR = <OneDrive>\NPI Information - Documents\General\Project Management\5. Tools for NPI
   WEEKLY_BOM_DIR = <NPI_TOOLS_DIR>\1. Weekly BOM
   ```
4. **Searches for dimension workbook** (priority): `WEEKLY_BOM_DIR` → `TOOL_DIR\1. Weekly BOM` → `NPI_TOOLS_DIR` → `TOOL_DIR`. Prefers `Products Physical Dimension - FG 1.Apr.25.xlsx`. `carton_tool_input.xlsx` is legacy fallback only.
5. **Searches for Weekly BOM** — scans `WEEKLY_BOM_DIR` and local paths for `BOM & FORMULA * D365.xlsx`. Sorted by **parsed filename date** (MONTH DD pattern, year fixed to 2026) descending. Falls back to `LastWriteTimeUtc` if filename date is unparseable. Fallback order: JUL 10 → JUL 03 → JUN 12. Dimension workbooks (containing "Physical Dimension") are excluded from BOM candidates.
6. **Writes two manifest files** beside the launcher. Each entry has `name`, `path` (Windows), `url` (file:/// URI), `modifiedUtc`, `sourceType`.
7. **Writes a detailed log** to `WVN_Master_Carton_Fit_Tool_launcher.log` with timestamps, paths searched, files found, and any errors.
8. **Launches Chrome or Edge** with `--allow-file-access-from-files` and an isolated temp profile.

For launcher validation without opening a browser, run `Open_WVN_Master_Carton_Fit_Tool.cmd --manifests-only`.

## Dimensional tolerance (10%)

A 10% proportional tolerance is applied on every axis for the approval scenario:

| Scenario | FG multiplier | Carton multiplier | Purpose |
|----------|--------------|-------------------|---------|
| Worst-case (approval) | ×1.10 | ×0.90 | Only this counts as "suitable" |
| Nominal | ×1.00 | ×1.00 | Diagnostic comparison |
| Best-case | ×0.90 | ×1.10 | Diagnostic only — never shown as approved |

This answers: "Will the target still fit when the FG is 10% larger and the carton is 10% smaller?"

### Fit classification

`ToleranceClassifier.classify()` evaluates all three scenarios and returns one of:

| Status | Meaning | Display |
|--------|---------|---------|
| `ROBUST_PASS` | Worst-case fits | Green — suitable carton |
| `NOMINAL_ONLY` | Worst-case fails, nominal passes | Orange — requires physical validation |
| `FAVORABLE_ONLY` | Only best-case fits | Red — high risk, never recommended |
| `FAIL` | Even best-case fails | Red |

### How tolerance is applied

`CartonFitEngine.evaluate()` accepts `options.scenario` (default `'worst'`). Before computing the grid or mixed-orientation packing, it creates toleranced copies of the FG and carton dimensions. All volume calculations and the `OrthogonalPackingSolver.find()` call use the toleranced copies. The cache key includes the scenario, so all three scenarios share a single bounded Map (~5000 entries).

`CartonFitEngine.evaluateAllScenarios(fg, carton, options)` returns `{ WORST, NOMINAL, BEST }` — three full evaluate results from one call.

### Result table behavior

- **Main table** shows only `ROBUST_PASS` cartons
- **Assigned BOM carton** shown only if `ROBUST_PASS`
- **Nominal-only section** appears below with "requires physical validation" header when `NOMINAL_ONLY` cartons exist
- **Tolerance status column** added to all result tables with colored badges
- **FG info display** shows both nominal and worst-case dimensions

### New carton proposal sizing

Proposals use an iterative worst-case formula:
```
requiredWorstCase = count × (nominalFgAxis × 1.10)
proposed = roundUpTo5(requiredWorstCase + clearance)
proposed = roundUpTo5(requiredWorstCase + clearance + 0.10 × proposed)
```
This ensures the nominal carton is large enough that after 10% shrinkage it still fits the worst-case FG. Every proposal is validated via worst-case evaluate — non-ROBUST_PASS proposals are rejected.

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

If the BOM manifest is unavailable or stale, the HTML also tries known BOM filenames in `./1. Weekly BOM` and the legacy sibling `../1. Weekly BOM` location. Candidates are sorted by **parsed filename date** (via `parseBomFilenameDate()`) descending first, then by `modifiedMs` descending.

### BOM detection error

If a user accidentally loads a BOM file as the Dimension Workbook (via Select Dimension Workbook or drag-and-drop), the tool detects the BOM filename pattern and shows a specific error message (EN/VI) explaining the two-workbook workflow. The error is produced when `parseBomFilenameDate()` matches the file name but the file lacks the required FG dimension or Master Carton dimension sheets.

## Data architecture (split-source)

Two independent workbooks replace the old single `carton_tool_input.xlsx`.

### Dimension workbook — FG EA dimensions + carton dimensions

Key sheets (detected by header row, not sheet name):
- **Dimension** / **Dimension_FG_packaging** — `Item number`, `Unit`, `Width`, `Depth`, `Height`. Only `Unit = EA`. cm → ×10 → mm. Duplicate EA rows with different dims → flagged as conflict, FG blocked.
- **Mastercarton_Dimensions** — `carton_id`, `carton_name`, `carton_width_cm`, `carton_depth_cm`, `carton_height_cm`, `carton_lifecycle`. `fg_per_carton` is **optional** — cartons load without it.
- **FG_MasterCarton** (optional) — provisional assignment only. Supports column aliases: `FinishGood_ID`, `FinishGood_Name`, `QtyPerMasterCarton`. Header may be on row 1+ (row 0 can be a note).

### Weekly BOM — authoritative FG→Carton assignment + target quantity

**Sheet selection**: Only Sheet3 and Sheet1 (2) are parsed (by index, 1 and 2). Sheet1 (index 0) — the raw D365 print layout — is skipped.

Parsers in `WeeklyBomLoader`:
- `parseNormalized()` — header-driven (FG / Item number / Product name), scans first 60 rows
- `parseFormat2()` — FG / Item number / Product name / Quantity / QTY / Per series
- `parseD365()` — state-machine for D365 export blocks

**All IDs (FG, Item number) are treated as text.** Values like `03259` are never converted to numbers.

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
- `CartonFitEngine`, `NewCartonGenerator`, `CartonScorer`, `PackingVisualizer3D`, `FitInterpretation`, `ToleranceClassifier`
- **Never add unit conversion inside these classes.**

## JavaScript class map (all in the HTML file)

| Class | Role |
|-------|------|
| `I18n` | EN/VI string dictionaries |
| `WorkbookLoader` | XLSX ZIP reading |
| `SimpleZip` | Browser-native ZIP deflater |
| `WeeklyBomLoader` | BOM parsing (3 format strategies, Sheet2/Sheet3 only) |
| `DataNormalizer` | Schema detection, header scanning, unit conversion |
| `BomTruthReconciler` | FG-to-carton assignment, target derivation |
| `OrthogonalPackingSolver` | Bounded mixed-orientation backtracking |
| `CartonFitEngine` | Multi-scenario tolerance + caching + six-orientation grid + mixed search |
| `ToleranceClassifier` | ROBUST_PASS / NOMINAL_ONLY / FAVORABLE_ONLY / FAIL from 3 scenarios |
| `FitInterpretation` | Packing-method status codes (separate axis from tolerance) |
| `NewCartonGenerator` | Worst-case proposal generation with iterative tolerance compensation |
| `CartonScorer` | Existing-carton ranking (filters by worst-case ROBUST_PASS) |
| `SharedCartonLineGenerator` | Manual NPI shared-carton line suggestions |
| `PackingVisualizer3D` | SVG 3D renderer |
| `UiController` | All UI logic, event handling, manifest loading |

Key helper functions: `resolveWorkbookSource()`, `parseBomFilenameDate()`, `fmt()`, `pct()`, `dims()`, `escapeHtml()`, `idKey()`.

## What must NOT change

- Six-orientation uniform grid test in `CartonFitEngine`
- Physical-capacity gating (volume alone must never pass a carton)
- `NewCartonGenerator` clearance (5 mm) and increment (5 mm)
- `OrthogonalPackingSolver` maxTarget (12), maxNodes (150,000)
- Downstream mm-only contract
- Manual NPI form logic and shared-carton generation
- `FitInterpretation` status codes
- Source workbook link infrastructure (manifest-driven, `file:///` URLs must NOT be filtered)

## Key behaviors for a new maintainer to verify

1. **Launch**: Run `.cmd` → `.log` file written → both manifests generated with OneDrive paths → HTML auto-loads dimension and BOM → no fallback warning → source workbook links appear
2. **Dynamic OneDrive**: Works on any machine with any OneDrive folder name (OneDrive, OneDrive (1), OneDrive - Wahl Clipper Corporation, etc.) as long as it contains `NPI Information - Documents`
3. **BOM selection**: Most recent BOM by parsed filename date (e.g., JUL 10 > JUL 03 > JUN 12). Run `--manifests-only` and check the log to verify
4. **Tolerance**: FG info shows nominal + worst-case dimensions. Result table shows ROBUST_PASS cartons with green tolerance badges. Nominal-only cartons appear in a separate section
5. **BOM detection error**: Load a BOM file via "Select Dimension Workbook" → specific error message appears (EN/VI)
6. **Data completeness**: FG dropdown shows ✓ Data complete / ⚠ Missing with specific missing fields
7. **Weekly BOM is sole source of truth**: Assignment, carton ID, and target quantity come exclusively from the Weekly BOM
8. **Fallback mode**: When BOM fails, amber banner + provisional FG_MasterCarton data used
9. **JS syntax**: Extract `<script>` content, `node --check` — must pass
10. **Browser console**: Zero uncaught errors after startup and after selecting an FG
11. **EN/VI**: Toggle after results displayed, confirm all labels translate including new tolerance labels
12. **Dark mode**: Toggle, verify tables, inputs, cards, SVG wireframe all render
13. **3D rotation**: Drag and buttons must not mutate FG/carton/fit dimensions
14. **Proposals**: New carton proposals sized for worst-case, validate as ROBUST_PASS
15. **`--manifests-only`**: Generates manifests and log only, skips browser

## Known constraints

- Requires `.cmd` launcher with `--allow-file-access-from-files` for auto-loading
- `.xlsx` / `.xlsm` only, not legacy `.xls`
- Formula cells use cached values
- Proposal generation is synchronous; large FG pools (~1000) take seconds
- No diagonal/nesting/compression/non-box modelling — physical sample validation always required
- No result is production-approved — tool is advisory only
