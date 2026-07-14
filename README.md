# WVN Master Carton Fit Tool — Maintainer Notes

Single-file offline browser tool. Checks whether an FG package fits into an existing master carton, proposes new carton dimensions when none fits. HTML + embedded JS (XLSX parser, packing engine, SVG 3D visualizer, EN/VI localization). No Node, Python, server, CDN, or external dependency.

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

1. **Detects USER_ROOT** from `%USERPROFILE%`
2. **Builds fixed company paths** (username is the only variable):
   ```
   NPI_TOOLS_DIR = %USER_ROOT%\OneDrive - Wahl Clipper Corporation (1)\NPI Information - Documents\General\Project Management\5. Tools for NPI
   WEEKLY_BOM_DIR = %NPI_TOOLS_DIR%\1. Weekly BOM
   MASTER_CARTON_TOOL_DIR = %NPI_TOOLS_DIR%\MasterCarton fit check tool
   ```
3. **Searches for dimension workbook** (priority): WEEKLY_BOM_DIR → MASTER_CARTON_TOOL_DIR → TOOL_DIR. Prefers `Products Physical Dimension - FG 1.Apr.25.xlsx`.
4. **Searches for Weekly BOM** (priority): WEEKLY_BOM_DIR → TOOL_DIR\1. Weekly BOM → TOOL_DIR. Matches `BOM & FORMULA * D365.xlsx`, sorted by filename month/day (JUL 03 > JUN 12) then by LastWriteTimeUtc.
5. **Writes two manifest files** beside the launcher. Each entry has `name`, `path` (Windows), `url` (file:/// URI), `modifiedUtc`, `sourceType`.
6. **Launches Chrome or Edge** with `--allow-file-access-from-files` and an isolated temp profile.

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

### Target quantity derivation chain (in `BomTruthReconciler.reconcile()`)

1. BOM `perSeries` > 1 → `floor(perSeries)`, source = `weekly_bom_per_series`
2. BOM `norm` / `quantity` — if < 1 → `round(1/value)` (e.g. 0.000625 → 1600), source = `weekly_bom_norm`
3. Provisional `QtyPerMasterCarton` from FG_MasterCarton
4. `carton.defaultQty` from Mastercarton_Dimensions (if column exists)
5. User manual override

### Master Carton candidate identification

`BomTruthReconciler.cartonCandidates()` — matches BOM rows to carton IDs (exact) or by product name regex. Excludes labels: `/carton\s*label|lith\s*label|fire\s*label|lithium\s*label|warning\s*label|rating\s*label|label\s*for/i`

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
| `PackingVisualizer3D` | ~349 | SVG 3D renderer |
| `UiController` | ~388 | All UI logic, event handling, manifest loading |

Key helper functions: `resolveWorkbookSource()` (~388), `fmt()`, `pct()`, `dims()`, `escapeHtml()`, `idKey()`.

## What must NOT change

- Six-orientation uniform grid test in `CartonFitEngine`
- Physical-capacity gating (volume alone must never pass a carton)
- `NewCartonGenerator` patterns and clearance (5 mm)
- `OrthogonalPackingSolver` maxTarget (12), maxNodes (150,000)
- Downstream mm-only contract
- Manual NPI form and audit form logic
- `FitInterpretation` status codes

## Key behaviors for a new maintainer to verify

1. **Launch**: Run `.cmd` → both manifests generated → HTML auto-loads dimension and BOM → no fallback warning
2. **Paths**: %USERPROFILE% drives all paths. Copy the folder to another machine, run the `.cmd`, it works.
3. **Data completeness**: FG dropdown shows ✓ Data complete / ⚠ Missing with specific missing fields
4. **Weekly BOM is authoritative**: Assignment and target overrule FG_MasterCarton
5. **Fallback mode**: When BOM fails, amber banner + provisional FG_MasterCarton data used
6. **JS syntax**: Extract `<script>` content, `node --check` — must pass
7. **Browser console**: Zero uncaught errors after startup and after selecting an FG
8. **EN/VI**: Toggle after results displayed, confirm all labels translate
9. **Dark mode**: Toggle, verify tables, inputs, cards, SVG wireframe all render
10. **3D rotation**: Drag and buttons must not mutate FG/carton/fit dimensions

## Known constraints

- Requires `.cmd` launcher with `--allow-file-access-from-files` for auto-loading
- `.xlsx` / `.xlsm` only, not legacy `.xls`
- Formula cells use cached values
- Proposal generation is synchronous; large FG pools (~1000) take seconds
- No diagonal/nesting/compression/non-box modelling — physical sample validation always required
- No result is production-approved — tool is advisory only
