# WAHL Vietnam Master Carton Fit Tool — Technical Handover

This document describes the current architecture, protected behavior, known failure modes, validation requirements, and safe change process.

Handover date: **2026-07-30**

## 1. Current repository inventory

Current maintained files:

```text
.gitignore
README.md
HANDOVER.md
Quick Start.txt
Open_WVN_Master_Carton_Fit_Tool.cmd
system/config.ini
system/launcher.ps1
system/WVN_Master_Carton_Fit_Tool.html
```

Generated files in `system/generated/`:

```text
fg_dimensions.csv
carton_dimensions.csv
bom_carton_assignments.csv
carton_fg_reverse_index.csv
carton_dimension_sources.csv
data_quality.csv
master_carton_picture_index.csv
cache_state.ini
```

Runtime logs are in `system/logs/`. `.wrangler/` is local tooling state. Source workbooks and carton pictures are in the sibling `1. Weekly BOM` folder, not the application root.

## 2. Architecture overview

```text
CMD
→ PowerShell launcher
→ workbook discovery and preprocessing
→ generated CSV layer
→ browser HTML application
→ JavaScript calculations and UI
```

PowerShell prepares trusted input data. JavaScript performs fit calculations and user interaction.

## 3. Runtime constraints

Windows, offline operation, standard user rights, Chrome or Edge, and Windows PowerShell are supported. Runtime operation requires no Node.js, Python, local server, CDN, Excel COM automation, administrator rights, or internet connection.

## 4. Folder and path rules

Standard names are `WVN_Master_Carton_Fit_Tool`, `1. Weekly BOM`, `master-carton-pictures`, `system`, `generated`, and `logs`. Launcher paths derive only from `$PSScriptRoot`.

Never commit an absolute user-profile path, tenant-specific OneDrive path, or company SharePoint path. The generated cache records resolved runtime source paths by design; they are runtime fingerprints, not hardcoded configuration.

Current deployment-name deviations: the application directory is not yet named `WVN_Master_Carton_Fit_Tool`, the picture folder is `Master_Carton_Pics`, and the quick-start file is `Quick Start.txt`.

## 5. PowerShell launcher architecture

### Configuration and paths

`Read-IniFile` reads non-path settings. Source folders are resolved as exact siblings from `$PSScriptRoot`; legacy `[Search]` path keys are ignored with a warning. `config.ini` currently contains file patterns and launcher settings only.

### Logging

Each run creates `WVN_Master_Carton_Fit_Tool_launcher_yyyyMMdd_HHmmss_fff_PID.log`. `Write-Log` always writes to the console, catches append failures, warns, and continues. A named mutex, `Local\WVN_Master_Carton_Fit_Tool_Launcher`, prevents concurrent preprocessing.

### Workbook discovery

Dimension candidates prefer the canonical filename, then matching valid alternatives. BOM candidates use `Get-BomBusinessDateFromFilename`, business-date sorting, modification-time fallback, and structural validation.

### Lock policy

`Test-FileReadAvailable` opens candidates with `FileAccess.Read` and `FileShare.Read`. `Get-UnlockedFileHash` repeats the strict availability check before hashing. Locked sources must not be hashed, validated, parsed, copied, or opened with shared-write access; they are skipped while discovery continues.

### Preprocessing

The launcher parses XLSX ZIP/XML directly. It builds EA FG dimensions, BOM carton assignments, CA measurements resolved through FG assignments, normalized axes, ×10 millimetre conversion, carton averages, reverse indexes, and data-quality records. Carton dimension source/version are `BOM_LINKED_CA_AVERAGE` and `BOM_LINKED_CA_AVERAGE_V1`.

### Update-data choice

Before any workbook is touched, `launcher.ps1` decides whether to run the discovery/hash/rebuild block at all (`$skipUpdate`, the block is wrapped in `if (-not $skipUpdate) { ... }`). Order: `ForceRebuild` / `--force-rebuild` and `--update` update without asking; missing generated files force an update; `--skip-update` or `[Launcher] UpdateData=skip` skip; `UpdateData=always` updates; otherwise `Show-UpdateDataDialog` (WinForms: Update / Skip / Cancel, 30 s countdown to Update that stops on any mouse or key activity) asks. `--manifests-only` never asks. `config.local.ini` overrides `config.ini` for these keys (`Get-LauncherSetting`).

### Cache

`CacheSchemaVersion` is `2.0`; preprocessing version is `BOM_LINKED_CA_AVERAGE_V4_WEIGHTS` (bump it whenever the generated columns change, and keep the literal in the update-data dialog block in step). Both workbook sections validate version, path, size, timestamp, and hash. Reuse also requires all six core CSVs. Locked-source fallback requires matching cached source metadata and stored fingerprints.

### Generated-file writing

`Publish-GeneratedDataset` stages and validates all CSVs, backs up existing destinations, replaces with retries, and rolls back on failure. `Move-StagedFileWithRetry` performs each same-volume replacement. `cache_state.ini` is published last. Picture-index publication is separately atomic and nonfatal.

### Browser launch

The launcher looks for Chrome and Edge in their standard Program Files locations, creates a temporary browser profile, and starts the HTML with `--allow-file-access-from-files`.

## 6. JavaScript architecture

Pallet code added after the classes listed below: `computePalletLayout` (stable interlocked auto stacking), `buildStack3dSvg` (shared 3D renderer), `PalletVisualizer` (used twice: the Pallet Visualizer tab and the compact `#fgPalletPanel` inside FG Lookup, fed by `UiController.renderFgLookupPallet`), and `ManualPalletVisualizer` plus the `manual*` helpers for the free drag-and-drop builder. Multi-select in the manual builder is `builder.selection` (indices into the active level): `manualResolveGroup` moves the selection as one rigid block, `manualRotateGroup` turns it 90 degrees about its centre (falling back to pinning each corner), and Select mode adds box-select, Shift/Ctrl+click, `R`, `Delete`, `Ctrl+A` and `Esc`.

Weights: the New Product form sends `gross_weight_kg`, `net_weight_kg`, `carton_gross_weight_kg`, `carton_net_weight_kg` to `/api/save-dimension`; `update_dimension_excel.ps1` writes them to the FG's `EA` row and its `CA` row on the `Dimension` sheet (`Gross weight` / `Net Weight`, first header occurrence). Blank means the cell is left alone. FG Lookup shows them from `fg_dimensions.csv`.

Active classes include `I18n`, `WorkbookLoader`, `SimpleZip`, `WeeklyBomLoader`, `DataNormalizer`, `BomTruthReconciler`, `OrthogonalPackingSolver`, `FitSettings`, `TaskProgressManager`, `DimensionSettings`, `AppDataStore`, `ToleranceClassifier`, `CartonFitEngine`, `FitInterpretation`, `NewCartonGenerator`, `ExistingCartonRanking`, `CartonScorer`, `RecommendationPolicy`, `SharedDecisionEngine`, `RecommendationEngine`, `MissingDimensionQuery`, `BatchAuditEngine`, `SharedCartonLineGenerator`, `CsvExporter`, `PackingVisualizer3D`, and `UiController`.

Classes prefixed `DisabledLegacy` or `DisabledFixed`/`DisabledContinuous` are retained legacy implementations and are not the active calculation engine.

## 7. Calculation contract

Quick Mode uses the active FG, fixed carton dimensions, six orthogonal orientations, and rigid rectangular-grid packing. Optimized Mode keeps the FG fixed and lets each carton axis grow by up to 3 % (`1.00–1.03` in 0.01 steps: four values on each of three axes, 64 combinations, six orientations, 384 checks). The range is one constant, `FIT_TOLERANCE = 0.03`; the legacy closest-fit search applies the same ±3 % to FG ratios (`0.97–1.03`, 0.01 steps).

Outside carton dimensions receive a 10 mm deduction per axis exactly once.

## Protected area: Carton Audit

Protected IDs include `auditFgId`, `auditMcId`, `auditMcNote`, `auditMcSelectLabel`, `auditMcSelect`, `auditMcSelectNote`, `auditForm`, and `auditClearButton`.

Protected functions include `bindAuditModule`, `clearAuditModule`, `autoFillAuditMasterCarton`, `setAuditCartonControlMode`, `hideAuditMcSelect`, `hideAuditMcInput`, `updateAuditMcNote`, `findWeeklyBomAssignedCartonIdForAudit`, `handleAuditSubmit`, `getFgForAudit`, `getCartonForAudit`, `getAuditTargetQty`, `renderAuditError`, and `renderAuditResults`.

Assignment modes:

- Single: manual input remains enabled and is autofilled; selector is disabled and hidden.
- None: manual input remains enabled for manager entry.
- Multiple/conflicting: manual input is disabled; the required selector becomes enabled and visible.

Previous regressions included a hidden enabled required selector blocking submit, a positive fallback assignment not being written to `auditMcId`, and deletion of selector elements. Unrelated patches must not touch Carton Audit JavaScript.

## 9. FG Lookup requirements

FG Lookup resolves the BOM carton, renders aligned fit-result columns with adjusted utilization third in Optimized Mode, and omits the former FG relationship column. Session carton overrides are stored per FG, do not mutate source records, and are isolated from Carton Audit and Batch Audit. Session remeasurement becomes the active FG. Results include fit evidence and 3D rendering.

## 10. Batch Audit requirements

Quick Batch must use `activeFg` for current-carton and alternative calculations and evidence. Optimized Batch screens candidates, keeps cartons fixed, evaluates FG ratios, and retains compact winning evidence; full 3D placement is deferred until requested.

Carton-line reduction is a result/feature, not a third calculation mode. Its manager-facing classification must stay simple. When a different carton is recommended, all capacity, utilization, arrangement, ratios, scenario dimensions, classification, and search evidence must come from that candidate.

## 11. Picture-index feature

Status: **Implemented.**

The preferred source is `1. Weekly BOM/master-carton-pictures/`; the current launcher temporarily accepts `Master_Carton_Pics` with a warning. The exact basename regex is:

```regex
^(?<cartonId>[^_]+)_image(?<sequence>\d+)(?:_(?<variant>\d+))?$
```

Sorting uses carton ID, image sequence, base before variant, variant number, then filename. `master_carton_picture_index.csv` contains carton ID, sequence, variant, filename, and file URI. FG Lookup and Carton Lookup render the matching gallery with `object-fit: contain`. Pictures are display-only.

## 12. Known failure modes

- **Folders:** wrong app name, duplicate nesting, missing `system`, missing sibling source, malformed config, committed absolute path.
- **Logs:** legacy fixed log locked, Notepad/monitor lock, read-only log directory, propagated logging exception.
- **Concurrency:** repeated double-click, parallel preprocessing, mixed outputs.
- **Workbooks:** Excel/Teams/sync/preview locks, locked newest or all BOMs, corrupt or zero-byte XLSX, invalid expected filename, removal during discovery.
- **Cache:** missing/unreadable/old cache, missing section or required CSV, stale path, timestamp, size, or hash.
- **Outputs:** CSV locked, disk full, read-only directory, partial generation, cache published early, mixed old/new set.
- **HTML:** missing DOM element, hidden required field, duplicate listener, encoding corruption, table misalignment, stale translation, deleted hidden dependency.
- **Calculation:** double wall deduction, varied carton ratios, Quick Batch using source instead of active FG, mixed recommendation evidence, capacity misclassified as nominal fit.
- **Pictures:** substring matching, lexical numeric sorting, deleted image, cropped image, or calculation coupling.

## 13. Current unresolved risks

### Nonstandard deployment names

- Status: Open.
- Evidence: current application, picture, and quick-start names differ from the standard.
- User impact: warnings and deployment confusion; legacy picture fallback is currently required.
- Recommended fix: rename during a controlled deployment, not from the launcher.
- Priority: Medium.

### Incomplete interactive regression coverage

- Status: Open.
- Evidence: local `file://` browser automation was blocked; Carton Audit and calculation cases used extracted-code runtime tests.
- User impact: visual layout and complete click-path regressions remain manually verified.
- Recommended fix: run the matrix on a Windows workstation in Chrome/Edge.
- Priority: Medium.

### Batch and consolidation runtime coverage

- Status: Open.
- Evidence: source inspection confirms active-FG calls and candidate evidence structures, but full Quick Batch and consolidation UI runs were not executed in the latest validation.
- User impact: later unrelated HTML changes could regress these flows without automated detection.
- Recommended fix: add a repeatable offline browser regression harness.
- Priority: High.

## 14. Regression test matrix

| Area | Test | Expected | Last evidence |
| --- | --- | --- | --- |
| Startup | Normal CMD launch | Prepared data, exit 0 | Runtime passed |
| Startup | Double-click twice | Second exits cleanly | Runtime passed |
| Logging | Legacy log locked | Preparation continues | Runtime passed |
| Dimension | Source locked | Skipped; valid cache reused | Runtime passed |
| BOM | Latest locked | Next valid unlocked selected | Runtime passed |
| BOM | All locked | Reuse valid cache or clear stop | Runtime passed |
| Cache | Complete/matching | Reuse | Runtime passed |
| Cache | Missing assignment CSV | Rebuild | Static only |
| Carton Audit | Single | Autofill | Extracted runtime passed |
| Carton Audit | None | Manual entry | Extracted runtime passed |
| Carton Audit | Multiple/conflict | Selector | Extracted runtime passed |
| Quick Batch | Remeasured FG | Uses active dimensions | Static only |
| Consolidation | Different recommendation | Candidate evidence only | Static only |
| FG Lookup | Override cleared | Return to BOM carton | Static only |
| Picture | Exact/no match | Correct gallery/empty state | Static only |
| Encoding | Mojibake scan | No corruption | Static passed |
| Regression FG | 3031116 | Reference result | Extracted runtime passed |
| Generated output | Locked assignment CSV | No final hashes change | Runtime passed |

## 15. Regression reference case

```text
FG: 3031116
Source: 125 × 45 × 210 mm
Remeasured: 120 × 50 × 200 mm
Carton: 305 × 246 × 206 mm
Target: 12
Orientation: 50 × 120 × 200
Arrangement: 6 × 2 × 1
Expected status: FIT_REMEASURED_NOMINAL
```

## 16. Safe modification procedure

1. Copy or commit current files as a baseline.
2. Inspect exact functions before editing.
3. Make the minimum possible change.
4. Do not regenerate the HTML.
5. Do not run a whole-file formatter.
6. Do not reorder translations.
7. Produce and inspect a diff.
8. Revert unrelated changes.
9. Run syntax checks.
10. Run Carton Audit regressions after every HTML change.
11. Run lock tests after every launcher change.
12. Report tests honestly.

## 17. Handover checklist

- [x] Current files inspected
- [x] README matches actual implementation
- [x] No hardcoded absolute paths remain
- [ ] Folder names standardized
- [x] Logging failure is nonfatal
- [x] Unique per-run logging confirmed
- [x] Single-instance mutex confirmed
- [x] Locked workbooks are skipped
- [x] Next unlocked BOM selection confirmed
- [x] Cache required-file list confirmed
- [x] Generated writes are atomic and rollback-protected
- [x] Cache state is written last
- [x] Carton Audit single assignment passed
- [x] Carton Audit no assignment passed
- [x] Carton Audit multiple assignment passed
- [ ] Quick Batch active FG runtime passed
- [ ] Consolidation evidence runtime passed
- [x] Regression FG 3031116 passed
- [x] JavaScript syntax passed
- [x] PowerShell parser passed
- [x] CMD relative paths passed
- [x] Mojibake scan passed

## Instructions for the next maintainer

Do not assume earlier prompts were implemented. Inspect the current source first.

The highest-risk areas are:

1. Launcher file-lock handling
2. Generated-data consistency
3. Carton Audit autofill
4. Quick Batch active-FG consistency
5. Consolidation recommendation evidence
6. UTF-8 encoding

A locked log must never stop startup.  
A locked workbook must never be loaded.  
A locked newest BOM must not block the next valid unlocked BOM.  
A failed output replacement must not destroy the previous valid dataset.  
Carton Audit selector elements must not be deleted.
