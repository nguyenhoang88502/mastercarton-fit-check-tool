# WVN Master Carton Fit Tool — Maintainer Notes

Offline, single-file decision-support tool for WAHL Vietnam carton engineering. The application reads the dimension workbook and latest Weekly BOM directly in Chrome or Edge, reconciles FG/carton truth, and provides New Product, FG Lookup, Carton Audit, Batch Audit, Carton Lookup, CSV export, and SVG packing comparison workflows.

The delivered application is `WVN_Master_Carton_Fit_Tool.html`. It has no CDN, framework, npm, Node, Python, server, or build-time dependency on the company machine.

## Launch and data loading

Run `Open_WVN_Master_Carton_Fit_Tool.cmd`. The existing launcher remains authoritative and:

1. Detects the local OneDrive/NPI tools path.
2. Selects the dimension workbook and newest Weekly BOM.
3. Writes `WVN Master Carton Workbook Manifest.json` and `Weekly BOM Manifest.json`.
4. Opens the HTML through `file:///` in Chrome or Edge with `--allow-file-access-from-files`.

The HTML loads the dimension workbook first, then the Weekly BOM, reconciles assignments and quantities, builds indexes, and only then marks data ready. Manual `.xlsx`/`.xlsm` selection remains available under Workbook. Use `Open_WVN_Master_Carton_Fit_Tool.cmd --manifests-only` to validate manifest discovery without opening the browser.

## Workflows

- **Dashboard** — source counts, data readiness, and engineering settings.
- **New Product** — manual NPI dimensions, existing-carton reuse, shared-carton suggestions, and new-carton proposals.
- **FG Lookup** — current assignment, BOM linkage, fit classification, recommendations, proposal, and 3D view.
- **Carton Audit** — preserved single-pair historical audit and existing-carton alternatives.
- **Batch Audit** — chunked, cancellable portfolio audit with filters, pagination, consolidation groups, and UTF-8 CSV.
- **Carton Lookup** — reverse index from carton ID to assigned FGs, audit-all handoff, and CSV.

## Dimension and uncertainty model

All normalized dimensions are millimetres. Source carton dimensions are treated as outside dimensions.

Default settings:

| Setting | Default |
|---|---:|
| FG measurement error | ±5 mm per axis |
| Carton measurement error | ±5 mm per axis |
| Carton wall thickness | 5 mm per side |
| Opening/fold deduction | 10 mm |
| Opening axis | Auto/unknown |
| Minimum utilization for standard reuse | 85% |
| Proposal clearance | 5 mm |
| Proposal rounding | 5 mm upward |

`estimateUsableCartonDimensions()` converts outside dimensions to estimated inside dimensions by subtracting two wall thicknesses. Conservative evaluation additionally expands FG dimensions, shrinks carton dimensions by measurement error, and applies the fold deduction. If the opening axis is unknown, length, width, and height opening cases are all evaluated and the least favorable result is used.

The six-axis orthogonal grid remains the primary packing model. `OrthogonalPackingSolver` performs bounded mixed-orientation search for eligible targets (default maximum target 12 and maximum 150,000 nodes). Node-limit or target-limit results are reported as search-incomplete, never silently promoted to a fit.

## Decision classifications

`SharedDecisionEngine` is the common fit contract for current-carton checks, recommendations, reverse lookup, and batch audit.

| Classification | Meaning |
|---|---|
| `STANDARD_REUSE_RECOMMENDED` | Conservative fit passes and utilization meets the configured minimum. |
| `FIT_BUT_LOW_UTILIZATION` | Conservative fit passes but utilization is below the minimum. |
| `NOMINAL_FIT_PHYSICAL_CHECK` | Nominal fit passes but conservative fit fails. |
| `FAVORABLE_ONLY_PHYSICAL_CHECK` | Only the favorable scenario passes. |
| `SEARCH_INCONCLUSIVE` | A bounded packing search ended without a definitive result. |
| `DOES_NOT_FIT` | Evaluated scenarios do not meet the target. |
| `DATA_INCOMPLETE` | Required dimensions, target, or assignment are missing. |

Only conservative fits at or above the utilization threshold enter standard existing-carton recommendations. Nominal-only candidates are separated for physical validation. New-carton proposals use integer layouts, conservative FG dimensions, configured clearance/wall/fold/measurement allowances, and configured upward rounding. All proposals remain advisory and require engineering/sample validation.

## Data architecture

`WeeklyBomLoader` preserves IDs as text, including leading zeros. `BomTruthReconciler` uses Weekly BOM as the authoritative assignment and target source when available. Target derivation order is `perSeries`, `norm`, then `quantity`; invalid or absent quantities remain missing.

`AppDataStore` builds and validates these indexes:

- `fgById`
- `cartonById`
- `bomRowsByFg`
- `fgsByCartonId`
- `assignedCartonByFg`
- `targetQtyByFg`
- `dataWarningsByFg`
- `cartonsByNormalizedDimensions`

Long-running reconciliation, index construction, and batch auditing yield cooperatively to the browser. `TaskProgressManager` owns the persistent task state, actual counters, elapsed time, and cancellation flag. Fit and decision caches are bounded to prevent unbounded memory growth.

## Key classes

| Class | Role |
|---|---|
| `WorkbookLoader`, `SimpleZip` | Browser-native XLSX reading |
| `WeeklyBomLoader`, `DataNormalizer` | Source parsing and mm normalization |
| `BomTruthReconciler`, `AppDataStore` | Source-of-truth reconciliation and indexes |
| `FitSettings`, `DimensionSettings` | Compatibility tolerance and engineering settings |
| `CartonFitEngine`, `OrthogonalPackingSolver` | Uniform and bounded mixed packing |
| `SharedDecisionEngine` | Common uncertainty-aware decision result |
| `RecommendationEngine` | Existing reuse, new proposals, consolidation |
| `BatchAuditEngine` | Chunked portfolio evaluation |
| `FitInterpretation`, `ToleranceClassifier` | Preserved legacy interpretation contracts |
| `PackingVisualizer3D` | Single and side-by-side SVG packing views |
| `UiController`, `TaskProgressManager` | Workflows, loading, progress, rendering, export |

## Maintainer validation checklist

1. Run the launcher and confirm both manifests load without fallback warnings.
2. Confirm Dashboard reports non-zero FG, carton, and BOM counts and “Data ready”.
3. Select an FG and verify linkage, existing results, audit, and SVG controls.
4. Enter an NPI and verify reuse is considered before a new-carton proposal.
5. Run a batch audit, change every filter, cancel a run, paginate, and export CSV.
6. Reverse-lookup a carton, export its FG list, and audit all users of that carton.
7. Toggle EN/VI and dark/light themes after results exist.
8. Confirm browser console has no uncaught errors.
9. Confirm every `data-i18n` key exists in both dictionaries and every referenced DOM ID exists.
10. Confirm no result claims production approval; physical sample validation remains visible.

Known modelling limits: no diagonal insertion, nesting, flexible compression, corrugate strength, regulatory compliance, pallet optimization, or production approval. Formula cells rely on cached workbook values.
