# WAHL Vietnam — Master Carton Fit Tool

Offline Windows tool for checking finished-good fit against existing master cartons, auditing current BOM cartons, comparing alternatives, and identifying potential carton-line consolidation opportunities.

No installation, administrator rights, Node.js, Python, internet connection, or local server is required.

## 1. What the tool does

- **New Product:** Enter FG dimensions and target quantity, check existing cartons, generate possible shared-carton proposals, and view packing in 3D.
- **FG Lookup:** Select an existing FG, resolve its BOM carton, compare alternatives, and review fit, capacity, utilization, and 3D packing. A session-only carton override is available per FG; it does not change the BOM or generated data.
- **Carton Audit:** Audit one FG/current-carton pair. Prepared BOM assignments support automatic single-carton entry, manual entry when no assignment exists, and a selector for multiple or conflicting assignments.
- **Batch Audit:** Run Quick Batch Audit, Optimized Batch Suggestions, and carton-line-reduction analysis, then export results.
- **Carton Lookup:** Search a carton, review its dimensions and linked FGs, and display carton pictures and packing context where available.

## 2. Quick and Optimized modes

**Quick Mode** uses the active FG dimensions, including a session remeasurement when present; otherwise it uses source FG dimensions. Carton dimensions remain fixed. It tests six orthogonal orientations using a rigid rectangular grid, with no compression, nesting, diagonal placement, or ratio adjustment.

**Optimized Mode** keeps launcher-prepared carton dimensions fixed and independently tests FG length, width, and height ratios of `0.90`, `0.95`, `1.00`, `1.05`, and `1.10`. This produces 125 FG combinations and six orientations, or 750 checks per FG-carton pair. Carton ratios remain fixed at `1.00`.

## 3. Required folder structure

```text
5. Tools for NPI/
├─ 1. Weekly BOM/
│  ├─ Products Physical Dimension - FG 1.Apr.25.xlsx
│  ├─ BOM & FORMULA [MONTH] [DAY] D365.xlsx
│  └─ master-carton-pictures/
└─ WVN_Master_Carton_Fit_Tool/
   ├─ Open_WVN_Master_Carton_Fit_Tool.cmd
   ├─ Quick_Start.txt
   ├─ README.md
   ├─ HANDOVER.md
   ├─ .gitignore
   └─ system/
      ├─ WVN_Master_Carton_Fit_Tool.html
      ├─ launcher.ps1
      ├─ config.ini
      ├─ generated/
      └─ logs/
```

The folder containing the CMD file is the application root. Do not create duplicated nesting such as `WVN_Master_Carton_Fit_Tool/WVN_Master_Carton_Fit_Tool/`.

The current checkout still uses `Quick Start.txt`, a nonstandard application-folder name, and the legacy picture-folder name `Master_Carton_Pics`; see `HANDOVER.md`.

## 4. How to launch

1. Close source workbooks currently open in Excel.
2. Confirm the latest BOM and dimension workbook are in `1. Weekly BOM`.
3. Double-click `Open_WVN_Master_Carton_Fit_Tool.cmd`.
4. Wait for preprocessing.
5. The tool opens in Chrome or Edge.

Only one launcher preparation process runs at a time. A second launch displays a friendly message and exits cleanly.

## 5. Source workbooks

The canonical dimension workbook is `Products Physical Dimension - FG 1.Apr.25.xlsx`; matching alternative dimension workbooks may be considered when the canonical file is unavailable or invalid.

Weekly BOM names follow `BOM & FORMULA [MONTH] [DAY] D365.xlsx`. The launcher sorts candidates by inferred business date, skips locked or invalid files, and selects the latest valid unlocked candidate. A locked file is not hashed, validated, parsed, or copied. If the newest BOM is locked, the next valid unlocked BOM is used.

When no valid unlocked source is available, the launcher reuses generated data only when all required files exist and the cache schema, preprocessing version, source paths, sizes, timestamps, and stored fingerprints match. Otherwise it stops with a clear message.

## 6. Generated data

The launcher converts workbook data into smaller files that the browser loads quickly:

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

Do not edit these files manually.

## 7. Carton dimension preparation

The BOM maps each FG to a carton. Valid CA measurement rows are linked through their FG, source values are multiplied by 10 to convert centimetres to millimetres, axes are normalized, and measurements are averaged by resolved carton ID. The dimension source is `BOM_LINKED_CA_AVERAGE`. Conflicting or missing records are reported in `data_quality.csv`.

## 8. Carton wall rule

Outside carton dimensions use a 5 mm wall per side: a 10 mm total deduction per outside axis, applied exactly once. Sources marked as internal dimensions are not deducted again.

## Locked files and safe startup

- Locked workbooks are skipped and never loaded.
- A locked newest BOM does not block the next valid unlocked BOM.
- Log-writing failures warn but do not stop startup.
- Generated outputs are staged, validated, and rollback-protected; replacement failure preserves the previous complete dataset.
- `cache_state.ini` is published last.

## 10. Carton Audit behavior

- **One assignment:** enter FG → carton ID autofills.
- **No assignment:** enter FG → manager enters the carton manually.
- **Multiple/conflicting assignments:** enter FG → selector appears → manager selects the current carton.

Maintainers must not delete `auditMcSelect`, `auditMcSelectLabel`, or `auditMcSelectNote`; they are part of autofill and form validation.

## 11. Master-carton pictures

Pictures are indexed by exact carton ID and displayed below the relevant FG Lookup or Carton Lookup 3D context. Put files in `1. Weekly BOM/master-carton-pictures/` using names such as:

```text
1008609_image1.png
1008609_image2.png
1008609_image2_1.png
```

Matching uses the exact filename pattern rather than substring matching. Images use `object-fit: contain`; they are display-only and do not affect fit, capacity, utilization, ranking, or recommendations.

## 12. Troubleshooting

| Symptom | Likely cause | Action |
| --- | --- | --- |
| Launcher says workbook is locked | Excel, preview, Teams, or synchronization lock | Close the file and relaunch |
| Latest BOM was skipped | File locked or invalid | Review the latest launcher log |
| Carton Audit does not autofill | Missing or stale `bom_carton_assignments.csv` | Close source files and relaunch |
| Browser opens with no data | Generated files missing or unavailable | Review `system/logs/` |
| Second launch does nothing | Another launcher is preparing | Wait for the first launch |
| No carton picture found | No exact indexed filename | Check the naming convention |
| Generated CSV cannot be replaced | CSV or preview is holding the file | Close it and relaunch |
| Wrong folder selected | Duplicate or incorrect nesting | Keep the CMD and `system` together |

## 13. Logs

Logs are stored in `system/logs/` with unique per-run names:

```text
WVN_Master_Carton_Fit_Tool_launcher_yyyyMMdd_HHmmss_fff_PID.log
```

Older fixed-name logs may remain from previous launcher versions but are no longer used.

## 14. Advisory statement

**All carton recommendations are advisory until confirmed by physical packing validation and the responsible engineering/production team.**

## Ownership

- Business owner: TBD
- Technical owner: TBD
- Source-data owner: TBD
- Weekly BOM owner: TBD
- Dimension workbook owner: TBD
