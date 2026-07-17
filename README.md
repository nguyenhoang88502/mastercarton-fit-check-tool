# WVN Master Carton Fit Tool

Offline-capable, single-file engineering tool for WAHL Vietnam that evaluates whether finished goods (FGs) physically fit inside master cartons and recommends optimized carton assignments to reduce material waste and carton-line proliferation.

## How it works

1. **Launcher** (`system/launcher.ps1`) — discovers the latest weekly BOM and product physical dimension Excel workbooks from OneDrive/SharePoint, preprocesses carton data, and opens the tool in the browser.
2. **Web tool** (`system/WVN_Master_Carton_Fit_Tool.html`) — a self-contained HTML/JS application with no external dependencies. Reads CSV data prepared by the launcher (or Excel files dropped manually), runs rigid-cuboid simple-grid packing calculations, and renders interactive 3D visualizations.

## Quick start

Double-click `Open_WVN_Master_Carton_Fit_Tool.cmd`. The launcher will:

1. Find the latest BOM and dimension workbooks
2. Preprocess carton data into CSV files under `system/generated/`
3. Open the tool in Chrome or Edge

If automatic discovery fails, use the web tool's manual file upload.

**Requirements:** Windows, PowerShell 5.1, Chrome or Edge, synced OneDrive/SharePoint Excel sources.

## Project structure

```
.
├── Open_WVN_Master_Carton_Fit_Tool.cmd   # Launch script (double-click to start)
├── Quick Start.txt                        # User-facing setup instructions
├── 1. Weekly BOM/                         # Source Excel workbooks (synced weekly)
├── system/
│   ├── WVN_Master_Carton_Fit_Tool.html    # Main application (single-file, zero-dependency)
│   ├── launcher.ps1                       # PowerShell launcher — discovers, preprocesses, caches
│   ├── config.ini                         # Search paths, file patterns, launcher settings
│   ├── generated/                         # CSV output from launcher preprocessing
│   │   ├── fg_dimensions.csv
│   │   ├── carton_dimensions.csv
│   │   ├── bom_carton_assignments.csv
│   │   ├── carton_dimension_sources.csv
│   │   ├── carton_fg_reverse_index.csv
│   │   ├── data_quality.csv
│   │   └── cache_state.ini
│   └── logs/                              # Launcher logs
└── .gitignore
```

## Key features

- **New Product / NPI Carton Check** — enter FG dimensions manually and check against the loaded master-carton list
- **FG Lookup** — select an existing FG and see which cartons fit with the target quantity
- **Carton Audit** — audit one FG/carton pairing and get ranked existing-carton alternatives with utilization comparisons
- **Batch Audit** — evaluate all eligible FGs in bulk with Quick Mode (exact dimensions) or Optimized Mode (multi-axis ratio search)
- **Carton-Line Reduction** — greedy set-cover analysis recommending how existing FGs can share fewer distinct carton IDs
- **Carton Lookup** — reverse lookup to see all FGs assigned to a given carton
- **Side-by-side 3D comparison** — compare current vs. suggested carton packing with synchronized rotation and drag-and-drop
- **Shared Carton Proposals** — generate new carton dimensions that serve multiple FGs
- **English / Vietnamese** — full bilingual UI

## Calculation modes

| Mode | Behavior |
|------|----------|
| **Quick Mode** | Exact source dimensions with 5 mm wall deduction per side. Six orthogonal orientations. No tolerance or ratio adjustment. |
| **Optimized Mode** | Screens candidates with Quick Mode, then applies 125 fixed-carton FG-ratio combinations (0.05 step, independent axes) for best utilization. |

## Engineering assumptions

- Rigid rectangular cuboid simple-grid packing (no compression, nesting, or diagonal placement)
- 5 mm carton wall deduction per side (unless source is marked as internal dimensions)
- Six orthogonal orientations tested
- Results are advisory only — confirm with physical samples before production release

## Configuration

Edit `system/config.ini` to adjust:

- `WeeklyBomFolder` — where to search for weekly BOM files
- `Root1`/`Root2`/`Root3` — additional search roots for dimension workbooks
- `MaxDepth` — folder search depth
- `ForceRebuild` — set to `true` to always regenerate CSVs
- `DimensionPattern` / `BomPattern` — file name globs for auto-discovery

## Git notes

- Generated CSVs, logs, and Wrangler state are gitignored
- Source Excel workbooks in `1. Weekly BOM/` are tracked for reproducibility
- The HTML tool is self-contained — no `npm install`, no build step
