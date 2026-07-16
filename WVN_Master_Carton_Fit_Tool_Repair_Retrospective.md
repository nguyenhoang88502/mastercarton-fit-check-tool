# WVN Master Carton Fit Tool — Repair Retrospective

**Date:** 16 July 2026  
**Repository:** `Mastercarton check tool`  
**Current Git HEAD:** `6e0a422` — `Add launcher-prepared carton data and audit comparison`

## Purpose

This document records what we went through while modifying the WAHL Vietnam Master Carton Fit Tool, what I broke or made unreliable, why those failures happened, what was repaired, and what must be done differently in future changes.

This is intentionally candid. The main failure was not one isolated typo. I repeatedly changed too much of a working application at once, mixed calculation-policy changes with UI and launcher changes, and handed changes back before testing the complete company workflow with the real files and paths.

## What the application is supposed to do

The tool is a standalone offline HTML application supported by:

- `WVN_Master_Carton_Fit_Tool.html`
- `launcher.ps1`
- `Open_WVN_Master_Carton_Fit_Tool.cmd`
- The product-dimension workbook
- Weekly BOM workbooks
- Launcher-generated JSON and JavaScript caches

Its core workflows are New Product, FG Lookup, Carton Audit, Batch Audit, and Carton Lookup. It must preserve workbook parsing, BOM assignment loading, exact carton-fit calculations, Quick and Optimized modes, proposal generation, CSV export, audit functions, and 3D visualization.

## What we went through

### 1. A UI redesign hid the actual engineering workflow

The interface was changed into a dashboard with a fixed sidebar. This hid the main inputs and made the tool harder for NPI managers to use. The user then had to request a return to top tabs, visible forms, and results directly below each form.

Commit `208520e` restored the workflow-focused interface, but it was a very large change: 1,474 insertions and 284 deletions across the HTML and README. That scale was risky for what should have been a focused UI correction.

### 2. The standalone HTML was left syntactically broken

Generated JavaScript contained corruption patterns such as missing spread operators, malformed object literals, unbalanced braces, and merged method/class boundaries. The application could not be trusted to parse, much less run.

The dumb mistake here was treating a nearly 500 KB single-file application like text that could safely be regenerated in large blocks. A standalone HTML file with many interdependent classes must be patched surgically and syntax-checked after every edit.

### 3. The packing model was repeatedly replaced instead of stabilized

The Git history shows repeated model changes:

| Commit | Change |
|---|---|
| `91e6c60` | Simplified the fit model and added BOM preprocessing |
| `2abe38b` | Replaced it with ten independent scenarios |
| `eac5c87` | Replaced that with thirty single-axis scenarios |
| `3e0662a` | Reverted the thirty-scenario change |
| `b49b7a9` | Added continuous single-axis ratio search |
| `4d09c81` | Reverted continuous single-axis search |
| `87b4728` | Added bounded multi-axis search |
| `edc4de9` | Repaired verification versus optimization behavior |

This churn made the behavior difficult to reason about and repeatedly invalidated caches, UI labels, result fields, and tests.

The underlying mistake was changing the calculation objective and implementation simultaneously. “Verify that the BOM carton fits” and “optimize utilization” are different objectives. They should never have shared the same stopping rule or evidence ranking.

### 4. Remeasured dimensions did not reliably replace workbook dimensions

When exact remeasured FG dimensions were entered, stale workbook dimensions could remain active through old session state or cache results. This was especially visible for FG `3031116`.

The required exact case was:

- Workbook FG: `125 × 45 × 210 mm`
- Remeasured FG: `120 × 50 × 200 mm`
- Carton: `305 × 246 × 206 mm`
- Target: `12`
- Expected orientation: `50 × 120 × 200 mm`
- Expected arrangement: `6 × 2 × 1`
- Expected status: `FIT_REMEASURED_NOMINAL`

The failure was allowing cached or estimated scenarios to compete with exact measurement evidence. Exact measurements should become the active nominal dimensions, invalidate every affected cache, and drive the current-carton result, rankings, proposals, and 3D views.

### 5. Optimization stopped at the first nearby fit

The search contained a premature stopping rule suitable only for “closest valid fit.” It was incorrectly used for utilization optimization, so the engine could stop shortly after finding a passing deviation instead of evaluating the bounded search and selecting the best candidate found.

This also produced misleading reporting such as `Search: 0 / 5000 · complete` even when an adjusted candidate had been selected. Search accounting, limit status, and optimality wording all had to be repaired.

### 6. Artificial ratio changes were allowed to overpower real evidence

Some iterations adjusted FG and carton dimensions to reproduce a known BOM fit. This could make an estimated result look better than a reliable nominal result and could recommend a carton based on artificial ratios.

That was conceptually wrong. Real remeasured FG dimensions must stay fixed. Recorded carton dimensions must stay fixed for current-carton verification. Ratio search is diagnostic evidence only when nominal checks fail or when the user explicitly requests optimization.

### 7. Recommendations could be worse than the current carton

Alternatives were sometimes presented even when they offered no genuine improvement, and proposal rows could displace valid existing cartons. Commit `bcf6b47`, `Prevent worse carton recommendations`, was needed to restore the correct rule.

The user also clarified that “Suggested alternative” should come from existing cartons. Generated proposals are a separate fallback and must not masquerade as an existing-carton alternative.

### 8. The 5 mm wall rule was added without protecting the launcher workflow

Quick Mode needed one clear conversion:

```text
outside dimension − 10 mm per axis = usable internal dimension
```

That represents a 5 mm wall on each side. Structured internal dimensions must not be deducted again.

The calculation rule itself was straightforward. The failure was coupling it with broader launcher and data-source changes. The tool reached a state where carton thickness behavior existed but the launcher did not reliably prepare or load the BOM. The user correctly asked to restore the working launcher while keeping carton thickness.

### 9. Launcher discovery assumed the wrong company environment

The launcher initially worked against a home-style or local folder layout but failed against the real company SharePoint/OneDrive layout. It selected fallback paths too early and did not consistently search the actual `1. Weekly BOM` location for both the BOM and product-dimension workbook.

The real company ZIP had to be used to reproduce the environment. A key repair was adding the Weekly BOM directory to dimension-workbook discovery instead of assuming the dimension workbook would always be in the tool root or a separately named directory.

The mistake was validating discovery with invented or local paths before validating it with the exact company package.

### 10. BOM preprocessing and browser loading drifted out of contract

The launcher-generated compact BOM index advanced to version 2 while one HTML loader still expected version 1. That kind of producer/consumer mismatch can silently force a slower fallback or make prepared data appear unavailable.

The broader contract also drifted:

- Launcher manifests used local `file:///C:/...` URLs.
- Those URLs work differently from HTTP URLs and cannot be used from a phone through a tunnel.
- Prepared payloads, JSON manifests, cache fingerprints, and browser fallback behavior were not always validated together.
- A cache could be technically present but rejected because its workbook fingerprint, file size, or source metadata did not match the selected workbook.

The dumb mistake was treating “the cache file exists” as equivalent to “the browser will accept and use the cache.” Every producer field must match what the HTML consumer validates.

### 11. The audit comparison selected suggestions too aggressively

The audit UI initially populated the suggested 3D pane automatically. The required behavior was to keep the current carton on the left and leave the right pane empty until the user clicks View 3D or drags an existing carton/proposal into it.

This was another example of adding UI behavior that was not requested and that obscured the baseline-versus-alternative decision.

### 12. Too many unrelated concerns were bundled into single commits

Examples from repository history:

- `91e6c60` changed the HTML, PowerShell launcher, CMD launcher, and `.gitignore` together.
- `6e0a422` added 397 lines and removed 107 lines across all three canonical application files.
- `bcf6b47` changed 365 HTML lines to repair recommendation behavior.

Large commits are not automatically wrong, but these combined algorithm changes, data preparation, UI rendering, cache contracts, and launcher behavior. When a regression appeared, rollback also removed unrelated good behavior.

## Current repaired state

The latest confirmed launcher run completed successfully on 16 July 2026 at 18:15 local time.

Confirmed launcher output:

- Selected BOM: `BOM & FORMULA JUL 10 D365.xlsx`
- Dimension workbook: `Products Physical Dimension - FG 1.Apr.25.xlsx`
- EA dimensions indexed: `687`
- CA dimensions indexed: `732`
- PL dimensions indexed: `707`
- Cartons derived from Dimension CA: `53`
- Cartons using `Mastercarton_Dimensions` backup: `102`
- Cartons still missing dimensions: `1`
- Carton variation warnings: `4`
- BOM carton assignments: `749`
- FGs indexed: `665`
- Cartons indexed from BOM: `180`
- BOM conflicts: `4`
- Browser launch: completed

The Cloudflare Quick Tunnel test also confirmed HTTP 200 responses for:

- The HTML application
- `WVN_Launcher_Prepared_Data.js`
- The product-dimension workbook
- The latest weekly BOM workbook

For the tunnel, a separate HTTP-safe staging copy was created under `.wrangler/mobile-test`. Local `file:///` manifest URLs were converted to relative HTTP URLs in that staging copy only. The canonical application and workbook files were not rewritten for the tunnel.

The Quick Tunnel is temporary, has no uptime guarantee, and exposes its contents to anyone who knows the generated URL. It should be stopped after phone testing.

## Repository state that still needs care

At the time this document was written, Git reported:

- `README.md` deleted in the working tree
- `1. Weekly BOM/BOM & FORMULA JUL 10 D365.xlsx` untracked
- `New folder/` untracked
- `WVN_MasterCarton_Derived_Dimensions.cache.json` untracked
- `backups/` untracked

These items were already present before this documentation file was created. They must not be blindly staged, deleted, or committed. Their ownership and intended retention should be reviewed first.

## Rules that should prevent another repeat

### Change discipline

1. Patch the existing implementation; do not regenerate the 500 KB HTML.
2. Make one concern per commit: UI, engine, launcher, cache contract, or documentation.
3. Never change the fit algorithm while repairing launcher discovery.
4. Never change the launcher while repairing a table column or 3D interaction.
5. Record the exact baseline commit before starting.
6. Use `git diff --check` and review the full diff before every commit.

### JavaScript validation

1. Extract and parse every `<script>` block before browser testing.
2. Verify required classes and unique class definitions.
3. Check referenced DOM IDs against the HTML.
4. Check all `data-i18n` keys in both English and Vietnamese.
5. Check event-handler targets and cache invalidation paths.
6. Run the application in a browser and capture console errors.

### Calculation validation

1. Keep nominal verification separate from utilization optimization.
2. Prefer evidence in this order: exact remeasurement, workbook nominal, bounded estimate.
3. Never adjust exact measured FG dimensions.
4. Apply carton wall deduction exactly once.
5. Rank nominal existing-carton fits ahead of estimated-ratio fits.
6. Never recommend a carton that is worse than the current carton.
7. Keep generated proposals separate from existing-carton alternatives.
8. Maintain fixed regression cases, especially FG `3031116`.

### Launcher and cache validation

1. Test against the exact latest company ZIP and real SharePoint/OneDrive layout.
2. Verify discovery separately for the dimension workbook and weekly BOM.
3. Verify the launcher-generated schema version against the HTML consumer.
4. Verify cache fingerprint, source size, source timestamp, and selected workbook together.
5. Test both prepared-cache loading and browser fallback loading.
6. Confirm BOM processing finishes before the browser is considered ready.
7. Preserve the canonical filenames.

### Release checklist

- HTML parses without JavaScript syntax errors.
- No duplicate required classes or methods.
- All required DOM IDs exist.
- All translation keys exist in EN and VI.
- Launcher finds the real company paths.
- Dimension cache is accepted by the HTML.
- BOM cache is accepted by the HTML.
- New Product works.
- FG Lookup works.
- Carton Audit works.
- Batch Audit works and remains responsive.
- Carton Lookup works.
- CSV exports contain the expected fields.
- 3D baseline and selected comparison render correctly.
- Quick Mode applies one wall deduction and no ratio search.
- Optimized Mode starts from the same internal baseline.
- FG `3031116` regression case passes.
- The final package is tested from a clean extracted ZIP, not only from the developer workspace.

## Bottom line

The most damaging pattern was over-editing: I made broad rewrites, changed the calculation policy repeatedly, and validated individual pieces instead of the full launcher-to-browser workflow. That created syntax failures, stale-cache behavior, incorrect search decisions, broken company-path discovery, and repeated rollback requests.

The correct approach is smaller patches, explicit producer/consumer contracts, real company-file testing, fixed regression cases, and end-to-end browser validation before claiming the tool is repaired.
