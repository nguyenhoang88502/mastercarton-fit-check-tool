You are writing HTML markup for one existing web app file. Output ONLY the HTML snippet requested below — no Markdown code fences, no explanation, no commentary, just the raw HTML that will be inserted directly into the page.

## Context

This is an internal manufacturing tool (single static HTML file). The "New Product / NPI Manual Check" section already has a form for entering a finished-good product and a "Save this product" button that only saves the product's own dimensions. It has NO way to add or fix a CARTON's dimensions from this tab today. I need a new block, gated behind a dropdown, that lets a non-technical user either fix an existing carton that is missing its measurements, or register a brand new carton line.

Here is the existing form, verbatim, so you match its exact style and conventions (indentation, class names, `data-i18n` usage, `data-i18n-placeholder` usage, `<span class="unit">mm</span>` pattern, `dimension-row` wrapper for L/W/H triplets):

```html
<form id="manualForm" class="manual-form">
  <div class="form-block identity-block">
    <h3 data-i18n="identityGroup">Product identity</h3><label><span data-i18n="fgIdOptional">New FG ID
        (optional)</span><input id="manualFgId" type="text" maxlength="80"
        data-i18n-placeholder="fgIdOptionalPlaceholder" placeholder="Optional"></label>
  </div>
  <div class="form-block specs-block">
    <h3 data-i18n="dimensionGroup">Finished Good Dimensions (mm)</h3>
    <div class="dimension-row"><label><span data-i18n="length">Length</span><input id="manualLength" type="number"
          min="0.1" step="0.1" required placeholder="mm"><span class="unit">mm</span></label><label><span
          data-i18n="width">Width</span><input id="manualWidth" type="number" min="0.1" step="0.1" required
          placeholder="mm"><span class="unit">mm</span></label><label><span data-i18n="height">Height</span><input
          id="manualHeight" type="number" min="0.1" step="0.1" required placeholder="mm"><span
          class="unit">mm</span></label></div>
  </div>
  <div class="form-block target-block">
    <h3 data-i18n="targetGroup">Target quantity</h3><label class="target-qty"><span data-i18n="targetQty">Target
        quantity</span><input id="manualQty" type="number" min="1" step="1" required value="3"
        placeholder="3"></label>
  </div>
  <div class="manual-actions"><button class="primary" type="submit" data-i18n="checkCartons">Check
      cartons</button><button id="suggestSharedCartonBtn" type="button" data-i18n="suggestSharedCarton">Suggest
      shared carton line</button><button id="saveProductButton" class="save-feature" type="button"
      data-i18n="saveProduct">Save this product</button></div>
  <div id="saveProductMessage" class="save-msg" role="status" aria-live="polite"></div>
</form>
```

A `<select>` elsewhere in the app that uses `data-i18n` directly on its `<option>` elements (this is the correct pattern for select options in this codebase, follow it exactly):

```html
<select id="auditProductionConfirmed">
  <option value="unknown" data-i18n="unknown">Unknown</option>
</select>
```

Save-only UI elements are gated with `class="save-feature"` (hidden unless a background helper is running — a global JS toggle handles showing it, you don't need to worry about that mechanism, just apply the class).

## What to generate

A new `<div class="form-block carton-source-block save-feature">...</div>` block, meant to be inserted as a new form-block INSIDE `<form id="manualForm">`, placed after the closing `</div>` of `target-block` and before `<div class="manual-actions">`. Structure exactly as follows (keep every id and value exactly as given — other code depends on these exact ids):

1. A heading: `<h3 data-i18n="cartonSourceGroup">` with visible text "If a carton needs measurements".
2. A label wrapping a `<select id="newProductCartonSourceMode">` with exactly these 3 options (id="newProductCartonSourceMode", each `<option>` uses `data-i18n` on the option tag itself per the pattern above):
   - `<option value="none" data-i18n="cartonSourceNone">` visible text "No carton changes needed" (this option should have the `selected` attribute)
   - `<option value="missing" data-i18n="cartonSourceMissing">` visible text "Fix a carton with missing measurements"
   - `<option value="new" data-i18n="cartonSourceNew">` visible text "Add a brand new carton"
   The label's own visible span text (data-i18n="cartonSourceLabel") should read "Carton measurement".
3. A `<div id="cartonSourceMissingFields" hidden>` block (hidden by default, JS elsewhere toggles the `hidden` attribute) containing:
   - A label wrapping `<select id="cartonSourceMissingSelect" disabled>` with one placeholder `<option value="" disabled selected data-i18n="cartonSourceMissingPrompt">` visible text "— Select a carton —". (JS elsewhere fills in the rest of the options at runtime, you don't need to generate them.) The label's own span text (data-i18n="cartonSourceMissingSelectLabel") should read "Which carton".
   - A `dimension-row` div with three labelled number inputs, following the EXACT pattern of the Length/Width/Height triplet in `specs-block` above (same `<span class="unit">mm</span>` markup, same `min="0.1" step="0.1"` attributes, NOT `required` since the field starts empty and only becomes relevant once a carton is picked), with ids `cartonSourceMissingLength`, `cartonSourceMissingWidth`, `cartonSourceMissingHeight`, reusing the existing `data-i18n="length"` / `data-i18n="width"` / `data-i18n="height"` keys for the labels (do not invent new i18n keys for these three — reuse the existing ones verbatim).
   - A `<button id="cartonSourceMissingSave" type="button" data-i18n="cartonSourceSaveMeasurements">` visible text "Save these measurements".
4. A `<div id="cartonSourceNewFields" hidden>` block (hidden by default) containing:
   - A label wrapping `<input id="cartonSourceNewId" type="text" maxlength="80" placeholder="e.g., 1008700">`, span text (data-i18n="cartonSourceNewId") "New Carton ID".
   - A label wrapping `<input id="cartonSourceNewName" type="text" maxlength="120" placeholder="Optional">`, span text (data-i18n="cartonSourceNewName") "Carton name (optional)".
   - A `dimension-row` div with three labelled number inputs identical in structure to the one in block 3, ids `cartonSourceNewLength`, `cartonSourceNewWidth`, `cartonSourceNewHeight`, reusing the existing `length`/`width`/`height` i18n keys.
   - A `<button id="cartonSourceNewSave" type="button" data-i18n="cartonSourceSaveNew">` visible text "Save new carton".
5. A single shared status line for both sub-blocks: `<div id="cartonSourceMessage" class="save-msg" role="status" aria-live="polite"></div>`, placed after both conditional field blocks (outside both, always present in the DOM).

Output only the raw HTML for this one `<div class="form-block carton-source-block save-feature">...</div>` block (nothing before or after it, no code fences, no other blocks). Match indentation style (2-space, compact single/double line style like the example) as closely as you reasonably can.
