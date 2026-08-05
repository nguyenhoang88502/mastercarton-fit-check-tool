You are writing vanilla JavaScript for one existing web app file. Output ONLY the JavaScript snippet requested below — no Markdown code fences, no explanation, no commentary, just the raw JS that will be inserted directly into an existing function body.

## Context

Internal manufacturing tool, single HTML file, vanilla JS, no framework, no build step, must run in a normal browser (ES2020 features are fine: optional chaining, arrow functions, async/await, template literals, destructuring).

A background helper (a local HTTP listener on `http://localhost:3000`) lets the page save new rows for finished-goods and cartons. It's already wired up through an object called `SavedEntries`, which exposes:

```js
SavedEntries.save(kind, payload) // async. kind is 'fg-dimensions' or 'carton-dimensions'.
                                  // For 'carton-dimensions', payload shape is:
                                  //   { carton_id: string, length_mm: number, width_mm: number, height_mm: number }
                                  // Returns { ok: boolean, body: { success: boolean, replaced: boolean, ...} | null }
                                  // On success, body.success is true and body.replaced tells you whether
                                  // it overwrote an existing row (true) or was a brand new row (false).
                                  // On failure, body may contain { details: [{field, message}, ...] } or
                                  // { detail: string } or { error: string }. Network/offline failures throw.
```

An already-working example of this exact pattern, saving a finished-good product from a form (for reference only — do not repeat this code, it already exists elsewhere in the file):

```js
$$('#saveProductButton').addEventListener('click', async () => {
  const box = $$('#saveProductMessage');
  const id = $$('#manualFgId').value.trim();
  const length = parseFloat($$('#manualLength').value);
  const width = parseFloat($$('#manualWidth').value);
  const height = parseFloat($$('#manualHeight').value);
  if (!id) { say(box, t('saveNeedId'), 'bad'); $$('#manualFgId').focus(); return; }
  if (!(length > 0 && width > 0 && height > 0)) { say(box, t('saveNeedSize'), 'bad'); return; }
  const button = $$('#saveProductButton');
  button.disabled = true;
  try {
    const payload = { fg_id: id, length_mm: length, width_mm: width, height_mm: height, source: 'MANUAL_ENTRY' };
    const res = await SavedEntries.save('fg-dimensions', payload);
    if (res.ok && res.body && res.body.success) {
      say(box, res.body.replaced ? t('savedReplaced') : t('savedOk'), 'ok');
    } else {
      const detail = res.body && Array.isArray(res.body.details)
        ? res.body.details.map(d => d.message).join(' ')
        : (res.body && (res.body.detail || res.body.error)) || '';
      say(box, t('saveFailed') + detail, 'bad');
    }
  } catch (e) {
    say(box, t('saveOffline'), 'bad');
    SavedEntries.probe();
  } finally { button.disabled = false; }
});
```

In that same enclosing scope (a function that runs once, after the page's data is loaded), these helpers already exist and you should use them exactly as-is, do not redefine them:
- `$$(selector)` — shorthand for `document.querySelector(selector)`
- `t(key)` — looks up an i18n string by key, returns the localized string
- `say(el, text, tone)` — sets a status message element's text and CSS tone class (`'ok'`, `'bad'`, `'warn'`, or `''`)
- `toolUi` — the single instance of the app's main controller class. `toolUi.data.cartons` is a plain JS array of carton objects, each shaped like `{ carton_id: string, carton_name: string, width: number, depth: number, height: number, lifecycle: string, ... }` (width/depth/height are the OUTSIDE dimensions in millimeters). A carton counts as having "missing measurements" when `!CartonFitEngine.validCarton(carton)` is true, where `CartonFitEngine.validCarton(c)` is a globally available function defined as `[c?.width, c?.depth, c?.height].every(n => Number.isFinite(n) && n > 0)` — i.e. any of the three being `NaN`, zero, or negative counts as missing. `CartonFitEngine` is a global class available in this scope, call it directly (`CartonFitEngine.validCarton(...)`), do not redefine it.
- `normalizeId(value)` — a globally available function that normalizes an ID string for comparison (trims, and is what the rest of the codebase uses to compare carton/FG ids). Use it when matching a carton_id against `toolUi.data.cartons`.

I already added the following markup to the page (ids given, so you can bind to them):

- `<select id="newProductCartonSourceMode">` with values `"none"` (default), `"missing"`, `"new"`.
- `<div id="cartonSourceMissingFields" hidden>` containing `<select id="cartonSourceMissingSelect" disabled>` (starts with just a placeholder `<option value="">`), three number inputs `#cartonSourceMissingLength` / `#cartonSourceMissingWidth` / `#cartonSourceMissingHeight`, and a button `#cartonSourceMissingSave`.
- `<div id="cartonSourceNewFields" hidden>` containing a text input `#cartonSourceNewId`, a text input `#cartonSourceNewName` (optional, may be blank), three number inputs `#cartonSourceNewLength` / `#cartonSourceNewWidth` / `#cartonSourceNewHeight`, and a button `#cartonSourceNewSave`.
- A single shared status element `<div id="cartonSourceMessage" class="save-msg" role="status" aria-live="polite"></div>` used by both sub-blocks.

These i18n keys already exist and you should use `t(...)` with them (do not invent replacements): `saveNeedId`, `saveFailed`, `saveOffline`, `savedOk`, `savedReplaced`.
These NEW i18n keys will be added separately (by someone else) to the EN/VI dictionaries — just reference them by these exact names via `t(...)`, do not worry about their translated text:
- `cartonSourceNeedId` — shown when the "new carton" ID field is empty
- `cartonSourceNeedSize` — shown when required dimension fields are missing/non-positive
- `cartonSourceNeedSelect` — shown when "missing" mode is active but no carton is selected in the dropdown
- `cartonSourcePicked` — shown after a carton is selected in "missing" mode, confirming it was picked (optional use, informational)

## What to generate

A block of JavaScript, written as plain statements (NOT wrapped in an IIFE, NOT wrapped in a function declaration — it will be pasted directly into the body of an existing function, after the `#saveProductButton` click handler shown above, in the same scope) that:

1. Grabs the 4 key elements once: mode select, missing-fields div, new-fields div, missing-select dropdown. (Use `$$(...)`.)

2. Adds a `'change'` listener on `#newProductCartonSourceMode` that:
   - Shows `#cartonSourceMissingFields` and hides `#cartonSourceNewFields` when value is `"missing"` (toggle each div's `.hidden` boolean property, not the attribute string).
   - Shows `#cartonSourceNewFields` and hides `#cartonSourceMissingFields` when value is `"new"`.
   - Hides both when value is `"none"`.
   - Clears `#cartonSourceMessage` text content whenever the mode changes.
   - When switching TO `"missing"` mode, call a function (define it, name it `populateMissingCartonOptions`) that rebuilds the `<option>` list inside `#cartonSourceMissingSelect`: it should read `toolUi?.data?.cartons || []`, filter to cartons where `!CartonFitEngine.validCarton(carton)`, sort them by `carton_id` (locale-aware, `localeCompare`), and set `cartonSourceMissingSelect.innerHTML` to a placeholder option (value="", disabled, selected, text = `t('cartonSourceMissingPrompt')`) followed by one `<option>` per filtered carton with `value` = the carton's `carton_id` and visible text = the carton_id (plus `— carton_name` appended in parentheses if `carton_name` is non-empty, e.g. `1008700 (Retail 6-pack)`). Escape any HTML-special characters in carton_id/carton_name when building this via string concatenation (write a tiny local escape helper inline, e.g. replacing `&`, `<`, `>`, `"` — do not assume one already exists in scope). Enable the select (`disabled = false`) if there is at least one option, otherwise leave/set it `disabled = true` and show `t('cartonSourceMissingPrompt')` as the only option (there are zero cartons with missing data — this is a normal, expected state, not an error).

3. Adds a `'change'` listener on `#cartonSourceMissingSelect` that, when a real carton_id is chosen (not the empty placeholder), looks up that carton in `toolUi.data.cartons` (match by `normalizeId(carton.carton_id) === normalizeId(selectedValue)`) and pre-fills `#cartonSourceMissingLength` / `#cartonSourceMissingWidth` / `#cartonSourceMissingHeight` with its current `width` / `depth` / `height` ONLY if those are finite positive numbers (leave the field blank/untouched if not, since the whole point is these are often missing) — use `Number.isFinite(n) && n > 0` as the check, and set `.value = ''` (not leave stale) when the check fails, so switching between two different missing cartons doesn't leave a stale number from the previous selection.

4. Adds a `'click'` listener on `#cartonSourceMissingSave` that:
   - Reads the selected carton_id from `#cartonSourceMissingSelect` (bail out with `say(box, t('cartonSourceNeedSelect'), 'bad')` if empty, where `box` is `$$('#cartonSourceMessage')`).
   - Reads length/width/height as `parseFloat` from the three inputs; bail out with `say(box, t('cartonSourceNeedSize'), 'bad')` if any is not `> 0`.
   - Disables the save button, calls `SavedEntries.save('carton-dimensions', { carton_id: <id>, length_mm: <length>, width_mm: <width>, height_mm: <height> })`, and on success (`res.ok && res.body && res.body.success`):
     - Shows `say(box, res.body.replaced ? t('savedReplaced') : t('savedOk'), 'ok')`.
     - Finds the matching carton object in `toolUi.data.cartons` (by `normalizeId`) and, if found, mutates it in place: sets its `width`, `depth`, `height` to the newly saved length/width/height, so the very next "Check cartons" run in this same session already sees the fixed measurements without a page reload.
     - Calls `populateMissingCartonOptions()` again to refresh the dropdown (the carton may no longer belong in the "missing" list).
   - On failure or thrown error, follow the exact same error-message pattern as the `#saveProductButton` handler shown above (details array / detail / error fallback chain on failure response; `t('saveOffline')` + `SavedEntries.probe()` on thrown/network error).
   - Re-enables the button in a `finally` block.

5. Adds a `'click'` listener on `#cartonSourceNewSave` that:
   - Reads `#cartonSourceNewId` (trimmed); bail out with `say(box, t('cartonSourceNeedId'), 'bad')` and focus that field if empty.
   - Reads length/width/height as `parseFloat`; bail out with `say(box, t('cartonSourceNeedSize'), 'bad')` if any is not `> 0`.
   - Reads `#cartonSourceNewName` (trimmed, may be empty — it is NOT sent to the server, the save API for carton-dimensions does not accept a name field, this is for local use only, see below).
   - Disables the save button, calls `SavedEntries.save('carton-dimensions', { carton_id: <id>, length_mm: <length>, width_mm: <width>, height_mm: <height> })`, and on success:
     - Shows `say(box, res.body.replaced ? t('savedReplaced') : t('savedOk'), 'ok')`.
     - If no carton with that `normalizeId` already exists in `toolUi.data.cartons`, pushes a new plain object onto that array shaped like the other entries: `{ carton_id: <id>, carton_name: <name or ''>, width: <length>, depth: <width>, height: <height>, defaultQty: 12, lifecycle: 'Active', dimensionType: 'outside', dimensionBasis: 'outside' }` (match the field names used elsewhere in this codebase for carton objects — width/depth/height, not length/width/height, at the object level). If a carton with that id already existed (edge case: user re-entered an id that turns out to already exist), just update its width/depth/height/carton_name in place instead of pushing a duplicate.
     - Clears the three new-carton input fields and the id/name fields back to empty, ready for the next entry.
     - If `#newProductCartonSourceMode` currently has value `"missing"`, calls `populateMissingCartonOptions()` too (harmless no-op if not relevant, but keeps things consistent if the user flips modes).
   - Same error-handling pattern as above (details/detail/error chain on failure, `saveOffline` + `SavedEntries.probe()` on throw).
   - Re-enables the button in a `finally` block.

Do not add HTML. Do not add CSS. Do not redefine `$$`, `t`, `say`, `toolUi`, `CartonFitEngine`, `normalizeId`, or `SavedEntries` — they already exist in the enclosing scope. Output only the raw JavaScript statements (no code fences, no surrounding function wrapper, no explanation before or after).
