You are writing JavaScript for one existing web app file. Output ONLY the code requested below in the three labeled sections — no Markdown code fences, no explanation, no commentary, just the raw code that will be pasted directly into the page. Label each section exactly as `=== SECTION 1 ===`, `=== SECTION 2 ===`, `=== SECTION 3 ===` so it can be split back out mechanically.

## Context

This is an internal manufacturing tool (single static HTML file, vanilla JS, no framework/build step). The page already has a `SavedEntries` module (an IIFE) that talks to a local HTTP listener at `http://localhost:3000` to persist rows the user enters into "New Product / NPI Manual Check" form. Two buttons already exist and already work: `#saveProductButton` ("Save Product") posts the finished-good fields to `/api/data/fg-dimensions`, and `#saveCartonButton` ("Save Carton") posts the carton fields to `/api/data/carton-dimensions`. Both already validate, save, show a message, and are fully independent of each other (saving a carton does not require a product ID and vice versa). Do not change any of that existing behavior.

A NEW backend endpoint, `/api/save-dimension`, now also exists on the same listener. It writes the same product/carton dimensions into the live source-of-truth Excel workbook via COM automation (not just the CSV overlay the existing buttons already write to), then kicks off a background CSV cache rebuild. It is authoritative but SLOW: end-to-end it takes roughly 5-20 seconds depending on machine load, because it has to start a real Excel process, open, write, save, and close the file. It must never block the fast CSV save the buttons already do - it runs strictly AFTER the existing CSV save succeeds, and it must not delay re-enabling the Save button or freeze the UI while it runs.

Request body the endpoint expects (JSON):
```json
{ "type": "product", "id": "NPI-2026-001", "length_mm": 225, "width_mm": 130, "height_mm": 55 }
```
or
```json
{ "type": "carton", "id": "CTN-NPI-01", "name": "Master Carton NPI", "length_mm": 400, "width_mm": 300, "height_mm": 250 }
```
`name` is optional and only meaningful for `type: "carton"` (omit the field entirely for `type: "product"`, or for a carton with no name typed).

Successful response (HTTP 200):
```json
{ "success": true, "status": "success", "message": "Saved to Excel and cache updated", "row": 2211, "is_new": true, "sheet": "Dimension", "backup_path": "...", "cache_rebuild_started": true }
```

Failure response (HTTP 502 or 503 - 503 specifically means the workbook is open elsewhere right now):
```json
{ "success": false, "status": 503, "error": "The workbook opened read-only, which means it is already open elsewhere. Close it there and try again." }
```
(400 is also possible for a malformed request, but that should never happen here since this call is only ever made with already-validated data - see below.)

## Existing code you must match the style of

The `SavedEntries` IIFE's private HTTP helper, its `save` method, and its public return statement, verbatim:

```js
const BASE = 'http://localhost:3000';

const call = async (path, options = {}, timeoutMs = 8000) => {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  try {
    const res = await fetch(BASE + path, { ...options, signal: controller.signal, cache: 'no-store' });
    let body = null;
    try { body = await res.json(); } catch (e) { body = null; }
    return { ok: res.ok, status: res.status, body };
  } finally { clearTimeout(timer); }
};

const isAvailable = () => available;
```

```js
const save = async (kind, payload) => call('/api/data/' + kind, {
  method: 'POST',
  headers: { 'Content-Type': 'application/json' },
  body: JSON.stringify(payload)
}, 20000);
```

```js
      return { KINDS, isAvailable, probe, save, list, stop, mergeInto };
    })();
```

The two existing button handlers, verbatim (these live in a different function, `initSavingUi`, outside the `SavedEntries` IIFE - they can only call `SavedEntries.<exported method>`, they cannot reach `call` or `BASE` directly since those are private to the IIFE):

```js
      /* ---- New Product tab: save the product currently typed into the form. Independent
         of the carton button below - either can be used on its own. */
      $$('#saveProductButton').addEventListener('click', async () => {
        const box = $$('#saveProductMessage');
        const id = $$('#manualFgId').value.trim();
        const length = parseFloat($$('#manualLength').value);
        const width = parseFloat($$('#manualWidth').value);
        const height = parseFloat($$('#manualHeight').value);
        const qty = parseInt($$('#manualQty').value, 10);

        if (!id) { say(box, t('saveNeedId'), 'bad'); $$('#manualFgId').focus(); return; }
        if (!(length > 0 && width > 0 && height > 0)) { say(box, t('saveNeedSize'), 'bad'); return; }

        const button = $$('#saveProductButton');
        button.disabled = true;
        try {
          const payload = { fg_id: id, length_mm: length, width_mm: width, height_mm: height, source: 'MANUAL_ENTRY' };
          if (Number.isFinite(qty) && qty > 0) payload.target_qty = qty;
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

      /* ---- New Product tab: save the carton currently typed into the carton block.
         Independent of the product button above. Auto-detect: an id already on the
         list is an update, anything else is a new carton - no separate mode to pick. */
      $$('#saveCartonButton').addEventListener('click', async () => {
        const box = $$('#cartonSourceMessage');
        const cartonId = $$('#manualCartonId').value.trim();
        const cartonLength = parseFloat($$('#manualCartonLength').value);
        const cartonWidth = parseFloat($$('#manualCartonWidth').value);
        const cartonHeight = parseFloat($$('#manualCartonHeight').value);

        if (!cartonId) { say(box, t('cartonSourceNeedId'), 'bad'); $$('#manualCartonId').focus(); return; }
        if (!(cartonLength > 0 && cartonWidth > 0 && cartonHeight > 0)) { say(box, t('cartonSourceNeedSize'), 'bad'); return; }

        const button = $$('#saveCartonButton');
        button.disabled = true;
        try {
          const cartonPayload = { carton_id: cartonId, length_mm: cartonLength, width_mm: cartonWidth, height_mm: cartonHeight, dimension_source: 'MANUAL_MEASUREMENT' };
          const cartonRes = await SavedEntries.save('carton-dimensions', cartonPayload);
          if (cartonRes.ok && cartonRes.body && cartonRes.body.success) {
            say(box, cartonRes.body.replaced ? t('cartonSavedReplaced') : t('cartonSavedOk'), 'ok');
            /* ... existing in-memory cartons/cartonById update, unchanged, omitted here ... */
            $$('#manualCartonId').value = '';
            $$('#manualCartonLength').value = '';
            $$('#manualCartonWidth').value = '';
            $$('#manualCartonHeight').value = '';
          } else {
            const detail = cartonRes.body && Array.isArray(cartonRes.body.details)
              ? cartonRes.body.details.map(d => d.message).join(' ')
              : (cartonRes.body && (cartonRes.body.detail || cartonRes.body.error)) || '';
            say(box, t('cartonSaveFailed') + detail, 'bad');
          }
        } catch (e) {
          say(box, t('saveOffline'), 'bad');
          SavedEntries.probe();
        } finally { button.disabled = false; }
      });
```

The `say` helper used throughout (sets a status box's class and text - `tone` is `'ok'`, `'bad'`, `'warn'`, or falsy):
```js
const say = (el, text, tone) => {
  el.className = 'save-msg' + (tone ? ' ' + tone : '');
  el.textContent = text;
};
```

## What to build

### SECTION 1 - one new method inside the `SavedEntries` IIFE

Add a `saveDimensionToExcel` method right after the existing `save` method (same indentation, same style), and add it to the module's public return statement. It must:
- Take `(type, payload)` where `type` is `'product'` or `'carton'` and `payload` already has `id`, `length_mm`, `width_mm`, `height_mm`, and optionally `name`.
- POST to `/api/save-dimension` using the existing `call` helper, with a 30000ms timeout (this call is slow by design - do not use the 8000ms default).
- Build the JSON body as `{ type, id: payload.id, length_mm: payload.length_mm, width_mm: payload.width_mm, height_mm: payload.height_mm }`, and only include `name: payload.name` in that object when `type === 'carton'` and `payload.name` is a non-empty string.
- Return whatever `call` returns (`{ ok, status, body }`), same as `save` does. Do not catch errors here - let the caller's try/catch handle them, exactly like `save` does today.

Then update the return statement to export it:
```js
return { KINDS, isAvailable, probe, save, saveDimensionToExcel, list, stop, mergeInto };
```

### SECTION 2 - modify both button handlers to also call it, fire-and-forget

For BOTH `#saveProductButton` and `#saveCartonButton`, inside the `if (res.ok && res.body && res.body.success)` / `if (cartonRes.ok && cartonRes.body && cartonRes.body.success)` branch (i.e. only after the existing fast CSV save already succeeded), add a call to `SavedEntries.saveDimensionToExcel(...)` that:
- Is NOT awaited by the `try` block - it must not delay `finally { button.disabled = false }` or block the user from clicking Save again. Fire it as its own `.then()/.catch()` chain (or an inner `async` IIFE you call without `await`), not with `await`.
- Before firing it, append `' ' + t('excelSaveInProgress')` to whatever message `say(box, ...)` already just displayed, so the box reads e.g. "Saved. Saving to Excel…" - use the SAME `box` element and re-call `say` with the combined text and the SAME tone ('ok') that was just used.
- When it resolves successfully (`ok && body && body.success`), re-call `say(box, ...)` replacing the message with the ORIGINAL saved-line (`savedOk`/`savedReplaced` for product, `cartonSavedOk`/`cartonSavedReplaced` for carton - reuse the same ternary already used above) plus `' ' + t('excelSaveOk')`, tone `'ok'`.
- When it resolves with a failure, re-call `say(box, ...)` with the original saved-line plus a space, plus: `t('excelSaveLocked')` if `body && body.status === 503`, otherwise `t('excelSaveFailed') + (body && body.error ? body.error : '')`. Tone stays `'ok'` (the CSV save itself still succeeded - only the Excel part failed, this is not a hard error for the user).
- On a thrown/network error (the `.catch()`), re-call `say(box, ...)` with the original saved-line plus `' ' + t('excelSaveFailed')`, tone `'ok'`.
- Guard against a stale message overwriting a newer one: capture the button's current click as a local generation - e.g. read `box.dataset.saveGen = String((Number(box.dataset.saveGen) || 0) + 1)` right when you start the Excel call, capture that value in a local `const gen = box.dataset.saveGen`, and before each later `say(box, ...)` call from the Excel-result branches, check `if (box.dataset.saveGen !== gen) return;` first so a second click (or the other button) doesn't get its message clobbered by a slow, stale first call resolving late.
- For the product handler, the payload passed to `saveDimensionToExcel` is `{ id, length_mm: length, width_mm: width, height_mm: height }` (reuse the exact same `id`/`length`/`width`/`height` variables already in that handler - do not re-read the DOM). Call it as `SavedEntries.saveDimensionToExcel('product', { id, length_mm: length, width_mm: width, height_mm: height })`.
- For the carton handler, the payload is `{ id: cartonId, name: '', length_mm: cartonLength, width_mm: cartonWidth, height_mm: cartonHeight }` (this form has no carton-name field today, so `name` is always empty - still pass the key so the shape matches). Call it as `SavedEntries.saveDimensionToExcel('carton', { id: cartonId, name: '', length_mm: cartonLength, width_mm: cartonWidth, height_mm: cartonHeight })`.

Return the FULL modified text of both event listener blocks (from `$$('#saveProductButton').addEventListener` through its closing `});`, then the same for `$$('#saveCartonButton')`), with every line of existing logic preserved exactly (including the in-memory cartons/cartonById update block in the carton handler, which was abbreviated above as a comment placeholder only to save space in this prompt - keep the REAL existing code there, do not replace it with a comment).

### SECTION 3 - four new i18n keys, English and Vietnamese

Match the exact object-literal style already used elsewhere in this file (single `Object.assign(this.text.en, { key: 'value', ... })` fragment and the matching `this.text.vi` fragment - just the key:value pairs, not the whole Object.assign wrapper). English wording should be short and match the tone of nearby keys like `savedOk: 'Saved.'`. Vietnamese should be a natural, correct translation, matching the tone of nearby keys like `savedOk:'Đã lưu.'`.

Keys needed:
- `excelSaveInProgress` - shown appended while the Excel write is in flight, e.g. "Saving to Excel…"
- `excelSaveOk` - shown appended on success, e.g. "Saved to Excel and cache updated."
- `excelSaveFailed` - a PREFIX (ends with a trailing space, like the existing `saveFailed:'Could not save. '`) that a short error detail gets appended to, e.g. "Could not save to Excel. "
- `excelSaveLocked` - shown appended when the workbook is open elsewhere (HTTP 503), e.g. "The Excel file is open elsewhere - close it and try again."

## Rules

- Minimal, surgical change. Do not touch anything not described above - not the validation logic, not the CSV save call, not the in-memory carton array update, not any other button or handler in the file.
- Do not invent new HTML elements, IDs, or classes. Use only `say`, `t`, `$$`, `box`, and the variables already named above.
- Do not add a loading spinner, disable other buttons, or otherwise change layout/CSS - text-only status updates via the existing `say` box.
- Plain ASCII punctuation only (regular hyphens and quotes) - no em dashes, curly quotes, or other non-ASCII punctuation anywhere in code or i18n strings, since past output from this same task type had a UTF-8 mis-encoding bug when using an em dash.
