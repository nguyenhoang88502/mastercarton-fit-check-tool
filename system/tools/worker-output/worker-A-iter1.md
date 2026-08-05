<div class="form-block carton-source-block save-feature">
  <h3 data-i18n="cartonSourceGroup">If a carton needs measurements</h3>
  <label><span data-i18n="cartonSourceLabel">Carton measurement</span>
    <select id="newProductCartonSourceMode">
      <option value="none" data-i18n="cartonSourceNone" selected>No carton changes needed</option>
      <option value="missing" data-i18n="cartonSourceMissing">Fix a carton with missing measurements</option>
      <option value="new" data-i18n="cartonSourceNew">Add a brand new carton</option>
    </select>
  </label>
  <div id="cartonSourceMissingFields" hidden>
    <label><span data-i18n="cartonSourceMissingSelectLabel">Which carton</span>
      <select id="cartonSourceMissingSelect" disabled>
        <option value="" disabled selected data-i18n="cartonSourceMissingPrompt">â Select a carton â</option>
      </select>
    </label>
    <div class="dimension-row">
      <label><span data-i18n="length">Length</span><input id="cartonSourceMissingLength" type="number" min="0.1" step="0.1" placeholder="mm"><span class="unit">mm</span></label>
      <label><span data-i18n="width">Width</span><input id="cartonSourceMissingWidth" type="number" min="0.1" step="0.1" placeholder="mm"><span class="unit">mm</span></label>
      <label><span data-i18n="height">Height</span><input id="cartonSourceMissingHeight" type="number" min="0.1" step="0.1" placeholder="mm"><span class="unit">mm</span></label>
    </div>
    <button id="cartonSourceMissingSave" type="button" data-i18n="cartonSourceSaveMeasurements">Save these measurements</button>
  </div>
  <div id="cartonSourceNewFields" hidden>
    <label><span data-i18n="cartonSourceNewId">New Carton ID</span><input id="cartonSourceNewId" type="text" maxlength="80" placeholder="e.g., 1008700"></label>
    <label><span data-i18n="cartonSourceNewName">Carton name (optional)</span><input id="cartonSourceNewName" type="text" maxlength="120" placeholder="Optional"></label>
    <div class="dimension-row">
      <label><span data-i18n="length">Length</span><input id="cartonSourceNewLength" type="number" min="0.1" step="0.1" placeholder="mm"><span class="unit">mm</span></label>
      <label><span data-i18n="width">Width</span><input id="cartonSourceNewWidth" type="number" min="0.1" step="0.1" placeholder="mm"><span class="unit">mm</span></label>
      <label><span data-i18n="height">Height</span><input id="cartonSourceNewHeight" type="number" min="0.1" step="0.1" placeholder="mm"><span class="unit">mm</span></label>
    </div>
    <button id="cartonSourceNewSave" type="button" data-i18n="cartonSourceSaveNew">Save new carton</button>
  </div>
  <div id="cartonSourceMessage" class="save-msg" role="status" aria-live="polite"></div>
</div>
