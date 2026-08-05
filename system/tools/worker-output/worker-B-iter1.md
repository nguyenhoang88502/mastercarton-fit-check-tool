const modeSelect = $$('#newProductCartonSourceMode');
const missingFieldsDiv = $$('#cartonSourceMissingFields');
const newFieldsDiv = $$('#cartonSourceNewFields');
const missingSelect = $$('#cartonSourceMissingSelect');
const msgBox = $$('#cartonSourceMessage');

const escapeHTML = (str) => String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

const populateMissingCartonOptions = () => {
  const cartons = toolUi?.data?.cartons || [];
  const missing = cartons.filter(c => !CartonFitEngine.validCarton(c));
  missing.sort((a, b) => a.carton_id.localeCompare(b.carton_id));
  let html = `<option value="" disabled selected>${escapeHTML(t('cartonSourceMissingPrompt'))}</option>`;
  missing.forEach(c => {
    const idEsc = escapeHTML(c.carton_id);
    const nameEsc = c.carton_name ? escapeHTML(c.carton_name) : '';
    const label = nameEsc ? `${idEsc} (${nameEsc})` : idEsc;
    html += `<option value="${idEsc}">${label}</option>`;
  });
  missingSelect.innerHTML = html;
  missingSelect.disabled = missing.length === 0;
};

modeSelect.addEventListener('change', () => {
  const mode = modeSelect.value;
  missingFieldsDiv.hidden = mode !== 'missing';
  newFieldsDiv.hidden = mode !== 'new';
  msgBox.textContent = '';
  if (mode === 'missing') {
    populateMissingCartonOptions();
    missingSelect.value = '';
  }
});

const missingLengthInput = $$('#cartonSourceMissingLength');
const missingWidthInput = $$('#cartonSourceMissingWidth');
const missingHeightInput = $$('#cartonSourceMissingHeight');

missingSelect.addEventListener('change', () => {
  const selectedId = missingSelect.value;
  if (!selectedId) return;
  const carton = (toolUi?.data?.cartons || []).find(c => normalizeId(c.carton_id) === normalizeId(selectedId));
  if (carton) {
    missingLengthInput.value = (Number.isFinite(carton.width) && carton.width > 0) ? carton.width.toString() : '';
    missingWidthInput.value = (Number.isFinite(carton.depth) && carton.depth > 0) ? carton.depth.toString() : '';
    missingHeightInput.value = (Number.isFinite(carton.height) && carton.height > 0) ? carton.height.toString() : '';
  }
});

$$('#cartonSourceMissingSave').addEventListener('click', async () => {
  const cartonId = missingSelect.value;
  if (!cartonId) {
    say(msgBox, t('cartonSourceNeedSelect'), 'bad');
    return;
  }
  const length = parseFloat(missingLengthInput.value);
  const width = parseFloat(missingWidthInput.value);
  const height = parseFloat(missingHeightInput.value);
  if (!(length > 0 && width > 0 && height > 0)) {
    say(msgBox, t('cartonSourceNeedSize'), 'bad');
    return;
  }
  const button = $$('#cartonSourceMissingSave');
  button.disabled = true;
  try {
    const payload = { carton_id: cartonId, length_mm: length, width_mm: width, height_mm: height };
    const res = await SavedEntries.save('carton-dimensions', payload);
    if (res.ok && res.body && res.body.success) {
      say(msgBox, res.body.replaced ? t('savedReplaced') : t('savedOk'), 'ok');
      const carton = (toolUi?.data?.cartons || []).find(c => normalizeId(c.carton_id) === normalizeId(cartonId));
      if (carton) {
        carton.width = length;
        carton.depth = width;
        carton.height = height;
      }
      populateMissingCartonOptions();
    } else {
      const detail = res.body && Array.isArray(res.body.details)
        ? res.body.details.map(d => d.message).join(' ')
        : (res.body && (res.body.detail || res.body.error)) || '';
      say(msgBox, t('saveFailed') + detail, 'bad');
    }
  } catch (e) {
    say(msgBox, t('saveOffline'), 'bad');
    SavedEntries.probe();
  } finally {
    button.disabled = false;
  }
});

const newIdInput = $$('#cartonSourceNewId');
const newNameInput = $$('#cartonSourceNewName');
const newLengthInput = $$('#cartonSourceNewLength');
const newWidthInput = $$('#cartonSourceNewWidth');
const newHeightInput = $$('#cartonSourceNewHeight');

$$('#cartonSourceNewSave').addEventListener('click', async () => {
  const id = newIdInput.value.trim();
  if (!id) {
    say(msgBox, t('cartonSourceNeedId'), 'bad');
    newIdInput.focus();
    return;
  }
  const length = parseFloat(newLengthInput.value);
  const width = parseFloat(newWidthInput.value);
  const height = parseFloat(newHeightInput.value);
  if (!(length > 0 && width > 0 && height > 0)) {
    say(msgBox, t('cartonSourceNeedSize'), 'bad');
    return;
  }
  const name = newNameInput.value.trim();
  const button = $$('#cartonSourceNewSave');
  button.disabled = true;
  try {
    const payload = { carton_id: id, length_mm: length, width_mm: width, height_mm: height };
    const res = await SavedEntries.save('carton-dimensions', payload);
    if (res.ok && res.body && res.body.success) {
      say(msgBox, res.body.replaced ? t('savedReplaced') : t('savedOk'), 'ok');
      const cartons = toolUi?.data?.cartons || [];
      const existing = cartons.find(c => normalizeId(c.carton_id) === normalizeId(id));
      if (existing) {
        existing.width = length;
        existing.depth = width;
        existing.height = height;
        if (name !== undefined) existing.carton_name = name;
      } else {
        cartons.push({
          carton_id: id,
          carton_name: name,
          width: length,
          depth: width,
          height: height,
          defaultQty: 12,
          lifecycle: 'Active',
          dimensionType: 'outside',
          dimensionBasis: 'outside'
        });
      }
      newIdInput.value = '';
      newNameInput.value = '';
      newLengthInput.value = '';
      newWidthInput.value = '';
      newHeightInput.value = '';
      if (modeSelect.value === 'missing') {
        populateMissingCartonOptions();
      }
    } else {
      const detail = res.body && Array.isArray(res.body.details)
        ? res.body.details.map(d => d.message).join(' ')
        : (res.body && (res.body.detail || res.body.error)) || '';
      say(msgBox, t('saveFailed') + detail, 'bad');
    }
  } catch (e) {
    say(msgBox, t('saveOffline'), 'bad');
    SavedEntries.probe();
  } finally {
    button.disabled = false;
  }
});
