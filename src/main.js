import './styles.css';
import * as pdfjsLib from 'pdfjs-dist';
import pdfjsWorker from 'pdfjs-dist/build/pdf.worker.min.mjs?url';
import { PDFArray, PDFDocument, PDFHexString, PDFName } from 'pdf-lib';
import JSZip from 'jszip';
import * as XLSX from 'xlsx';

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfjsWorker;

const els = {
  modeTabs: [...document.querySelectorAll('.mode-tab')],
  splitMode: document.querySelector('#split-mode'),
  buildMode: document.querySelector('#build-mode'),

  binderInput: document.querySelector('#binder-input'),
  detectTabs: document.querySelector('#detect-tabs'),
  clearMarkers: document.querySelector('#clear-markers'),
  splitDownload: document.querySelector('#split-download'),
  preserveTabPdfs: document.querySelector('#preserve-tab-pdfs'),
  splitStatus: document.querySelector('#split-status'),
  thumbnailList: document.querySelector('#thumbnail-list'),
  previewStage: document.querySelector('#preview-stage'),
  previewLabel: document.querySelector('#preview-label'),
  toggleCurrentMarker: document.querySelector('#toggle-current-marker'),
  splitAfterCurrent: document.querySelector('#split-after-current'),
  nextTab: document.querySelector('#next-tab'),

  currentSectionEditor: document.querySelector('#current-section-editor'),
  currentSectionPages: document.querySelector('#current-section-pages'),
  currentSectionFilenameMode: document.querySelector('#current-section-filename-mode'),
  currentSectionLabel: document.querySelector('#current-section-label'),
  currentSectionDescription: document.querySelector('#current-section-description'),
  currentSectionDate: document.querySelector('#current-section-date'),
  currentSectionFilename: document.querySelector('#current-section-filename'),

  manifestBody: document.querySelector('#manifest-body'),
  importSplitManifest: document.querySelector('#import-split-manifest'),
  exportManifestJson: document.querySelector('#export-manifest-json'),
  exportManifestCsv: document.querySelector('#export-manifest-csv'),
  exportManifestXlsx: document.querySelector('#export-manifest-xlsx'),
  filenameTemplate: document.querySelector('#filename-template'),
  applyFilenameTemplate: document.querySelector('#apply-filename-template'),
  manifestValidationSummary: document.querySelector('#manifest-validation-summary'),
  pasteIndexNames: document.querySelector('#paste-index-names'),
  bulkPasteDialog: document.querySelector('#bulk-paste-dialog'),
  bulkPasteInput: document.querySelector('#bulk-paste-input'),
  bulkPasteSummary: document.querySelector('#bulk-paste-summary'),
  bulkPastePreview: document.querySelector('#bulk-paste-preview'),
  bulkRegenerateFilenames: document.querySelector('#bulk-regenerate-filenames'),
  applyBulkPaste: document.querySelector('#apply-bulk-paste'),
  cancelBulkPaste: document.querySelector('#cancel-bulk-paste'),
  closeBulkPaste: document.querySelector('#close-bulk-paste'),

  buildFiles: document.querySelector('#build-files'),
  buildTabs: document.querySelector('#build-tabs'),
  clearBuildTabs: document.querySelector('#clear-build-tabs'),
  buildIndex: document.querySelector('#build-index'),
  clearBuildIndex: document.querySelector('#clear-build-index'),
  createIndexLinks: document.querySelector('#create-index-links'),
  indexDetectionSummary: document.querySelector('#index-detection-summary'),
  indexLinkBody: document.querySelector('#index-link-body'),
  insertBuildTabs: document.querySelector('#insert-build-tabs'),
  createBuildBookmarks: document.querySelector('#create-build-bookmarks'),
  buildManifest: document.querySelector('#build-manifest'),
  buildOutputName: document.querySelector('#build-output-name'),
  buildDownload: document.querySelector('#build-download'),
  buildStatus: document.querySelector('#build-status'),
  buildBody: document.querySelector('#build-body'),
  exportBuildXlsx: document.querySelector('#export-build-xlsx'),
};

const splitState = {
  file: null,
  bytes: null,
  pdfjsDoc: null,
  selectedPage: 0,
  currentSectionKey: null,
  tabMarkers: new Set(),
  splitAfter: new Set(),
  sectionMeta: new Map(),
  detectedLabels: new Map(),
  renderToken: 0,
  bulkPaste: { rows: [], canApply: false, source: '', issues: [] },
};

const buildState = {
  entries: [],
  tabFiles: [],
  manifest: null,
  draggedEntryId: null,
  indexFile: null,
  indexInfo: null,
  indexLoading: false,
};

els.modeTabs.forEach((tab) => {
  tab.addEventListener('click', () => {
    const mode = tab.dataset.mode;
    els.modeTabs.forEach((item) => item.classList.toggle('active', item === tab));
    els.splitMode.classList.toggle('active', mode === 'split');
    els.buildMode.classList.toggle('active', mode === 'build');
  });
});

els.binderInput.addEventListener('change', async (event) => {
  const [file] = event.target.files;
  if (!file) return;
  await loadBinder(file);
});

els.detectTabs.addEventListener('click', detectTabPages);
els.clearMarkers.addEventListener('click', resetDivisions);
els.toggleCurrentMarker.addEventListener('click', () => toggleTabMarker(splitState.selectedPage));
els.splitAfterCurrent.addEventListener('click', () => toggleSplitAfter(splitState.selectedPage));
els.nextTab.addEventListener('click', goToNextTab);
els.splitDownload.addEventListener('click', splitAndDownload);
els.preserveTabPdfs.addEventListener('change', rebuildManifest);
els.importSplitManifest.addEventListener('change', importSplitManifest);
els.exportManifestJson.addEventListener('click', exportManifestJson);
els.exportManifestCsv.addEventListener('click', exportManifestCsv);
els.exportManifestXlsx.addEventListener('click', exportManifestXlsx);
els.filenameTemplate.addEventListener('input', () => {
  renderManifestTable();
  updateCurrentSectionEditor();
});
els.applyFilenameTemplate.addEventListener('click', () => {
  splitState.sectionMeta.forEach((meta) => {
    delete meta.filename;
    meta.filenameLocked = false;
  });
  rebuildManifest();
  setSplitStatus('Filename template reapplied to every section. Manually edited filenames were reset to automatic naming.', 'success');
});

els.currentSectionLabel.addEventListener('input', () => updateCurrentSectionField('label', els.currentSectionLabel.value));
els.currentSectionDescription.addEventListener('input', () => updateCurrentSectionField('description', els.currentSectionDescription.value));
els.currentSectionDate.addEventListener('input', () => updateCurrentSectionField('date', els.currentSectionDate.value));
els.currentSectionFilename.addEventListener('input', () => updateCurrentSectionField('filename', els.currentSectionFilename.value, true));

els.pasteIndexNames.addEventListener('click', openBulkPasteDialog);
els.bulkPasteInput.addEventListener('paste', handleBulkPasteEvent);
els.bulkPasteInput.addEventListener('input', () => analyzeBulkPaste('', els.bulkPasteInput.value));
els.bulkRegenerateFilenames.addEventListener('change', renderBulkPastePreview);
els.applyBulkPaste.addEventListener('click', applyBulkPasteEntries);
els.cancelBulkPaste.addEventListener('click', closeBulkPasteDialog);
els.closeBulkPaste.addEventListener('click', closeBulkPasteDialog);
els.bulkPasteDialog.addEventListener('cancel', (event) => {
  event.preventDefault();
  closeBulkPasteDialog();
});

els.buildFiles.addEventListener('change', (event) => {
  addBuildFiles([...event.target.files]);
  event.target.value = '';
});
els.buildTabs.addEventListener('change', (event) => {
  addBuildTabFiles([...event.target.files]);
  event.target.value = '';
});
els.clearBuildTabs.addEventListener('click', () => {
  buildState.tabFiles = [];
  renderBuildTable();
  updateBuildStatusSummary();
});
els.buildIndex.addEventListener('change', async (event) => {
  const [file] = event.target.files;
  if (file) await loadBuildIndex(file);
  event.target.value = '';
});
els.clearBuildIndex.addEventListener('click', () => {
  buildState.indexFile = null;
  buildState.indexInfo = null;
  buildState.indexLoading = false;
  renderIndexPanel();
  renderBuildTable();
  updateBuildStatusSummary();
});
els.insertBuildTabs.addEventListener('change', () => {
  renderBuildTable();
  updateBuildStatusSummary();
});
els.createBuildBookmarks.addEventListener('change', updateBuildStatusSummary);
els.createIndexLinks.addEventListener('change', () => {
  renderIndexPanel();
  renderBuildTable();
  updateBuildStatusSummary();
});
els.buildManifest.addEventListener('change', importBuildManifest);
els.buildDownload.addEventListener('click', buildAndDownload);
els.exportBuildXlsx.addEventListener('click', exportBuildXlsx);


const SAFE_FILENAME_MAX_LENGTH = 120;
const WINDOWS_RESERVED_BASENAMES = /^(?:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])$/i;

function openBulkPasteDialog() {
  const sections = getSections();
  if (!sections.length) {
    setSplitStatus('Create at least one output section before pasting index entries.', 'error');
    return;
  }

  splitState.bulkPaste = { rows: [], canApply: false, source: '', issues: [] };
  els.bulkPasteInput.value = '';
  renderBulkPastePreview();

  if (typeof els.bulkPasteDialog.showModal === 'function') {
    els.bulkPasteDialog.showModal();
  } else {
    els.bulkPasteDialog.setAttribute('open', '');
  }
  setTimeout(() => els.bulkPasteInput.focus(), 0);
}

function closeBulkPasteDialog() {
  if (typeof els.bulkPasteDialog.close === 'function' && els.bulkPasteDialog.open) {
    els.bulkPasteDialog.close();
  } else {
    els.bulkPasteDialog.removeAttribute('open');
  }
}

function handleBulkPasteEvent(event) {
  const html = event.clipboardData?.getData('text/html') || '';
  const text = event.clipboardData?.getData('text/plain') || '';
  if (!html && !text) return;

  event.preventDefault();
  els.bulkPasteInput.value = text;
  analyzeBulkPaste(html, text);
}

function analyzeBulkPaste(html, text) {
  const htmlRows = html ? parseWordTableHtml(html) : [];
  const rows = htmlRows.length ? htmlRows : parsePlainIndexText(text);
  const source = htmlRows.length ? 'Word/HTML table' : rows.length ? 'plain text' : '';
  const analysis = validateBulkPasteRows(rows, getSections());

  splitState.bulkPaste = {
    rows: analysis.rows,
    canApply: analysis.canApply,
    source,
    issues: analysis.issues,
  };
  renderBulkPastePreview();
}

function parseWordTableHtml(html) {
  try {
    const doc = new DOMParser().parseFromString(html, 'text/html');
    const candidates = [...doc.querySelectorAll('table')].map((table) => {
      const rows = [];
      [...table.querySelectorAll('tr')].forEach((tr) => {
        const cells = [...tr.children]
          .filter((node) => ['TD', 'TH'].includes(node.tagName))
          .map(extractPastedCellText);
        const parsed = rowFromCells(cells);
        if (parsed) rows.push(parsed);
      });
      return rows;
    });
    candidates.sort((a, b) => b.length - a.length);
    return candidates[0] || [];
  } catch (error) {
    console.warn('Could not parse clipboard HTML table; falling back to plain text.', error);
    return [];
  }
}

function extractPastedCellText(cell) {
  const chunks = [...cell.childNodes]
    .map((node) => normalizePastedCell(node.textContent))
    .filter(Boolean);
  return normalizePastedCell(chunks.join(' '));
}

function rowFromCells(cells) {
  if (cells.length < 2) return null;
  const numberIndex = cells.findIndex((cell) => parsePastedRowNumber(cell) != null);
  if (numberIndex < 0) return null;

  const number = parsePastedRowNumber(cells[numberIndex]);
  const trailing = cells.slice(numberIndex + 1);
  if (!trailing.length) return null;

  const description = normalizePastedCell(trailing[0]);
  const date = normalizePastedCell(trailing.length > 1 ? trailing[trailing.length - 1] : '');
  if (!description || normalizeHeader(description) === 'description') return null;

  return { number, description, date };
}

function parsePlainIndexText(text) {
  const source = String(text || '').replace(/\u00a0/g, ' ');
  if (!source.trim()) return [];

  const lines = source.split(/\r?\n/).map((line) => line.trim()).filter(Boolean);

  // First choice for plain clipboard text from Word tables: tab-separated cells.
  const tabRows = [];
  lines.forEach((line) => {
    if (!line.includes('\t')) return;
    const parsed = rowFromCells(line.split(/\t+/).map(normalizePastedCell));
    if (parsed) tabRows.push(parsed);
  });
  if (tabRows.length) return tabRows;

  // Fallback: group lines under explicit numbered entries. This also handles
  // clipboard text where Word emits each cell on its own line.
  const groups = [];
  let current = null;
  lines.forEach((line) => {
    const standalone = line.match(/^\s*(\d{1,3})\s*[.)]?\s*$/);
    const inline = line.match(/^\s*(\d{1,3})\s*[.)]\s+(.+)$/);

    if (standalone || inline) {
      if (current) groups.push(current);
      current = {
        number: Number((standalone || inline)[1]),
        parts: inline ? [inline[2]] : [],
      };
      return;
    }

    if (current) current.parts.push(line);
  });
  if (current) groups.push(current);

  return groups.map((group) => {
    const parts = group.parts.map(normalizePastedCell).filter(Boolean);
    let date = '';
    if (parts.length > 1 && looksLikePastedIndexDate(parts[parts.length - 1])) date = parts.pop();
    return {
      number: group.number,
      description: normalizePastedCell(parts.join(' ')),
      date,
    };
  }).filter((row) => row.description);
}

function parsePastedRowNumber(value) {
  const match = String(value || '').trim().match(/^(\d{1,3})\s*[.)]?$/);
  return match ? Number(match[1]) : null;
}

function looksLikePastedIndexDate(value) {
  const text = normalizePastedCell(value);
  return /^(?:\d{1,2}[/-]\d{1,2}[/-]\d{2,4}|\d{4}[/-]\d{1,2}[/-]\d{1,2}|\d{4}\s*[-–]\s*\d{4}|\d{4})$/.test(text);
}

function normalizePastedCell(value) {
  return String(value || '')
    .replace(/\u00a0/g, ' ')
    .replace(/[\r\n\t]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function validateBulkPasteRows(rows, sections) {
  const cleaned = rows.map((row) => ({
    number: Number(row.number),
    description: normalizePastedCell(row.description),
    date: normalizePastedCell(row.date),
  })).filter((row) => Number.isInteger(row.number) && row.number > 0);

  const issues = [];
  const seen = new Map();
  cleaned.forEach((row) => seen.set(row.number, (seen.get(row.number) || 0) + 1));
  const duplicates = [...seen.entries()].filter(([, count]) => count > 1).map(([number]) => number);
  if (duplicates.length) issues.push(`Duplicate row number${duplicates.length === 1 ? '' : 's'}: ${duplicates.join(', ')}.`);

  const expected = sections.map((section) => section.order);
  const actual = new Set(cleaned.map((row) => row.number));
  const missing = expected.filter((number) => !actual.has(number));
  const extra = [...actual].filter((number) => number < 1 || number > sections.length).sort((a, b) => a - b);

  if (cleaned.length !== sections.length) {
    issues.push(`Detected ${cleaned.length} entr${cleaned.length === 1 ? 'y' : 'ies'}, but there are ${sections.length} output sections.`);
  }
  if (missing.length) issues.push(`Missing position${missing.length === 1 ? '' : 's'}: ${missing.join(', ')}.`);
  if (extra.length) issues.push(`Unexpected position${extra.length === 1 ? '' : 's'}: ${extra.join(', ')}.`);

  const blankDescriptions = cleaned.filter((row) => !row.description).map((row) => row.number);
  if (blankDescriptions.length) issues.push(`Blank description at position${blankDescriptions.length === 1 ? '' : 's'}: ${blankDescriptions.join(', ')}.`);

  const canApply = cleaned.length === sections.length
    && !duplicates.length
    && !missing.length
    && !extra.length
    && !blankDescriptions.length;

  return { rows: cleaned.sort((a, b) => a.number - b.number), issues, canApply };
}

function renderBulkPastePreview() {
  const { rows, canApply, source, issues } = splitState.bulkPaste;
  const sections = getSections();
  els.applyBulkPaste.disabled = !canApply;

  if (!rows.length) {
    els.bulkPasteSummary.textContent = 'Paste an index table to preview detected entries.';
    els.bulkPasteSummary.className = 'bulk-paste-summary';
    els.bulkPastePreview.innerHTML = '<tr><td colspan="5" class="empty-cell">Nothing pasted yet.</td></tr>';
    return;
  }

  if (canApply) {
    els.bulkPasteSummary.textContent = `${rows.length} entries detected from ${source || 'clipboard text'}. Numbering matches all ${sections.length} current sections.`;
    els.bulkPasteSummary.className = 'bulk-paste-summary success';
  } else {
    els.bulkPasteSummary.textContent = issues.join(' ');
    els.bulkPasteSummary.className = 'bulk-paste-summary error';
  }

  els.bulkPastePreview.innerHTML = '';
  rows.forEach((row) => {
    const section = sections[row.number - 1];
    const filenameInfo = section ? previewFilenameForPastedRow(row, section) : null;
    const tr = document.createElement('tr');
    if (!section) tr.classList.add('row-has-error');

    const adjustmentText = filenameInfo?.adjusted ? 'Adjusted' : 'Ready';
    const statusClass = filenameInfo?.adjusted ? 'warning' : section ? 'ok' : 'error';
    const statusTitle = filenameInfo?.reasons?.join('\n') || (section ? 'Filename already cross-platform safe.' : 'No matching output section.');

    tr.innerHTML = `
      <td>${row.number}</td>
      <td>${escapeHtml(row.description)}</td>
      <td>${escapeHtml(row.date || '')}</td>
      <td class="bulk-filename-cell">${escapeHtml(filenameInfo?.safe || '—')}</td>
      <td><span class="row-status ${statusClass}" title="${escapeAttr(statusTitle)}">${section ? (filenameInfo?.adjusted ? '!' : '✓') : '!'}</span><span class="bulk-status-text">${adjustmentText}</span></td>
    `;
    els.bulkPastePreview.append(tr);
  });
}

function previewFilenameForPastedRow(row, section) {
  if (!els.bulkRegenerateFilenames.checked) {
    const safe = sanitizePdfFilename(section.filename);
    return { safe, adjusted: safe !== section.filename, reasons: filenameAdjustmentReasons(section.filename, safe) };
  }

  const context = {
    order: row.number,
    label: section.label,
    description: row.description,
    date: row.date,
    startPage: section.startPage,
    endPage: section.endPage,
    original: stripExtension(splitState.file?.name || 'Source'),
  };
  const raw = renderFilenameTemplateRaw(els.filenameTemplate.value, context);
  const safe = sanitizePdfFilename(raw);
  return { safe, adjusted: safe !== raw, reasons: filenameAdjustmentReasons(raw, safe) };
}

function filenameAdjustmentReasons(rawValue, safeValue) {
  const raw = String(rawValue || '');
  const reasons = [];
  if (/[<>:"/\\|?*\x00-\x1F]/.test(raw)) reasons.push('Removed or replaced characters that are unsafe on Windows/macOS.');
  if (/[. ](?:\.pdf)?$/i.test(raw.trimEnd()) || /[. ]$/.test(stripExtension(raw))) reasons.push('Removed a trailing period or space.');
  const rawBase = stripExtension(raw).trim().replace(/[. ]+$/g, '');
  if (WINDOWS_RESERVED_BASENAMES.test(rawBase)) reasons.push('Adjusted a Windows-reserved filename.');
  if ([...raw].length > SAFE_FILENAME_MAX_LENGTH) reasons.push(`Shortened to the ${SAFE_FILENAME_MAX_LENGTH}-character filename limit used by this app.`);
  if (!reasons.length && raw !== safeValue) reasons.push('Normalized the filename for cross-platform compatibility.');
  return reasons;
}

function applyBulkPasteEntries() {
  const { rows, canApply } = splitState.bulkPaste;
  if (!canApply) return;

  const sections = getSections();
  const regenerate = els.bulkRegenerateFilenames.checked;
  rows.forEach((row) => {
    const section = sections[row.number - 1];
    if (!section) return;
    const meta = getOrCreateSectionMeta(section.key);
    meta.description = row.description;
    meta.date = row.date;
    if (regenerate) {
      delete meta.filename;
      meta.filenameLocked = false;
    }
  });

  rebuildManifest();
  closeBulkPasteDialog();
  const filenamePhrase = regenerate ? ' Safe filenames were regenerated from the current template.' : ' Existing filenames were preserved.';
  setSplitStatus(`Applied ${rows.length} pasted index entr${rows.length === 1 ? 'y' : 'ies'} to the manifest.${filenamePhrase}`, 'success');
}

async function loadBinder(file) {
  try {
    setSplitStatus(`Opening ${file.name}…`);
    resetSplitState();

    const arrayBuffer = await file.arrayBuffer();
    splitState.file = file;
    splitState.bytes = new Uint8Array(arrayBuffer);
    splitState.pdfjsDoc = await pdfjsLib.getDocument({ data: splitState.bytes.slice() }).promise;
    splitState.selectedPage = 0;

    const count = splitState.pdfjsDoc.numPages;
    setSplitStatus(`${file.name} — ${count} page${count === 1 ? '' : 's'}.`);
    setSplitControls(true);
    rebuildManifest();
    await renderThumbnails();
    await renderPreview(0);
  } catch (error) {
    console.error(error);
    resetSplitState();
    setSplitControls(false);
    setSplitStatus(`Could not open that PDF: ${friendlyError(error)}`, 'error');
  }
}

function resetSplitState() {
  splitState.renderToken += 1;
  splitState.file = null;
  splitState.bytes = null;
  splitState.pdfjsDoc = null;
  splitState.selectedPage = 0;
  splitState.currentSectionKey = null;
  splitState.tabMarkers = new Set();
  splitState.splitAfter = new Set();
  splitState.sectionMeta = new Map();
  splitState.detectedLabels = new Map();
  splitState.bulkPaste = { rows: [], canApply: false, source: '', issues: [] };
  if (els.bulkPasteDialog?.open) closeBulkPasteDialog();

  els.thumbnailList.innerHTML = 'No PDF loaded.';
  els.thumbnailList.className = 'thumbnail-list empty-state';
  els.previewStage.innerHTML = 'Open a PDF to preview pages.';
  els.previewStage.className = 'preview-stage empty-state';
  els.previewLabel.textContent = 'No page selected.';
  els.manifestBody.innerHTML = '<tr><td colspan="8" class="empty-cell">No divisions yet.</td></tr>';
  els.manifestValidationSummary.textContent = 'No sections yet.';
  els.manifestValidationSummary.className = 'validation-summary';
  hideCurrentSectionEditor();
}

function setSplitControls(enabled) {
  els.detectTabs.disabled = !enabled;
  els.clearMarkers.disabled = !enabled;
  els.splitDownload.disabled = !enabled;
  els.toggleCurrentMarker.disabled = !enabled;
  els.splitAfterCurrent.disabled = !enabled;
  els.nextTab.disabled = !enabled;
  els.exportManifestJson.disabled = !enabled;
  els.exportManifestCsv.disabled = !enabled;
  els.exportManifestXlsx.disabled = !enabled;
  els.applyFilenameTemplate.disabled = !enabled;
  els.pasteIndexNames.disabled = !enabled;
}

async function renderThumbnails() {
  const doc = splitState.pdfjsDoc;
  if (!doc) return;

  const token = ++splitState.renderToken;
  els.thumbnailList.className = 'thumbnail-list';
  els.thumbnailList.innerHTML = '';

  for (let index = 0; index < doc.numPages; index += 1) {
    if (token !== splitState.renderToken) return;

    const card = document.createElement('button');
    card.type = 'button';
    card.className = 'thumbnail-card';
    card.dataset.pageIndex = String(index);
    card.innerHTML = `
      <span class="thumb-page-num">${index + 1}</span>
      <canvas class="thumb-canvas" aria-label="Page ${index + 1} thumbnail"></canvas>
      <span class="thumb-meta">
        <span class="thumb-label"></span>
        <span class="thumb-hint">Click to preview</span>
      </span>
    `;
    card.addEventListener('click', async () => {
      splitState.selectedPage = index;
      updateThumbnailClasses();
      await renderPreview(index);
    });
    els.thumbnailList.append(card);

    try {
      const page = await doc.getPage(index + 1);
      const naturalViewport = page.getViewport({ scale: 1 });
      const targetWidth = 164;
      const viewport = page.getViewport({ scale: targetWidth / naturalViewport.width });
      const canvas = card.querySelector('canvas');
      const ctx = canvas.getContext('2d', { alpha: false });
      canvas.width = Math.ceil(viewport.width);
      canvas.height = Math.ceil(viewport.height);
      await page.render({ canvasContext: ctx, viewport }).promise;
    } catch (error) {
      console.warn(`Thumbnail render failed for page ${index + 1}`, error);
    }
  }

  updateThumbnailClasses();
}

async function renderPreview(pageIndex) {
  const doc = splitState.pdfjsDoc;
  if (!doc) return;

  const page = await doc.getPage(pageIndex + 1);
  const naturalViewport = page.getViewport({ scale: 1 });
  const stageWidth = Math.max(320, els.previewStage.clientWidth - 50);
  const scale = Math.min(1.5, stageWidth / naturalViewport.width);
  const viewport = page.getViewport({ scale });

  const canvas = document.createElement('canvas');
  const ctx = canvas.getContext('2d', { alpha: false });
  canvas.width = Math.ceil(viewport.width);
  canvas.height = Math.ceil(viewport.height);

  els.previewStage.className = 'preview-stage';
  els.previewStage.innerHTML = '';
  els.previewStage.append(canvas);
  els.previewLabel.textContent = `Page ${pageIndex + 1} of ${doc.numPages}`;
  updateMarkerButtons();
  updateCurrentSectionEditor();

  await page.render({ canvasContext: ctx, viewport }).promise;
}

function toggleTabMarker(pageIndex) {
  if (!splitState.pdfjsDoc) return;

  if (splitState.tabMarkers.has(pageIndex)) {
    splitState.tabMarkers.delete(pageIndex);
  } else {
    splitState.tabMarkers.add(pageIndex);
    splitState.splitAfter.delete(pageIndex);
  }

  rebuildManifest();
  updateThumbnailClasses();
  updateMarkerButtons();
}

function toggleSplitAfter(pageIndex) {
  if (!splitState.pdfjsDoc) return;

  if (splitState.splitAfter.has(pageIndex)) {
    splitState.splitAfter.delete(pageIndex);
  } else {
    splitState.splitAfter.add(pageIndex);
    splitState.tabMarkers.delete(pageIndex);
  }

  rebuildManifest();
  updateThumbnailClasses();
  updateMarkerButtons();
}

async function goToNextTab() {
  if (!splitState.pdfjsDoc) return;

  const nextTabIndex = [...splitState.tabMarkers]
    .filter((index) => index > splitState.selectedPage)
    .sort((a, b) => a - b)[0];
  if (nextTabIndex == null) return;

  splitState.selectedPage = nextTabIndex;
  updateThumbnailClasses();

  const card = els.thumbnailList.querySelector(`[data-page-index="${nextTabIndex}"]`);
  card?.scrollIntoView({ behavior: 'smooth', block: 'center' });

  await renderPreview(nextTabIndex);
}

function updateMarkerButtons() {
  if (!splitState.pdfjsDoc) return;
  const pageIndex = splitState.selectedPage;
  const isLastPage = pageIndex === splitState.pdfjsDoc.numPages - 1;
  const nextTabIndex = [...splitState.tabMarkers]
    .filter((index) => index > pageIndex)
    .sort((a, b) => a - b)[0];

  els.nextTab.disabled = nextTabIndex == null;
  els.nextTab.textContent = nextTabIndex == null ? 'No later tab' : `Next tab → ${nextTabIndex + 1}`;

  els.toggleCurrentMarker.disabled = false;
  els.toggleCurrentMarker.textContent = splitState.tabMarkers.has(pageIndex)
    ? 'Unmark tab divider'
    : 'Mark as tab divider';

  els.splitAfterCurrent.disabled = isLastPage;
  els.splitAfterCurrent.textContent = splitState.splitAfter.has(pageIndex)
    ? 'Remove split after page'
    : isLastPage
      ? 'Last page'
      : 'Split after page';
}

function updateThumbnailClasses() {
  const cards = [...els.thumbnailList.querySelectorAll('.thumbnail-card')];
  cards.forEach((card) => {
    const index = Number(card.dataset.pageIndex);
    const isTab = splitState.tabMarkers.has(index);
    const isCut = splitState.splitAfter.has(index);
    card.classList.toggle('selected', index === splitState.selectedPage);
    card.classList.toggle('marker', isTab);
    card.classList.toggle('cut-marker', isCut);

    const label = card.querySelector('.thumb-label');
    const detected = splitState.detectedLabels.get(index);
    if (isTab) {
      label.textContent = detected ? `${detected} — EXCLUDED` : 'TAB DIVIDER — EXCLUDED';
    } else if (isCut) {
      label.textContent = 'SPLIT AFTER — PAGE KEPT';
    } else {
      label.textContent = detected ? `Suggested: ${detected}` : '';
    }
  });
}

async function detectTabPages() {
  const doc = splitState.pdfjsDoc;
  if (!doc) return;

  const pattern = /\bTAB\s*(?:NO\.?\s*)?[:#\-–—]?\s*([A-Z]|\d{1,3})\b/i;
  let found = 0;
  els.detectTabs.disabled = true;

  try {
    for (let index = 0; index < doc.numPages; index += 1) {
      setSplitStatus(`Scanning page ${index + 1} of ${doc.numPages} for tab labels…`);
      const page = await doc.getPage(index + 1);
      const content = await page.getTextContent();
      const text = content.items.map((item) => item.str).join(' ').replace(/\s+/g, ' ').trim();
      const match = text.slice(0, 1800).match(pattern);
      if (match) {
        const label = `TAB ${String(match[1]).toUpperCase()}`;
        splitState.detectedLabels.set(index, label);
        if (!splitState.tabMarkers.has(index)) {
          splitState.tabMarkers.add(index);
          splitState.splitAfter.delete(index);
          found += 1;
        }
      }
    }

    rebuildManifest();
    updateThumbnailClasses();
    updateMarkerButtons();
    setSplitStatus(
      found
        ? `Detected ${found} new tab page${found === 1 ? '' : 's'}. Review the markers before splitting.`
        : 'No new tab labels were detected. Add normal split cuts or mark tab-divider pages manually as needed.',
      found ? 'success' : '',
    );
  } catch (error) {
    console.error(error);
    setSplitStatus(`Tab detection failed: ${friendlyError(error)}`, 'error');
  } finally {
    els.detectTabs.disabled = false;
  }
}

function resetDivisions() {
  if (!splitState.pdfjsDoc) return;
  splitState.tabMarkers = new Set();
  splitState.splitAfter = new Set();
  splitState.sectionMeta = new Map();
  splitState.detectedLabels = new Map();
  rebuildManifest();
  updateThumbnailClasses();
  updateMarkerButtons();
  setSplitStatus('Divisions reset. Use normal split cuts to keep every page, or mark tab-divider pages to exclude them.');
}

function boundaryKey(boundary) {
  return boundary.type === 'start' ? 'start' : `${boundary.type}:${boundary.pageIndex}`;
}

function boundaryLabel(boundary) {
  if (boundary.type === 'tab') return `Tab page ${boundary.pageIndex + 1} (excluded)`;
  if (boundary.type === 'cut') return `After page ${boundary.pageIndex + 1}`;
  return 'Start of PDF';
}

function getSections() {
  if (!splitState.pdfjsDoc) return [];

  const pageCount = splitState.pdfjsDoc.numPages;
  const tabPages = [...splitState.tabMarkers].sort((a, b) => a - b);
  const cutPages = [...splitState.splitAfter].sort((a, b) => a - b);
  const firstTabIndex = tabPages.length ? tabPages[0] : null;
  const hasLeadingCut = firstTabIndex != null && cutPages.some((pageIndex) => pageIndex < firstTabIndex);

  // Preserve the binder behavior established in testing: if the first structural
  // marker is a tab divider, unmarked front matter before it is omitted. A normal
  // cut before that first tab signals that leading pages should be preserved.
  let outputActive = firstTabIndex == null || hasLeadingCut;
  let startBoundary = { type: 'start', pageIndex: null };
  let currentPages = [];
  const rawSections = [];

  const finalize = () => {
    if (!currentPages.length) return;
    rawSections.push({ pageIndexes: currentPages, boundary: startBoundary });
    currentPages = [];
  };

  for (let pageIndex = 0; pageIndex < pageCount; pageIndex += 1) {
    if (splitState.tabMarkers.has(pageIndex)) {
      if (outputActive) finalize();
      outputActive = true;
      startBoundary = { type: 'tab', pageIndex };
      continue;
    }

    if (!outputActive) continue;

    currentPages.push(pageIndex);
    if (splitState.splitAfter.has(pageIndex) && pageIndex < pageCount - 1) {
      finalize();
      startBoundary = { type: 'cut', pageIndex };
    }
  }
  finalize();

  return rawSections.map((raw, index) => {
    const order = index + 1;
    const key = boundaryKey(raw.boundary);
    const existing = splitState.sectionMeta.get(key) || {};
    const tabIndex = raw.boundary.type === 'tab' ? raw.boundary.pageIndex : null;
    const detectedLabel = tabIndex == null ? null : splitState.detectedLabels.get(tabIndex);
    const defaultLabel = detectedLabel ?? (tabIndex == null ? `SECTION ${order}` : `TAB ${order}`);
    const defaultDescription = detectedLabel ?? `Section ${order}`;
    const label = existing.label ?? defaultLabel;
    const description = existing.description ?? defaultDescription;
    const date = existing.date ?? '';
    const startIndex = raw.pageIndexes[0];
    const endIndex = raw.pageIndexes[raw.pageIndexes.length - 1];

    const filenameContext = {
      order,
      label,
      description,
      date,
      startPage: startIndex + 1,
      endPage: endIndex + 1,
      original: stripExtension(splitState.file?.name || 'Source'),
    };
    const filename = existing.filenameLocked
      ? (existing.filename ?? '')
      : renderFilenameTemplate(els.filenameTemplate.value, filenameContext);

    return {
      order,
      key,
      boundary: raw.boundary,
      boundaryText: boundaryLabel(raw.boundary),
      tabIndex,
      tabPage: tabIndex == null ? null : tabIndex + 1,
      splitAfterPage: raw.boundary.type === 'cut' ? raw.boundary.pageIndex + 1 : null,
      startIndex,
      endIndex,
      startPage: startIndex + 1,
      endPage: endIndex + 1,
      pageIndexes: raw.pageIndexes,
      pageCount: raw.pageIndexes.length,
      label,
      description,
      date,
      filename,
      filenameLocked: Boolean(existing.filenameLocked),
    };
  });
}

function getTabAssets() {
  const sectionsByTab = new Map(
    getSections()
      .filter((section) => section.tabIndex != null)
      .map((section) => [section.tabIndex, section]),
  );

  return [...splitState.tabMarkers]
    .sort((a, b) => a - b)
    .map((tabIndex, index) => {
      const section = sectionsByTab.get(tabIndex);
      const label = section?.label ?? splitState.detectedLabels.get(tabIndex) ?? `TAB ${index + 1}`;
      return {
        order: index + 1,
        tabIndex,
        tabPage: tabIndex + 1,
        label,
        filename: sanitizePdfFilename(label || `TAB ${index + 1}`),
      };
    });
}

function getSectionForPage(pageIndex) {
  const sections = getSections();
  return sections.find((section) => section.tabIndex === pageIndex || section.pageIndexes.includes(pageIndex)) || null;
}

function getSectionByKey(key) {
  return getSections().find((section) => section.key === key) || null;
}

function getOrCreateSectionMeta(key) {
  const meta = splitState.sectionMeta.get(key) || {};
  splitState.sectionMeta.set(key, meta);
  return meta;
}

function updateSectionMeta(key, field, value, lockFilename = false) {
  if (!key) return;
  const meta = getOrCreateSectionMeta(key);
  meta[field] = value;
  if (field === 'filename' || lockFilename) meta.filenameLocked = true;
}

function rebuildManifest() {
  renderManifestTable();
  updateCurrentSectionEditor();
}

function renderManifestTable() {
  const sections = getSections();
  const validation = validateSections(sections);
  const hasSections = sections.length > 0;
  const hasBlockingErrors = validation.errorCount > 0;

  els.splitDownload.disabled = !hasSections || hasBlockingErrors;
  els.exportManifestJson.disabled = !hasSections;
  els.exportManifestCsv.disabled = !hasSections;
  els.exportManifestXlsx.disabled = !hasSections;
  els.applyFilenameTemplate.disabled = !hasSections;
  els.pasteIndexNames.disabled = !hasSections;

  updateManifestValidationSummary(validation, sections.length);

  if (!sections.length) {
    els.manifestBody.innerHTML = '<tr><td colspan="8" class="empty-cell">No output sections. Add a normal split cut, detect tabs, or mark divider pages manually.</td></tr>';
    return;
  }

  els.manifestBody.innerHTML = '';
  sections.forEach((section) => {
    const issues = validation.byKey.get(section.key) || { errors: [], warnings: [] };
    const statusClass = issues.errors.length ? 'error' : issues.warnings.length ? 'warning' : 'ok';
    const statusText = issues.errors.length ? '!' : issues.warnings.length ? '!' : '✓';
    const statusTitle = [...issues.errors, ...issues.warnings].join('\n') || 'Ready';

    const row = document.createElement('tr');
    row.dataset.sectionKey = section.key;
    if (issues.errors.length) row.classList.add('row-has-error');
    if (issues.warnings.length) row.classList.add('row-has-warning');
    row.innerHTML = `
      <td><span class="row-status ${statusClass}" title="${escapeAttr(statusTitle)}">${statusText}</span></td>
      <td>${section.order}</td>
      <td>${escapeHtml(section.boundaryText)}</td>
      <td>${section.startPage}–${section.endPage} <span class="muted">(${section.pageCount})</span></td>
      <td><input data-field="label" value="${escapeAttr(section.label)}" /></td>
      <td><input data-field="description" value="${escapeAttr(section.description)}" /></td>
      <td><input data-field="date" value="${escapeAttr(section.date)}" /></td>
      <td><input data-field="filename" value="${escapeAttr(section.filename)}" title="${section.filenameLocked ? 'Manual filename' : 'Generated from filename template'}" /></td>
    `;

    row.querySelectorAll('input').forEach((input) => {
      input.addEventListener('input', () => {
        const field = input.dataset.field;
        updateSectionMeta(section.key, field, input.value, field === 'filename');

        const refreshed = getSectionByKey(section.key);
        if (refreshed && field !== 'filename' && !refreshed.filenameLocked) {
          const filenameInput = row.querySelector('[data-field="filename"]');
          filenameInput.value = refreshed.filename;
          filenameInput.title = 'Generated from filename template';
        } else if (field === 'filename') {
          input.title = 'Manual filename';
        }

        refreshManifestValidationRows();
        if (splitState.currentSectionKey === section.key) syncCurrentSectionEditorValues();
      });
    });

    els.manifestBody.append(row);
  });
}

function refreshManifestValidationRows() {
  const sections = getSections();
  const validation = validateSections(sections);
  updateManifestValidationSummary(validation, sections.length);
  els.splitDownload.disabled = !sections.length || validation.errorCount > 0;

  sections.forEach((section) => {
    const row = [...els.manifestBody.querySelectorAll('tr')].find((candidate) => candidate.dataset.sectionKey === section.key);
    if (!row) return;
    const issues = validation.byKey.get(section.key) || { errors: [], warnings: [] };
    const status = row.querySelector('.row-status');
    if (!status) return;
    const statusClass = issues.errors.length ? 'error' : issues.warnings.length ? 'warning' : 'ok';
    status.className = `row-status ${statusClass}`;
    status.textContent = issues.errors.length || issues.warnings.length ? '!' : '✓';
    status.title = [...issues.errors, ...issues.warnings].join('\n') || 'Ready';
    row.classList.toggle('row-has-error', issues.errors.length > 0);
    row.classList.toggle('row-has-warning', issues.warnings.length > 0);
  });
}

function updateManifestValidationSummary(validation, sectionCount) {
  if (!sectionCount) {
    els.manifestValidationSummary.textContent = 'No sections yet.';
    els.manifestValidationSummary.className = 'validation-summary';
    return;
  }

  if (validation.errorCount) {
    els.manifestValidationSummary.textContent = `${validation.errorCount} blocking issue${validation.errorCount === 1 ? '' : 's'}; fix before splitting.`;
    els.manifestValidationSummary.className = 'validation-summary error';
    return;
  }

  if (validation.warningCount) {
    els.manifestValidationSummary.textContent = `${sectionCount} section${sectionCount === 1 ? '' : 's'}; ${validation.warningCount} warning${validation.warningCount === 1 ? '' : 's'}.`;
    els.manifestValidationSummary.className = 'validation-summary warning';
    return;
  }

  els.manifestValidationSummary.textContent = `${sectionCount} section${sectionCount === 1 ? '' : 's'} ready.`;
  els.manifestValidationSummary.className = 'validation-summary success';
}

function validateSections(sections) {
  const byKey = new Map(sections.map((section) => [section.key, { errors: [], warnings: [] }]));
  const filenameGroups = new Map();

  sections.forEach((section) => {
    const issues = byKey.get(section.key);
    const rawFilename = String(section.filename || '').trim();
    const description = String(section.description || '').trim();
    const label = String(section.label || '').trim();

    if (!rawFilename) issues.errors.push('Output filename is blank.');
    if (!description) issues.warnings.push('Description is blank.');
    if (!label) issues.warnings.push('Label is blank.');
    if (section.pageCount > 150) issues.warnings.push(`Large section: ${section.pageCount} pages.`);

    if (rawFilename) {
      const sanitized = sanitizePdfFilename(rawFilename);
      if (sanitized !== rawFilename) issues.warnings.push(`Filename will be saved as “${sanitized}”.`);
      const normalized = sanitized.toLocaleLowerCase();
      if (!filenameGroups.has(normalized)) filenameGroups.set(normalized, []);
      filenameGroups.get(normalized).push(section.key);
    }
  });

  filenameGroups.forEach((keys) => {
    if (keys.length < 2) return;
    keys.forEach((key) => byKey.get(key).errors.push('Duplicate output filename.'));
  });

  if (els.preserveTabPdfs.checked) {
    const tabGroups = new Map();
    getTabAssets().forEach((tab) => {
      const normalized = tab.filename.toLocaleLowerCase();
      if (!tabGroups.has(normalized)) tabGroups.set(normalized, []);
      tabGroups.get(normalized).push(tab);
    });
    tabGroups.forEach((tabs) => {
      if (tabs.length < 2) return;
      tabs.forEach((tab) => {
        const section = sections.find((candidate) => candidate.tabPage === tab.tabPage);
        if (section) byKey.get(section.key).errors.push('Duplicate preserved tab filename; use distinct tab labels.');
      });
    });
  }

  let errorCount = 0;
  let warningCount = 0;
  byKey.forEach((issues) => {
    errorCount += issues.errors.length;
    warningCount += issues.warnings.length;
  });

  return { byKey, errorCount, warningCount };
}

function hideCurrentSectionEditor() {
  splitState.currentSectionKey = null;
  els.currentSectionEditor.classList.add('is-hidden');
}

function updateCurrentSectionEditor() {
  if (!splitState.pdfjsDoc) {
    hideCurrentSectionEditor();
    return;
  }

  const section = getSectionForPage(splitState.selectedPage);
  if (!section) {
    hideCurrentSectionEditor();
    return;
  }

  splitState.currentSectionKey = section.key;
  els.currentSectionEditor.classList.remove('is-hidden');
  els.currentSectionPages.textContent = `Pages ${section.startPage}–${section.endPage}`;
  els.currentSectionLabel.value = section.label;
  els.currentSectionDescription.value = section.description;
  els.currentSectionDate.value = section.date;
  els.currentSectionFilename.value = section.filename;
  els.currentSectionFilenameMode.textContent = section.filenameLocked ? 'manual filename' : 'auto filename';
  els.currentSectionFilenameMode.className = `filename-mode${section.filenameLocked ? ' manual' : ''}`;
}

function syncCurrentSectionEditorValues() {
  const section = getSectionByKey(splitState.currentSectionKey);
  if (!section) return;
  if (document.activeElement !== els.currentSectionLabel) els.currentSectionLabel.value = section.label;
  if (document.activeElement !== els.currentSectionDescription) els.currentSectionDescription.value = section.description;
  if (document.activeElement !== els.currentSectionDate) els.currentSectionDate.value = section.date;
  if (document.activeElement !== els.currentSectionFilename) els.currentSectionFilename.value = section.filename;
  els.currentSectionPages.textContent = `Pages ${section.startPage}–${section.endPage}`;
  els.currentSectionFilenameMode.textContent = section.filenameLocked ? 'manual filename' : 'auto filename';
  els.currentSectionFilenameMode.className = `filename-mode${section.filenameLocked ? ' manual' : ''}`;
}

function updateCurrentSectionField(field, value, lockFilename = false) {
  const key = splitState.currentSectionKey;
  if (!key) return;
  updateSectionMeta(key, field, value, lockFilename);
  renderManifestTable();

  if (field !== 'filename') {
    const refreshed = getSectionByKey(key);
    if (refreshed && !refreshed.filenameLocked) els.currentSectionFilename.value = refreshed.filename;
  }

  const refreshed = getSectionByKey(key);
  if (refreshed) {
    els.currentSectionFilenameMode.textContent = refreshed.filenameLocked ? 'manual filename' : 'auto filename';
    els.currentSectionFilenameMode.className = `filename-mode${refreshed.filenameLocked ? ' manual' : ''}`;
  }
}

function renderFilenameTemplate(template, context) {
  return sanitizePdfFilename(renderFilenameTemplateRaw(template, context));
}

function renderFilenameTemplateRaw(template, context) {
  const source = String(template || '').trim() || '{order:02} - {description}.pdf';
  let rendered = source.replace(/\{([^}]+)\}/g, (_, tokenText) => {
    const [token, format] = String(tokenText).split(':');
    const key = token.trim();
    let value = context[key] ?? '';

    if (key === 'order' && format && /^0?\d+$/.test(format)) {
      const width = Number(format.replace(/^0/, '')) || Number(format);
      value = String(value).padStart(width, '0');
    }

    return String(value);
  });

  rendered = rendered.replace(/\s+/g, ' ').trim();
  if (!rendered.toLowerCase().endsWith('.pdf')) rendered += '.pdf';
  return rendered;
}

async function splitAndDownload() {
  const sections = getSections();
  if (!splitState.bytes || !sections.length) return;

  const validation = validateSections(sections);
  if (validation.errorCount) {
    setSplitStatus('Fix the blocking manifest issues before splitting.', 'error');
    return;
  }

  els.splitDownload.disabled = true;
  try {
    setSplitStatus('Preparing source PDF…');
    const source = await PDFDocument.load(splitState.bytes.slice());
    const zip = new JSZip();
    const documentsFolder = zip.folder('documents');
    const tabsFolder = els.preserveTabPdfs.checked ? zip.folder('tabs') : null;

    for (let i = 0; i < sections.length; i += 1) {
      const section = sections[i];
      setSplitStatus(`Creating document ${i + 1} of ${sections.length}: ${section.filename}`);

      const output = await PDFDocument.create();
      const pages = await output.copyPages(source, section.pageIndexes);
      pages.forEach((page) => output.addPage(page));
      const bytes = await output.save();
      documentsFolder.file(sanitizePdfFilename(section.filename), bytes);
    }

    const tabAssets = tabsFolder ? getTabAssets() : [];
    if (tabsFolder) {
      for (let i = 0; i < tabAssets.length; i += 1) {
        const tab = tabAssets[i];
        setSplitStatus(`Preserving tab divider ${i + 1} of ${tabAssets.length}: ${tab.filename}`);
        const tabOutput = await PDFDocument.create();
        const [tabPage] = await tabOutput.copyPages(source, [tab.tabIndex]);
        tabOutput.addPage(tabPage);
        const tabBytes = await tabOutput.save();
        tabsFolder.file(tab.filename, tabBytes);
      }
    }

    const manifest = makeManifestObject();
    zip.file('binder-manifest.json', JSON.stringify(manifest, null, 2));
    zip.file('binder-manifest.csv', manifestToCsv(manifest.sections));
    zip.file('binder-manifest.xlsx', makeManifestWorkbookBytes(manifest));

    setSplitStatus('Packaging ZIP…');
    const blob = await zip.generateAsync({ type: 'blob' });
    const base = stripExtension(splitState.file.name);
    downloadBlob(blob, `${base} - Split PDF.zip`);
    const tabNote = tabAssets.length ? ` and ${tabAssets.length} separate tab PDF${tabAssets.length === 1 ? '' : 's'}` : '';
    setSplitStatus(`Created ${sections.length} content PDF${sections.length === 1 ? '' : 's'}${tabNote}, plus JSON, CSV, and Excel manifests.`, 'success');
  } catch (error) {
    console.error(error);
    setSplitStatus(`Could not split the PDF: ${friendlyError(error)}`, 'error');
  } finally {
    renderManifestTable();
  }
}

function makeManifestObject() {
  const tabAssets = els.preserveTabPdfs.checked ? getTabAssets() : [];
  const tabFilenameByPage = new Map(tabAssets.map((tab) => [tab.tabPage, tab.filename]));

  return {
    schema: 'pdf-binder-tool/6',
    sourceFile: splitState.file?.name ?? null,
    generatedAt: new Date().toISOString(),
    tabSemantics: 'position',
    filenameTemplate: els.filenameTemplate.value,
    splitSemantics: {
      normalCut: 'split-after-page; source page retained',
      tabDivider: 'divider page excluded from content output',
      leadingPages: 'omitted before first tab divider unless a normal split cut is placed before that divider',
    },
    tabs: tabAssets.map((tab) => ({
      order: tab.order,
      tabPage: tab.tabPage,
      label: tab.label,
      filename: tab.filename,
    })),
    sections: getSections().map((section) => ({
      order: section.order,
      boundaryType: section.boundary.type,
      boundaryPage: section.boundary.pageIndex == null ? null : section.boundary.pageIndex + 1,
      tabPage: section.tabPage,
      splitAfterPage: section.splitAfterPage,
      startPage: section.startPage,
      endPage: section.endPage,
      pageCount: section.pageCount,
      label: section.label,
      description: section.description,
      date: section.date,
      filename: sanitizePdfFilename(section.filename),
      filenameMode: section.filenameLocked ? 'manual' : 'template',
      tabFilename: section.tabPage == null ? null : tabFilenameByPage.get(section.tabPage) ?? null,
    })),
  };
}

function exportManifestJson() {
  const manifest = makeManifestObject();
  downloadBlob(new Blob([JSON.stringify(manifest, null, 2)], { type: 'application/json' }), 'binder-manifest.json');
}

function exportManifestCsv() {
  const manifest = makeManifestObject();
  downloadBlob(new Blob([manifestToCsv(manifest.sections)], { type: 'text/csv;charset=utf-8' }), 'binder-manifest.csv');
}

function exportManifestXlsx() {
  const manifest = makeManifestObject();
  downloadBlob(
    new Blob([makeManifestWorkbookBytes(manifest)], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }),
    'binder-manifest.xlsx',
  );
}

function manifestToCsv(sections) {
  const headers = ['Order', 'Boundary Type', 'Boundary Page', 'Tab Page (Excluded)', 'Split After Page', 'Start Page', 'End Page', 'Page Count', 'Label', 'Description', 'Date', 'Filename', 'Filename Mode', 'Tab Filename'];
  const rows = sections.map((s) => [s.order, s.boundaryType, s.boundaryPage, s.tabPage, s.splitAfterPage, s.startPage, s.endPage, s.pageCount, s.label, s.description, s.date, s.filename, s.filenameMode, s.tabFilename]);
  return [headers, ...rows].map((row) => row.map(csvEscape).join(',')).join('\r\n');
}

function makeManifestWorkbookBytes(manifest) {
  const workbook = XLSX.utils.book_new();
  const rows = manifest.sections.map((section) => ({
    Order: section.order,
    Label: section.label,
    Description: section.description,
    Date: section.date,
    Filename: section.filename,
    'Filename Mode': section.filenameMode,
    'Boundary Type': section.boundaryType,
    'Boundary Page': section.boundaryPage,
    'Tab Page (Excluded)': section.tabPage,
    'Split After Page': section.splitAfterPage,
    'Start Page': section.startPage,
    'End Page': section.endPage,
    'Page Count': section.pageCount,
    'Tab Filename': section.tabFilename,
  }));
  const manifestSheet = XLSX.utils.json_to_sheet(rows);
  manifestSheet['!cols'] = [
    { wch: 8 }, { wch: 18 }, { wch: 38 }, { wch: 18 }, { wch: 44 }, { wch: 14 },
    { wch: 14 }, { wch: 14 }, { wch: 18 }, { wch: 18 }, { wch: 12 },
    { wch: 12 }, { wch: 12 }, { wch: 24 },
  ];
  XLSX.utils.book_append_sheet(workbook, manifestSheet, 'Manifest');

  const instructions = [
    ['PDF Binder Tool Manifest'],
    ['Source PDF', manifest.sourceFile || ''],
    ['Schema', manifest.schema],
    ['Filename template', manifest.filenameTemplate || ''],
    [],
    ['Editing guidance'],
    ['You may edit Order, Label, Description, Date, Filename, and Tab Filename for binder planning.'],
    ['When importing this workbook back into an already-open source PDF, the app uses Label, Description, Date, and Filename only.'],
    ['Page/boundary columns are reference data; spreadsheet edits do not change source-PDF split points.'],
  ];
  const instructionsSheet = XLSX.utils.aoa_to_sheet(instructions);
  instructionsSheet['!cols'] = [{ wch: 24 }, { wch: 100 }];
  XLSX.utils.book_append_sheet(workbook, instructionsSheet, 'Instructions');

  return XLSX.write(workbook, { type: 'array', bookType: 'xlsx' });
}

async function importSplitManifest(event) {
  const [file] = event.target.files;
  if (!file) return;

  try {
    if (!splitState.pdfjsDoc) throw new Error('Open the source PDF before importing manifest edits.');
    const parsed = await readManifestFile(file);
    if (!parsed?.sections?.length) throw new Error('Manifest does not contain any sections.');

    const currentSections = getSections();
    let applied = 0;
    parsed.sections.forEach((imported, index) => {
      const order = Number(imported.order) || index + 1;
      const current = currentSections[order - 1];
      if (!current) return;
      const meta = getOrCreateSectionMeta(current.key);
      if (imported.label != null && imported.label !== '') meta.label = String(imported.label);
      if (imported.description != null && imported.description !== '') meta.description = String(imported.description);
      if (imported.date != null && imported.date !== '') meta.date = String(imported.date);
      if (imported.filename != null && imported.filename !== '') {
        meta.filename = String(imported.filename);
        meta.filenameLocked = true;
      }
      applied += 1;
    });

    if (parsed.filenameTemplate) els.filenameTemplate.value = String(parsed.filenameTemplate);
    rebuildManifest();
    setSplitStatus(`Imported manifest edits for ${applied} of ${currentSections.length} current section${currentSections.length === 1 ? '' : 's'}. Page divisions were not changed.`, 'success');
  } catch (error) {
    console.error(error);
    setSplitStatus(`Could not import manifest: ${friendlyError(error)}`, 'error');
  } finally {
    event.target.value = '';
  }
}

function addBuildFiles(files) {
  const pdfs = files.filter(isPdfFile);
  buildState.entries.push(...pdfs.map(makeBuildEntry));
  applyBuildManifestOrder();
  applyBuildManifestMetadata();
  renderBuildTable();
  updateBuildStatusSummary();
}

function makeBuildEntry(file) {
  return {
    id: makeId(),
    file,
    description: humanizeFilename(file.name),
    date: '',
  };
}

async function addBuildTabFiles(files) {
  const pdfs = files.filter(isPdfFile);
  const tabEntries = pdfs.map((file) => ({
    id: makeId(),
    file,
    bookmarkTitle: fallbackBookmarkTitle(file.name),
    bookmarkSource: 'filename',
    bookmarkStatus: 'reading',
  }));

  buildState.tabFiles.push(...tabEntries);
  applyBuildTabOrder();
  renderBuildTable();
  updateBuildStatusSummary();

  await Promise.all(tabEntries.map(readTabBookmarkTitle));
  applyBuildTabOrder();
  renderBuildTable();
  updateBuildStatusSummary();
}

async function readTabBookmarkTitle(tabEntry) {
  let loadingTask = null;
  try {
    const bytes = new Uint8Array(await tabEntry.file.arrayBuffer());
    loadingTask = pdfjsLib.getDocument({ data: bytes });
    const pdf = await loadingTask.promise;
    const outline = await pdf.getOutline();
    const title = firstOutlineTitle(outline);

    if (title) {
      tabEntry.bookmarkTitle = title;
      tabEntry.bookmarkSource = 'existing';
    }
    tabEntry.bookmarkStatus = 'ready';
  } catch (error) {
    console.warn(`Could not read bookmark from ${tabEntry.file.name}; using filename fallback.`, error);
    tabEntry.bookmarkStatus = 'fallback';
  } finally {
    if (loadingTask) {
      try {
        await loadingTask.destroy();
      } catch (cleanupError) {
        console.warn(`Could not fully release PDF.js resources for ${tabEntry.file.name}.`, cleanupError);
      }
    }
  }
}

function firstOutlineTitle(items) {
  if (!Array.isArray(items)) return '';
  for (const item of items) {
    const title = String(item?.title || '').trim();
    if (title) return title;
    const nested = firstOutlineTitle(item?.items);
    if (nested) return nested;
  }
  return '';
}

function fallbackBookmarkTitle(filename) {
  const title = stripExtension(filename)
    .replace(/[_]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  return title || 'Tab';
}

async function loadBuildIndex(file) {
  if (!isPdfFile(file)) {
    setBuildStatus('The index must be a PDF file.', 'error');
    return;
  }

  buildState.indexFile = file;
  buildState.indexInfo = null;
  buildState.indexLoading = true;
  renderIndexPanel();
  renderBuildTable();
  updateBuildStatusSummary();

  let loadingTask = null;
  try {
    const bytes = new Uint8Array(await file.arrayBuffer());
    loadingTask = pdfjsLib.getDocument({ data: bytes });
    const pdf = await loadingTask.promise;
    const rows = [];
    let existingLinkRectsUsed = 0;

    for (let pageIndex = 0; pageIndex < pdf.numPages; pageIndex += 1) {
      const page = await pdf.getPage(pageIndex + 1);
      const pageRows = await detectIndexRowsOnPage(page, pageIndex);
      existingLinkRectsUsed += pageRows.filter((row) => row.rectSource === 'existing-link').length;
      rows.push(...pageRows);
    }

    rows.sort((a, b) => a.number - b.number || a.pageIndex - b.pageIndex || b.rect[3] - a.rect[3]);
    buildState.indexInfo = {
      pageCount: pdf.numPages,
      rows,
      existingLinkRectsUsed,
    };
  } catch (error) {
    console.error(`Could not analyze index ${file.name}.`, error);
    buildState.indexInfo = {
      pageCount: 0,
      rows: [],
      existingLinkRectsUsed: 0,
      error: friendlyError(error),
    };
  } finally {
    if (loadingTask) {
      try {
        await loadingTask.destroy();
      } catch (cleanupError) {
        console.warn(`Could not fully release PDF.js resources for ${file.name}.`, cleanupError);
      }
    }
    buildState.indexLoading = false;
    renderIndexPanel();
    renderBuildTable();
    updateBuildStatusSummary();
  }
}

async function detectIndexRowsOnPage(page, pageIndex) {
  const textContent = await page.getTextContent();
  const items = textContent.items
    .map(normalizeIndexTextItem)
    .filter(Boolean);

  const pageView = page.view || [0, 0, 612, 792];
  const pageX0 = Number(pageView[0]) || 0;
  const pageY0 = Number(pageView[1]) || 0;
  const pageX1 = Number(pageView[2]) || 612;
  const pageY1 = Number(pageView[3]) || 792;
  const pageWidth = Math.max(1, pageX1 - pageX0);

  const descriptionHeader = items.find((item) => normalizeIndexText(item.str) === 'description');
  const dateHeader = items.find((item) => normalizeIndexText(item.str) === 'date');
  const noHeader = items.find((item) => ['no', 'no.'].includes(normalizeIndexText(item.str)));
  const headerY = descriptionHeader?.y ?? noHeader?.y ?? dateHeader?.y ?? null;

  let numberItems = items
    .map((item) => ({ item, number: parseIndexRowNumber(item.str) }))
    .filter(({ item, number }) => Number.isInteger(number)
      && number > 0
      && number < 1000
      && item.x < pageX0 + pageWidth * 0.24
      && (headerY == null || item.y < headerY - 2))
    .sort((a, b) => b.item.y - a.item.y || a.item.x - b.item.x);

  // Keep only one visible number marker per row number/page location.
  numberItems = numberItems.filter((candidate, index, list) => !list.slice(0, index).some((prior) => (
    prior.number === candidate.number && Math.abs(prior.item.y - candidate.item.y) < 4
  )));

  if (!numberItems.length) return [];

  let existingLinks = [];
  try {
    const annotations = await page.getAnnotations({ intent: 'any' });
    existingLinks = annotations
      .filter((annotation) => annotation?.subtype === 'Link' && Array.isArray(annotation.rect) && annotation.rect.length === 4)
      .map((annotation) => ({ rect: normalizePdfRect(annotation.rect) }))
      .filter((annotation) => annotation.rect);
  } catch (error) {
    console.warn(`Could not inspect existing index links on page ${pageIndex + 1}.`, error);
  }

  const dateLikeItems = items.filter((item) => looksLikeIndexDate(item.str) && item.x > pageX0 + pageWidth * 0.62);
  const dateLeft = dateLikeItems.length
    ? Math.min(...dateLikeItems.map((item) => item.x))
    : dateHeader
      ? dateHeader.x - pageWidth * 0.025
      : pageX0 + pageWidth * 0.79;

  const rows = [];
  numberItems.forEach((candidate, index) => {
    const currentY = candidate.item.y;
    const previousY = index > 0 ? numberItems[index - 1].item.y : null;
    const nextY = index < numberItems.length - 1 ? numberItems[index + 1].item.y : null;

    let upper = previousY != null
      ? (previousY + currentY) / 2
      : headerY != null
        ? (headerY + currentY) / 2
        : currentY + Math.max(candidate.item.height * 1.4, 12);
    let lower = nextY != null
      ? (currentY + nextY) / 2
      : currentY - Math.max(previousY != null ? (previousY - currentY) / 2 : 12, candidate.item.height * 1.4, 12);

    upper = Math.min(pageY1, upper);
    lower = Math.max(pageY0, lower);

    const rowItems = items.filter((item) => item.y <= upper + 2 && item.y >= lower - 2);
    const descriptionItems = rowItems.filter((item) => (
      item.x > candidate.item.x + Math.max(candidate.item.width, 8) + 8
      && item.x < dateLeft - 4
      && parseIndexRowNumber(item.str) == null
    ));
    const dateItems = rowItems.filter((item) => item.x >= dateLeft - 4 && item.x <= pageX1);

    const description = joinIndexTextItems(descriptionItems);
    const date = joinIndexTextItems(dateItems);
    if (!description) return;

    const descriptionLeft = Math.max(
      pageX0,
      descriptionItems.length ? Math.min(...descriptionItems.map((item) => item.x)) - 6 : candidate.item.x + 24,
    );
    const descriptionRight = Math.min(pageX1, Math.max(descriptionLeft + 20, dateLeft - 10));
    let rect = normalizePdfRect([descriptionLeft, lower + 1, descriptionRight, upper - 1]);
    let rectSource = 'text-detection';

    const linkMatch = existingLinks
      .map((link) => ({
        link,
        distance: Math.abs(((link.rect[1] + link.rect[3]) / 2) - currentY),
        overlaps: link.rect[3] >= lower && link.rect[1] <= upper,
      }))
      .filter((match) => match.overlaps)
      .sort((a, b) => a.distance - b.distance)[0];

    if (linkMatch?.link?.rect) {
      rect = linkMatch.link.rect;
      rectSource = 'existing-link';
    }

    if (!rect) return;
    rows.push({
      number: candidate.number,
      pageIndex,
      description,
      date,
      rect,
      rectSource,
    });
  });

  return rows;
}

function normalizeIndexTextItem(item) {
  const str = String(item?.str || '').trim();
  const transform = item?.transform;
  if (!str || !Array.isArray(transform) || transform.length < 6) return null;
  const x = Number(transform[4]);
  const y = Number(transform[5]);
  if (!Number.isFinite(x) || !Number.isFinite(y)) return null;
  const width = Number.isFinite(Number(item.width)) ? Math.abs(Number(item.width)) : 0;
  const heightCandidates = [
    Number(item.height),
    Math.hypot(Number(transform[0]) || 0, Number(transform[1]) || 0),
    Math.hypot(Number(transform[2]) || 0, Number(transform[3]) || 0),
  ].filter((value) => Number.isFinite(value) && value > 0);
  const height = heightCandidates.length ? Math.max(...heightCandidates) : 10;
  return { str, x, y, width, height };
}

function parseIndexRowNumber(value) {
  const match = String(value || '').trim().match(/^(\d{1,3})\s*[.)]?$/);
  if (!match) return null;
  return Number(match[1]);
}

function looksLikeIndexDate(value) {
  const text = String(value || '').trim();
  return /^\d{1,2}\/\d{1,2}\/\d{2,4}$/.test(text)
    || /^\d{4}\s*[-–]\s*\d{4}$/.test(text)
    || /^\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}$/.test(text);
}

function normalizeIndexText(value) {
  return String(value || '').trim().toLowerCase().replace(/\s+/g, ' ');
}

function joinIndexTextItems(items) {
  if (!items.length) return '';
  const ordered = [...items].sort((a, b) => {
    if (Math.abs(a.y - b.y) > 3) return b.y - a.y;
    return a.x - b.x;
  });
  return ordered.map((item) => item.str).join(' ').replace(/\s+/g, ' ').trim();
}

function normalizePdfRect(rect) {
  if (!Array.isArray(rect) || rect.length !== 4) return null;
  const values = rect.map(Number);
  if (values.some((value) => !Number.isFinite(value))) return null;
  return [
    Math.min(values[0], values[2]),
    Math.min(values[1], values[3]),
    Math.max(values[0], values[2]),
    Math.max(values[1], values[3]),
  ];
}

function getIndexLinkValidation() {
  if (!buildState.indexFile || !els.createIndexLinks.checked) return { ok: true, message: '' };
  if (buildState.indexLoading) return { ok: false, message: 'Index analysis is still running.' };
  if (buildState.indexInfo?.error) return { ok: false, message: `Could not analyze the index: ${buildState.indexInfo.error}` };

  const rows = buildState.indexInfo?.rows || [];
  const contentCount = buildState.entries.length;
  if (!rows.length) return { ok: false, message: 'No numbered index rows were detected. Turn off index linking to build without links.' };
  if (!contentCount) return { ok: false, message: 'Add content PDFs so the detected index rows have destinations.' };

  const counts = new Map();
  rows.forEach((row) => counts.set(row.number, (counts.get(row.number) || 0) + 1));
  const duplicates = [...counts.entries()].filter(([, count]) => count > 1).map(([number]) => number);
  if (duplicates.length) {
    return { ok: false, message: `Duplicate index row number${duplicates.length === 1 ? '' : 's'} detected: ${duplicates.join(', ')}.` };
  }

  const missing = [];
  for (let number = 1; number <= contentCount; number += 1) {
    if (!counts.has(number)) missing.push(number);
  }
  const extra = rows.map((row) => row.number).filter((number) => number > contentCount);
  if (missing.length || extra.length || rows.length !== contentCount) {
    const details = [];
    if (missing.length) details.push(`missing row${missing.length === 1 ? '' : 's'} ${missing.join(', ')}`);
    if (extra.length) details.push(`row${extra.length === 1 ? '' : 's'} beyond the loaded content count: ${extra.join(', ')}`);
    if (!details.length) details.push(`${rows.length} rows detected for ${contentCount} content PDFs`);
    return {
      ok: false,
      message: `Index/content mismatch (${details.join('; ')}). For this first version, linked index rows must be numbered 1 through ${contentCount}.`,
    };
  }

  return { ok: true, message: '' };
}

function renderIndexPanel() {
  els.clearBuildIndex.disabled = !buildState.indexFile;
  els.createIndexLinks.disabled = !buildState.indexFile || buildState.indexLoading || !buildState.indexInfo?.rows?.length;

  if (!buildState.indexFile) {
    els.indexDetectionSummary.textContent = 'No index PDF loaded.';
    els.indexDetectionSummary.className = 'index-detection-summary';
    els.indexLinkBody.innerHTML = '<tr><td colspan="5" class="empty-cell">Add an index PDF to detect its numbered rows.</td></tr>';
    return;
  }

  if (buildState.indexLoading) {
    els.indexDetectionSummary.textContent = `Analyzing ${buildState.indexFile.name}…`;
    els.indexDetectionSummary.className = 'index-detection-summary';
    els.indexLinkBody.innerHTML = '<tr><td colspan="5" class="empty-cell">Detecting numbered index rows and Description cells…</td></tr>';
    return;
  }

  if (buildState.indexInfo?.error) {
    els.indexDetectionSummary.textContent = `Could not analyze ${buildState.indexFile.name}.`;
    els.indexDetectionSummary.className = 'index-detection-summary error';
    els.indexLinkBody.innerHTML = `<tr><td colspan="5" class="empty-cell">${escapeHtml(buildState.indexInfo.error)}</td></tr>`;
    return;
  }

  const info = buildState.indexInfo || { rows: [], pageCount: 0, existingLinkRectsUsed: 0 };
  const validation = getIndexLinkValidation();
  const reusedNote = info.existingLinkRectsUsed
    ? ` ${info.existingLinkRectsUsed} existing link rectangle${info.existingLinkRectsUsed === 1 ? '' : 's'} reused for placement.`
    : '';
  els.indexDetectionSummary.textContent = `${buildState.indexFile.name} — ${info.pageCount} page${info.pageCount === 1 ? '' : 's'}, ${info.rows.length} numbered row${info.rows.length === 1 ? '' : 's'} detected.${reusedNote}`;
  els.indexDetectionSummary.className = `index-detection-summary${validation.ok || !els.createIndexLinks.checked ? ' success' : ' error'}`;

  if (!info.rows.length) {
    els.indexLinkBody.innerHTML = '<tr><td colspan="5" class="empty-cell">No numbered Description rows detected.</td></tr>';
    return;
  }

  els.indexLinkBody.innerHTML = '';
  info.rows.forEach((row) => {
    const entry = buildState.entries[row.number - 1];
    const destination = entry
      ? `Position ${row.number} — ${entry.description || entry.file.name}`
      : `No content PDF at position ${row.number}`;
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${row.number}</td>
      <td>${row.pageIndex + 1}</td>
      <td>${escapeHtml(row.description)}</td>
      <td>${escapeHtml(row.date || '')}</td>
      <td class="${entry ? 'index-destination' : 'index-destination missing'}">${escapeHtml(destination)}</td>
    `;
    els.indexLinkBody.append(tr);
  });
}

async function importBuildManifest(event) {
  const [file] = event.target.files;
  if (!file) return;

  try {
    const parsed = await readManifestFile(file);
    if (!parsed || !Array.isArray(parsed.sections)) throw new Error('Manifest does not contain a sections array.');
    buildState.manifest = parsed;
    applyBuildManifestOrder();
    applyBuildManifestMetadata();
    applyBuildTabOrder();
    renderBuildTable();

    const matched = countManifestMatches();
    const tabMatched = countManifestTabMatches();
    const tabPhrase = parsed.sections.some((section) => section.tabFilename)
      ? ` ${tabMatched} tab filename${tabMatched === 1 ? '' : 's'} also match loaded tab PDFs.`
      : '';
    setBuildStatus(`Manifest loaded. ${matched} of ${parsed.sections.length} content filenames currently match loaded PDFs.${tabPhrase}`, matched === parsed.sections.length && matched > 0 ? 'success' : '');
  } catch (error) {
    console.error(error);
    setBuildStatus(`Could not read manifest: ${friendlyError(error)}`, 'error');
  } finally {
    event.target.value = '';
  }
}

async function readManifestFile(file) {
  const lower = file.name.toLowerCase();
  if (lower.endsWith('.json')) {
    const parsed = JSON.parse(await file.text());
    if (!parsed?.sections) throw new Error('JSON manifest does not contain a sections array.');
    return parsed;
  }

  if (lower.endsWith('.xlsx') || lower.endsWith('.xls')) {
    const workbook = XLSX.read(await file.arrayBuffer(), { type: 'array' });
    const sheet = workbook.Sheets.Manifest || workbook.Sheets[workbook.SheetNames[0]];
    if (!sheet) throw new Error('Workbook does not contain a readable worksheet.');
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
    if (!rows.length) throw new Error('Manifest worksheet is empty.');
    return manifestFromSpreadsheetRows(rows);
  }

  throw new Error('Use a JSON or Excel (.xlsx/.xls) manifest.');
}

function manifestFromSpreadsheetRows(rows) {
  const sections = rows.map((row, index) => ({
    order: numberOrNull(getSpreadsheetValue(row, ['Order'])) ?? index + 1,
    label: getSpreadsheetValue(row, ['Label']),
    description: getSpreadsheetValue(row, ['Description']),
    date: getSpreadsheetValue(row, ['Date', 'Document Date']),
    filename: getSpreadsheetValue(row, ['Filename', 'Output Filename', 'Content PDF']),
    filenameMode: getSpreadsheetValue(row, ['Filename Mode']),
    boundaryType: getSpreadsheetValue(row, ['Boundary Type']),
    boundaryPage: numberOrNull(getSpreadsheetValue(row, ['Boundary Page'])),
    tabPage: numberOrNull(getSpreadsheetValue(row, ['Tab Page (Excluded)', 'Tab Page'])),
    splitAfterPage: numberOrNull(getSpreadsheetValue(row, ['Split After Page'])),
    startPage: numberOrNull(getSpreadsheetValue(row, ['Start Page'])),
    endPage: numberOrNull(getSpreadsheetValue(row, ['End Page'])),
    pageCount: numberOrNull(getSpreadsheetValue(row, ['Page Count'])),
    tabFilename: getSpreadsheetValue(row, ['Tab Filename']),
  }));

  return {
    schema: 'pdf-binder-tool/spreadsheet-import',
    sections,
  };
}

function getSpreadsheetValue(row, candidates) {
  const normalized = new Map(
    Object.entries(row).map(([key, value]) => [normalizeHeader(key), value]),
  );
  for (const candidate of candidates) {
    const value = normalized.get(normalizeHeader(candidate));
    if (value != null && value !== '') return value;
  }
  return '';
}

function normalizeHeader(value) {
  return String(value).toLowerCase().replace(/[^a-z0-9]+/g, '');
}

function numberOrNull(value) {
  if (value == null || value === '') return null;
  const number = Number(value);
  return Number.isFinite(number) ? number : null;
}

function applyBuildManifestOrder() {
  if (!buildState.manifest?.sections?.length || !buildState.entries.length) return;
  const orderMap = new Map(buildState.manifest.sections.map((section, index) => [section.filename, section.order ?? index + 1]));
  buildState.entries.sort((a, b) => {
    const aOrder = orderMap.get(a.file.name);
    const bOrder = orderMap.get(b.file.name);
    if (aOrder == null && bOrder == null) return naturalCompare(a.file.name, b.file.name);
    if (aOrder == null) return 1;
    if (bOrder == null) return -1;
    return aOrder - bOrder;
  });
}

function applyBuildManifestMetadata() {
  if (!buildState.manifest?.sections?.length) return;
  const byFilename = new Map(buildState.manifest.sections.map((section) => [section.filename, section]));
  buildState.entries.forEach((entry) => {
    const match = byFilename.get(entry.file.name);
    if (match?.description != null && match.description !== '') entry.description = String(match.description);
    if (match?.date != null && match.date !== '') entry.date = String(match.date);
  });
}

function applyBuildTabOrder() {
  if (!buildState.tabFiles.length) return;

  const manifestTabOrder = new Map(
    (buildState.manifest?.sections || [])
      .filter((section) => section.tabFilename)
      .map((section, index) => [section.tabFilename, section.order ?? index + 1]),
  );

  buildState.tabFiles.sort((a, b) => {
    const aOrder = manifestTabOrder.get(a.file.name);
    const bOrder = manifestTabOrder.get(b.file.name);
    if (aOrder == null && bOrder == null) return naturalCompare(a.file.name, b.file.name);
    if (aOrder == null) return 1;
    if (bOrder == null) return -1;
    return aOrder - bOrder;
  });
}

function countManifestMatches() {
  if (!buildState.manifest?.sections) return 0;
  const loaded = new Set(buildState.entries.map((entry) => entry.file.name));
  return buildState.manifest.sections.filter((section) => loaded.has(section.filename)).length;
}

function countManifestTabMatches() {
  if (!buildState.manifest?.sections) return 0;
  const loaded = new Set(buildState.tabFiles.map((tab) => tab.file.name));
  return buildState.manifest.sections.filter((section) => section.tabFilename && loaded.has(section.tabFilename)).length;
}

function renderBuildTable() {
  renderIndexPanel();
  els.clearBuildTabs.disabled = buildState.tabFiles.length === 0;
  els.exportBuildXlsx.disabled = buildState.entries.length === 0;
  const tabsEnabled = els.insertBuildTabs.checked;
  const hasEnoughTabs = !tabsEnabled || buildState.tabFiles.length >= buildState.entries.length;
  const indexValidation = getIndexLinkValidation();
  els.buildDownload.disabled = buildState.entries.length === 0 || !hasEnoughTabs || !indexValidation.ok;

  if (!buildState.entries.length) {
    els.buildBody.innerHTML = '<tr><td colspan="8" class="empty-cell">No content PDFs loaded.</td></tr>';
    return;
  }

  els.buildBody.innerHTML = '';
  buildState.entries.forEach((entry, index) => {
    const tabFile = tabsEnabled ? buildState.tabFiles[index] : null;
    const tabCell = !tabsEnabled
      ? '<span class="tab-missing">Not inserted</span>'
      : tabFile
        ? `<span class="tab-file-name">${escapeHtml(tabFile.file.name)}</span>
           <span class="tab-bookmark">${tabFile.bookmarkStatus === 'reading'
             ? 'Bookmark: checking PDF…'
             : `Bookmark: ${escapeHtml(tabFile.bookmarkTitle)}${tabFile.bookmarkSource === 'existing' ? ' (from PDF)' : ' (filename fallback)'}`}</span>`
        : '<span class="tab-missing">Missing tab PDF</span>';

    const row = document.createElement('tr');
    row.dataset.entryId = entry.id;
    row.innerHTML = `
      <td><span class="drag-handle" draggable="true" title="Drag to reorder" aria-label="Drag to reorder">☰</span></td>
      <td>${index + 1}</td>
      <td>${tabCell}</td>
      <td><input class="build-description-input" data-action="description" value="${escapeAttr(entry.description)}" /></td>
      <td>${escapeHtml(entry.file.name)}</td>
      <td>${formatBytes(entry.file.size)}</td>
      <td>
        <div class="order-buttons">
          <button type="button" data-action="up" ${index === 0 ? 'disabled' : ''} aria-label="Move content up">↑</button>
          <button type="button" data-action="down" ${index === buildState.entries.length - 1 ? 'disabled' : ''} aria-label="Move content down">↓</button>
        </div>
      </td>
      <td><button type="button" class="remove-file" data-action="remove" aria-label="Remove content PDF">×</button></td>
    `;

    row.querySelector('[data-action="description"]').addEventListener('input', (event) => {
      entry.description = event.target.value;
    });
    row.querySelector('[data-action="up"]').addEventListener('click', () => moveBuildEntry(index, -1));
    row.querySelector('[data-action="down"]').addEventListener('click', () => moveBuildEntry(index, 1));
    row.querySelector('[data-action="remove"]').addEventListener('click', () => {
      buildState.entries.splice(index, 1);
      renderBuildTable();
      updateBuildStatusSummary();
    });

    attachBuildDragEvents(row, entry.id);
    els.buildBody.append(row);
  });
}

function attachBuildDragEvents(row, entryId) {
  const handle = row.querySelector('.drag-handle');
  handle.addEventListener('dragstart', (event) => {
    buildState.draggedEntryId = entryId;
    row.classList.add('dragging');
    event.dataTransfer.effectAllowed = 'move';
    event.dataTransfer.setData('text/plain', entryId);
  });
  handle.addEventListener('dragend', () => {
    buildState.draggedEntryId = null;
    [...els.buildBody.querySelectorAll('tr')].forEach((item) => item.classList.remove('dragging', 'drag-over'));
  });

  row.addEventListener('dragover', (event) => {
    if (!buildState.draggedEntryId || buildState.draggedEntryId === entryId) return;
    event.preventDefault();
    event.dataTransfer.dropEffect = 'move';
    row.classList.add('drag-over');
  });
  row.addEventListener('dragleave', () => row.classList.remove('drag-over'));
  row.addEventListener('drop', (event) => {
    event.preventDefault();
    row.classList.remove('drag-over');
    const draggedId = buildState.draggedEntryId || event.dataTransfer.getData('text/plain');
    reorderBuildEntry(draggedId, entryId);
  });
}

function reorderBuildEntry(draggedId, targetId) {
  const from = buildState.entries.findIndex((entry) => entry.id === draggedId);
  const to = buildState.entries.findIndex((entry) => entry.id === targetId);
  if (from < 0 || to < 0 || from === to) return;
  const [entry] = buildState.entries.splice(from, 1);
  buildState.entries.splice(to, 0, entry);
  renderBuildTable();
  updateBuildStatusSummary();
}

function moveBuildEntry(index, delta) {
  const target = index + delta;
  if (target < 0 || target >= buildState.entries.length) return;
  [buildState.entries[index], buildState.entries[target]] = [buildState.entries[target], buildState.entries[index]];
  renderBuildTable();
  updateBuildStatusSummary();
}

function updateBuildStatusSummary() {
  const contentCount = buildState.entries.length;
  const tabCount = buildState.tabFiles.length;
  els.createBuildBookmarks.disabled = !els.insertBuildTabs.checked || tabCount === 0;
  if (!contentCount) {
    setBuildStatus(tabCount
      ? `${tabCount} positional tab PDF${tabCount === 1 ? '' : 's'} loaded. Add content PDFs to build the binder.`
      : 'Add content PDFs in the order you want, or import a manifest.');
    return;
  }

  const duplicateNames = findDuplicateNames(buildState.entries.map((entry) => entry.file.name));
  const duplicateNote = duplicateNames.length ? ` Warning: duplicate content filename${duplicateNames.length === 1 ? '' : 's'} (${duplicateNames.join(', ')}) can make manifest matching ambiguous.` : '';
  const indexValidation = getIndexLinkValidation();
  if (!indexValidation.ok) {
    setBuildStatus(indexValidation.message, 'error');
    return;
  }
  const indexNote = buildState.indexFile
    ? els.createIndexLinks.checked
      ? ` The ${buildState.indexInfo?.pageCount || 0}-page index will be inserted first with ${buildState.indexInfo?.rows?.length || 0} Description link${buildState.indexInfo?.rows?.length === 1 ? '' : 's'}.`
      : ` The ${buildState.indexInfo?.pageCount || 0}-page index will be inserted first without creating new links.`
    : '';

  if (!els.insertBuildTabs.checked) {
    setBuildStatus(`${contentCount} content PDF${contentCount === 1 ? '' : 's'} loaded. Positional tab insertion is off.${indexNote}${duplicateNote}`, duplicateNames.length ? 'error' : '');
    return;
  }

  if (tabCount < contentCount) {
    const missing = contentCount - tabCount;
    setBuildStatus(`${contentCount} content PDF${contentCount === 1 ? '' : 's'} and ${tabCount} tab PDF${tabCount === 1 ? '' : 's'} loaded. Add ${missing} more tab PDF${missing === 1 ? '' : 's'} or turn off tab insertion.${indexNote}${duplicateNote}`, 'error');
    return;
  }

  const extra = tabCount - contentCount;
  const extraNote = extra ? ` ${extra} extra tab PDF${extra === 1 ? '' : 's'} will not be used.` : '';
  const bookmarkNote = els.createBuildBookmarks.checked ? ` ${contentCount} top-level tab bookmark${contentCount === 1 ? '' : 's'} will be created.` : '';
  setBuildStatus(`${contentCount} content PDF${contentCount === 1 ? '' : 's'} will be paired with the first ${contentCount} positional tab PDF${contentCount === 1 ? '' : 's'}.${bookmarkNote}${indexNote}${extraNote}${duplicateNote}`, duplicateNames.length ? 'error' : 'success');
}

function findDuplicateNames(names) {
  const counts = new Map();
  names.forEach((name) => {
    const key = name.toLocaleLowerCase();
    counts.set(key, (counts.get(key) || 0) + 1);
  });
  return [...counts.entries()].filter(([, count]) => count > 1).map(([name]) => name);
}

async function appendPdfToOutput(output, file) {
  const source = await PDFDocument.load(await file.arrayBuffer());
  const pages = await output.copyPages(source, source.getPageIndices());
  pages.forEach((page) => output.addPage(page));
  return pages.length;
}

async function appendIndexPdfToOutput(output, file) {
  const source = await PDFDocument.load(await file.arrayBuffer());
  const pages = await output.copyPages(source, source.getPageIndices());
  let removedLinks = 0;
  pages.forEach((page) => {
    output.addPage(page);
    removedLinks += removeLinkAnnotations(output, page);
  });
  return { pageCount: pages.length, removedLinks };
}

function removeLinkAnnotations(pdfDoc, page) {
  const annotsName = PDFName.of('Annots');
  const annots = page.node.lookupMaybe(annotsName, PDFArray);
  if (!annots) return 0;

  const kept = PDFArray.withContext(pdfDoc.context);
  let removed = 0;
  for (let index = 0; index < annots.size(); index += 1) {
    const entry = annots.get(index);
    let annotation = null;
    try {
      annotation = pdfDoc.context.lookup(entry);
    } catch {
      annotation = null;
    }
    const subtype = annotation?.get?.(PDFName.of('Subtype'));
    if (String(subtype) === '/Link') {
      removed += 1;
    } else {
      kept.push(entry);
    }
  }
  page.node.set(annotsName, kept);
  return removed;
}

function addInternalLinkAnnotation(pdfDoc, sourcePage, rect, targetPageIndex) {
  const normalizedRect = normalizePdfRect(rect);
  if (!normalizedRect) throw new Error('Invalid index link rectangle.');
  const targetPage = pdfDoc.getPage(targetPageIndex);
  if (!targetPage) throw new Error(`Invalid index link destination page ${targetPageIndex + 1}.`);

  const context = pdfDoc.context;
  const destination = PDFArray.withContext(context);
  destination.push(targetPage.ref);
  destination.push(PDFName.of('Fit'));

  const link = context.obj({
    Type: 'Annot',
    Subtype: 'Link',
    Rect: normalizedRect,
    Border: [0, 0, 0],
    H: 'I',
    Dest: destination,
  });
  const linkRef = context.register(link);
  const annotsName = PDFName.of('Annots');
  let annots = sourcePage.node.lookupMaybe(annotsName, PDFArray);
  if (!annots) {
    annots = PDFArray.withContext(context);
    sourcePage.node.set(annotsName, annots);
  }
  annots.push(linkRef);
}

function addTopLevelBookmarks(pdfDoc, bookmarks) {
  if (!bookmarks.length) return;

  const context = pdfDoc.context;
  const outlinesRef = context.nextRef();
  const itemRefs = bookmarks.map(() => context.nextRef());

  bookmarks.forEach((bookmark, index) => {
    const page = pdfDoc.getPage(bookmark.pageIndex);
    const destination = PDFArray.withContext(context);
    destination.push(page.ref);
    destination.push(PDFName.of('Fit'));

    const item = context.obj({
      Title: PDFHexString.fromText(bookmark.title),
      Parent: outlinesRef,
      Prev: index > 0 ? itemRefs[index - 1] : undefined,
      Next: index < itemRefs.length - 1 ? itemRefs[index + 1] : undefined,
      Dest: destination,
    });
    context.assign(itemRefs[index], item);
  });

  const outlines = context.obj({
    Type: 'Outlines',
    First: itemRefs[0],
    Last: itemRefs[itemRefs.length - 1],
    Count: itemRefs.length,
  });
  context.assign(outlinesRef, outlines);
  pdfDoc.catalog.set(PDFName.of('Outlines'), outlinesRef);
}

async function buildAndDownload() {
  if (!buildState.entries.length) return;
  if (els.insertBuildTabs.checked && buildState.tabFiles.length < buildState.entries.length) {
    updateBuildStatusSummary();
    return;
  }
  const indexValidation = getIndexLinkValidation();
  if (!indexValidation.ok) {
    updateBuildStatusSummary();
    return;
  }

  els.buildDownload.disabled = true;
  try {
    const output = await PDFDocument.create();
    const bookmarks = [];
    const contentPageTargets = [];
    let indexPageStart = null;
    let indexPageCount = 0;
    let indexLinksCreated = 0;
    let oldIndexLinksRemoved = 0;

    if (buildState.indexFile) {
      indexPageStart = output.getPageCount();
      setBuildStatus(`Adding index: ${buildState.indexFile.name}`);
      const indexResult = await appendIndexPdfToOutput(output, buildState.indexFile);
      indexPageCount = indexResult.pageCount;
      oldIndexLinksRemoved = indexResult.removedLinks;
    }

    for (let i = 0; i < buildState.entries.length; i += 1) {
      const entry = buildState.entries[i];
      const tabFile = els.insertBuildTabs.checked ? buildState.tabFiles[i] : null;

      if (tabFile) {
        const tabPageIndex = output.getPageCount();
        setBuildStatus(`Adding positional tab ${i + 1}: ${tabFile.file.name}`);
        await appendPdfToOutput(output, tabFile.file);
        if (els.createBuildBookmarks.checked) {
          bookmarks.push({ title: tabFile.bookmarkTitle || `TAB ${i + 1}`, pageIndex: tabPageIndex });
        }
      }

      const contentPageIndex = output.getPageCount();
      contentPageTargets.push(contentPageIndex);
      setBuildStatus(`Adding content ${i + 1} of ${buildState.entries.length}: ${entry.file.name}`);
      await appendPdfToOutput(output, entry.file);
    }

    if (buildState.indexFile && els.createIndexLinks.checked) {
      const rows = buildState.indexInfo?.rows || [];
      setBuildStatus(`Creating ${rows.length} linked index entr${rows.length === 1 ? 'y' : 'ies'}…`);
      rows.forEach((row) => {
        const targetPageIndex = contentPageTargets[row.number - 1];
        const sourcePage = output.getPage(indexPageStart + row.pageIndex);
        if (targetPageIndex == null || !sourcePage) {
          throw new Error(`Could not map index row ${row.number} to its final content page.`);
        }
        addInternalLinkAnnotation(output, sourcePage, row.rect, targetPageIndex);
        indexLinksCreated += 1;
      });
    }

    if (bookmarks.length) {
      setBuildStatus(`Creating ${bookmarks.length} tab bookmark${bookmarks.length === 1 ? '' : 's'}…`);
      addTopLevelBookmarks(output, bookmarks);
    }

    setBuildStatus('Saving rebuilt binder…');
    const bytes = await output.save();
    const filename = sanitizePdfFilename(els.buildOutputName.value || 'Rebuilt Binder.pdf');
    downloadBlob(new Blob([bytes], { type: 'application/pdf' }), filename);
    const indexPhrase = indexPageCount
      ? ` with ${indexPageCount} index page${indexPageCount === 1 ? '' : 's'}${indexLinksCreated ? ` and ${indexLinksCreated} linked index entr${indexLinksCreated === 1 ? 'y' : 'ies'}` : ''}`
      : '';
    const tabPhrase = els.insertBuildTabs.checked ? `, ${buildState.entries.length} positional tab PDF${buildState.entries.length === 1 ? '' : 's'}` : '';
    const bookmarkPhrase = bookmarks.length ? `, and ${bookmarks.length} bookmark${bookmarks.length === 1 ? '' : 's'}` : '';
    const cleanupPhrase = oldIndexLinksRemoved ? ` Existing index link annotation${oldIndexLinksRemoved === 1 ? ' was' : 's were'} replaced.` : '';
    setBuildStatus(`Built ${filename}${indexPhrase}${tabPhrase}${bookmarkPhrase}.${cleanupPhrase}`, 'success');
  } catch (error) {
    console.error(error);
    setBuildStatus(`Could not build the binder: ${friendlyError(error)}`, 'error');
  } finally {
    renderBuildTable();
  }
}

function exportBuildXlsx() {
  if (!buildState.entries.length) return;

  const workbook = XLSX.utils.book_new();
  const indexRowsByNumber = new Map((buildState.indexInfo?.rows || []).map((row) => [row.number, row]));
  const rows = buildState.entries.map((entry, index) => ({
    Order: index + 1,
    'Tab Position': els.insertBuildTabs.checked ? index + 1 : '',
    'Tab Filename': els.insertBuildTabs.checked ? (buildState.tabFiles[index]?.file.name || '') : '',
    'Bookmark Title': els.insertBuildTabs.checked && els.createBuildBookmarks.checked ? (buildState.tabFiles[index]?.bookmarkTitle || '') : '',
    'Index Description': indexRowsByNumber.get(index + 1)?.description || '',
    'Index Date': indexRowsByNumber.get(index + 1)?.date || '',
    Description: entry.description,
    Filename: entry.file.name,
    'File Size': formatBytes(entry.file.size),
  }));
  const sheet = XLSX.utils.json_to_sheet(rows);
  sheet['!cols'] = [{ wch: 8 }, { wch: 14 }, { wch: 26 }, { wch: 28 }, { wch: 50 }, { wch: 16 }, { wch: 44 }, { wch: 50 }, { wch: 14 }];
  XLSX.utils.book_append_sheet(workbook, sheet, 'Manifest');

  const configSheet = XLSX.utils.aoa_to_sheet([
    ['PDF Binder Tool - Final Binder Order'],
    ['Output filename', els.buildOutputName.value || 'Rebuilt Binder.pdf'],
    ['Tab semantics', 'position'],
    ['Tabs inserted', els.insertBuildTabs.checked ? 'Yes' : 'No'],
    ['Tab bookmarks created', els.insertBuildTabs.checked && els.createBuildBookmarks.checked ? 'Yes' : 'No'],
    ['Index PDF', buildState.indexFile?.name || ''],
    ['Index pages', buildState.indexInfo?.pageCount || ''],
    ['Index Description links created', buildState.indexFile && els.createIndexLinks.checked ? 'Yes' : 'No'],
    ['Detected index rows', buildState.indexInfo?.rows?.length || ''],
  ]);
  configSheet['!cols'] = [{ wch: 22 }, { wch: 70 }];
  XLSX.utils.book_append_sheet(workbook, configSheet, 'Binder Settings');

  const bytes = XLSX.write(workbook, { type: 'array', bookType: 'xlsx' });
  downloadBlob(new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), 'final-binder-order.xlsx');
}

function isPdfFile(file) {
  return file.type === 'application/pdf' || file.name.toLowerCase().endsWith('.pdf');
}

function makeId() {
  if (globalThis.crypto?.randomUUID) return crypto.randomUUID();
  return `entry-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function naturalCompare(a, b) {
  return a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' });
}

function humanizeFilename(filename) {
  return stripExtension(filename)
    .replace(/^\s*\d+\s*[-_.]\s*/, '')
    .replace(/[_]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function sanitizePdfFilename(value, maxLength = SAFE_FILENAME_MAX_LENGTH) {
  let raw = String(value || 'section.pdf').normalize('NFC');
  raw = raw.replace(/\.pdf$/i, '');

  let base = raw
    .replace(/(\d):(?=\d)/g, '$1-')
    .replace(/(\d)\/(?=\d)/g, '$1-')
    .replace(/[:/\\]/g, ' - ')
    .replace(/[<>|*\x00-\x1F]/g, '-')
    .replace(/["?]/g, '')
    .replace(/-{2,}/g, '-')
    .replace(/\s+/g, ' ')
    .replace(/\s+-\s+/g, ' - ')
    .trim()
    .replace(/[. -]+$/g, '');

  if (!base || base === '.' || base === '..') base = 'section';
  if (WINDOWS_RESERVED_BASENAMES.test(base)) base = `_${base}`;

  const extension = '.pdf';
  const maxBaseLength = Math.max(1, maxLength - extension.length);
  if ([...base].length > maxBaseLength) {
    const chars = [...base].slice(0, maxBaseLength).join('');
    const lastSpace = chars.lastIndexOf(' ');
    base = lastSpace >= Math.floor(maxBaseLength * 0.72) ? chars.slice(0, lastSpace) : chars;
    base = base.trim().replace(/[. -]+$/g, '');
  }

  if (!base) base = 'section';
  if (WINDOWS_RESERVED_BASENAMES.test(base)) base = `_${base}`;
  return `${base}${extension}`;
}

function stripExtension(filename) {
  return String(filename || '').replace(/\.[^.]+$/, '');
}

function csvEscape(value) {
  const text = String(value ?? '');
  return /[",\r\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = filename;
  document.body.append(anchor);
  anchor.click();
  anchor.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function setSplitStatus(message, type = '') {
  els.splitStatus.textContent = message;
  els.splitStatus.className = `status ${type}`.trim();
}

function setBuildStatus(message, type = '') {
  els.buildStatus.textContent = message;
  els.buildStatus.className = `status ${type}`.trim();
}

function friendlyError(error) {
  return error?.message || String(error);
}

function formatBytes(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 ** 2) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / 1024 ** 2).toFixed(1)} MB`;
}

function escapeHtml(value) {
  return String(value)
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#039;');
}

function escapeAttr(value) {
  return escapeHtml(value).replaceAll('`', '&#096;');
}
