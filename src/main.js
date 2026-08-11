import './styles.css';
import * as pdfjsLib from 'pdfjs-dist';
import pdfjsWorker from 'pdfjs-dist/build/pdf.worker.min.mjs?url';
import { PDFDocument } from 'pdf-lib';
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

  currentSectionEditor: document.querySelector('#current-section-editor'),
  currentSectionPages: document.querySelector('#current-section-pages'),
  currentSectionFilenameMode: document.querySelector('#current-section-filename-mode'),
  currentSectionLabel: document.querySelector('#current-section-label'),
  currentSectionDescription: document.querySelector('#current-section-description'),
  currentSectionFilename: document.querySelector('#current-section-filename'),

  manifestBody: document.querySelector('#manifest-body'),
  importSplitManifest: document.querySelector('#import-split-manifest'),
  exportManifestJson: document.querySelector('#export-manifest-json'),
  exportManifestCsv: document.querySelector('#export-manifest-csv'),
  exportManifestXlsx: document.querySelector('#export-manifest-xlsx'),
  filenameTemplate: document.querySelector('#filename-template'),
  applyFilenameTemplate: document.querySelector('#apply-filename-template'),
  manifestValidationSummary: document.querySelector('#manifest-validation-summary'),

  buildFiles: document.querySelector('#build-files'),
  buildTabs: document.querySelector('#build-tabs'),
  clearBuildTabs: document.querySelector('#clear-build-tabs'),
  insertBuildTabs: document.querySelector('#insert-build-tabs'),
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
};

const buildState = {
  entries: [],
  tabFiles: [],
  manifest: null,
  draggedEntryId: null,
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
els.currentSectionFilename.addEventListener('input', () => updateCurrentSectionField('filename', els.currentSectionFilename.value, true));

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
els.insertBuildTabs.addEventListener('change', () => {
  renderBuildTable();
  updateBuildStatusSummary();
});
els.buildManifest.addEventListener('change', importBuildManifest);
els.buildDownload.addEventListener('click', buildAndDownload);
els.exportBuildXlsx.addEventListener('click', exportBuildXlsx);

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

  els.thumbnailList.innerHTML = 'No PDF loaded.';
  els.thumbnailList.className = 'thumbnail-list empty-state';
  els.previewStage.innerHTML = 'Open a PDF to preview pages.';
  els.previewStage.className = 'preview-stage empty-state';
  els.previewLabel.textContent = 'No page selected.';
  els.manifestBody.innerHTML = '<tr><td colspan="7" class="empty-cell">No divisions yet.</td></tr>';
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
  els.exportManifestJson.disabled = !enabled;
  els.exportManifestCsv.disabled = !enabled;
  els.exportManifestXlsx.disabled = !enabled;
  els.applyFilenameTemplate.disabled = !enabled;
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

function updateMarkerButtons() {
  if (!splitState.pdfjsDoc) return;
  const pageIndex = splitState.selectedPage;
  const isLastPage = pageIndex === splitState.pdfjsDoc.numPages - 1;

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
    const startIndex = raw.pageIndexes[0];
    const endIndex = raw.pageIndexes[raw.pageIndexes.length - 1];

    const filenameContext = {
      order,
      label,
      description,
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

  updateManifestValidationSummary(validation, sections.length);

  if (!sections.length) {
    els.manifestBody.innerHTML = '<tr><td colspan="7" class="empty-cell">No output sections. Add a normal split cut, detect tabs, or mark divider pages manually.</td></tr>';
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
  els.currentSectionFilename.value = section.filename;
  els.currentSectionFilenameMode.textContent = section.filenameLocked ? 'manual filename' : 'auto filename';
  els.currentSectionFilenameMode.className = `filename-mode${section.filenameLocked ? ' manual' : ''}`;
}

function syncCurrentSectionEditorValues() {
  const section = getSectionByKey(splitState.currentSectionKey);
  if (!section) return;
  if (document.activeElement !== els.currentSectionLabel) els.currentSectionLabel.value = section.label;
  if (document.activeElement !== els.currentSectionDescription) els.currentSectionDescription.value = section.description;
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
    schema: 'pdf-binder-tool/5',
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
  const headers = ['Order', 'Boundary Type', 'Boundary Page', 'Tab Page (Excluded)', 'Split After Page', 'Start Page', 'End Page', 'Page Count', 'Label', 'Description', 'Filename', 'Filename Mode', 'Tab Filename'];
  const rows = sections.map((s) => [s.order, s.boundaryType, s.boundaryPage, s.tabPage, s.splitAfterPage, s.startPage, s.endPage, s.pageCount, s.label, s.description, s.filename, s.filenameMode, s.tabFilename]);
  return [headers, ...rows].map((row) => row.map(csvEscape).join(',')).join('\r\n');
}

function makeManifestWorkbookBytes(manifest) {
  const workbook = XLSX.utils.book_new();
  const rows = manifest.sections.map((section) => ({
    Order: section.order,
    Label: section.label,
    Description: section.description,
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
    { wch: 8 }, { wch: 18 }, { wch: 38 }, { wch: 44 }, { wch: 14 },
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
    ['You may edit Order, Label, Description, Filename, and Tab Filename for binder planning.'],
    ['When importing this workbook back into an already-open source PDF, the app uses Label, Description, and Filename only.'],
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
  };
}

function addBuildTabFiles(files) {
  const pdfs = files.filter(isPdfFile);
  buildState.tabFiles.push(...pdfs);
  applyBuildTabOrder();
  renderBuildTable();
  updateBuildStatusSummary();
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
    const aOrder = manifestTabOrder.get(a.name);
    const bOrder = manifestTabOrder.get(b.name);
    if (aOrder == null && bOrder == null) return naturalCompare(a.name, b.name);
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
  const loaded = new Set(buildState.tabFiles.map((file) => file.name));
  return buildState.manifest.sections.filter((section) => section.tabFilename && loaded.has(section.tabFilename)).length;
}

function renderBuildTable() {
  els.clearBuildTabs.disabled = buildState.tabFiles.length === 0;
  els.exportBuildXlsx.disabled = buildState.entries.length === 0;
  const tabsEnabled = els.insertBuildTabs.checked;
  const hasEnoughTabs = !tabsEnabled || buildState.tabFiles.length >= buildState.entries.length;
  els.buildDownload.disabled = buildState.entries.length === 0 || !hasEnoughTabs;

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
        ? `<span class="tab-file-name">${escapeHtml(tabFile.name)}</span>`
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
  if (!contentCount) {
    setBuildStatus(tabCount
      ? `${tabCount} positional tab PDF${tabCount === 1 ? '' : 's'} loaded. Add content PDFs to build the binder.`
      : 'Add content PDFs in the order you want, or import a manifest.');
    return;
  }

  const duplicateNames = findDuplicateNames(buildState.entries.map((entry) => entry.file.name));
  const duplicateNote = duplicateNames.length ? ` Warning: duplicate content filename${duplicateNames.length === 1 ? '' : 's'} (${duplicateNames.join(', ')}) can make manifest matching ambiguous.` : '';

  if (!els.insertBuildTabs.checked) {
    setBuildStatus(`${contentCount} content PDF${contentCount === 1 ? '' : 's'} loaded. Positional tab insertion is off.${duplicateNote}`, duplicateNames.length ? 'error' : '');
    return;
  }

  if (tabCount < contentCount) {
    const missing = contentCount - tabCount;
    setBuildStatus(`${contentCount} content PDF${contentCount === 1 ? '' : 's'} and ${tabCount} tab PDF${tabCount === 1 ? '' : 's'} loaded. Add ${missing} more tab PDF${missing === 1 ? '' : 's'} or turn off tab insertion.${duplicateNote}`, 'error');
    return;
  }

  const extra = tabCount - contentCount;
  const extraNote = extra ? ` ${extra} extra tab PDF${extra === 1 ? '' : 's'} will not be used.` : '';
  setBuildStatus(`${contentCount} content PDF${contentCount === 1 ? '' : 's'} will be paired with the first ${contentCount} positional tab PDF${contentCount === 1 ? '' : 's'}.${extraNote}${duplicateNote}`, duplicateNames.length ? 'error' : 'success');
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
}

async function buildAndDownload() {
  if (!buildState.entries.length) return;
  if (els.insertBuildTabs.checked && buildState.tabFiles.length < buildState.entries.length) {
    updateBuildStatusSummary();
    return;
  }

  els.buildDownload.disabled = true;
  try {
    const output = await PDFDocument.create();
    for (let i = 0; i < buildState.entries.length; i += 1) {
      const entry = buildState.entries[i];
      const tabFile = els.insertBuildTabs.checked ? buildState.tabFiles[i] : null;

      if (tabFile) {
        setBuildStatus(`Adding positional tab ${i + 1}: ${tabFile.name}`);
        await appendPdfToOutput(output, tabFile);
      }

      setBuildStatus(`Adding content ${i + 1} of ${buildState.entries.length}: ${entry.file.name}`);
      await appendPdfToOutput(output, entry.file);
    }

    setBuildStatus('Saving rebuilt binder…');
    const bytes = await output.save();
    const filename = sanitizePdfFilename(els.buildOutputName.value || 'Rebuilt Binder.pdf');
    downloadBlob(new Blob([bytes], { type: 'application/pdf' }), filename);
    const tabPhrase = els.insertBuildTabs.checked ? ` with ${buildState.entries.length} positional tab PDF${buildState.entries.length === 1 ? '' : 's'}` : '';
    setBuildStatus(`Built ${filename} from ${buildState.entries.length} content PDF${buildState.entries.length === 1 ? '' : 's'}${tabPhrase}.`, 'success');
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
  const rows = buildState.entries.map((entry, index) => ({
    Order: index + 1,
    'Tab Position': els.insertBuildTabs.checked ? index + 1 : '',
    'Tab Filename': els.insertBuildTabs.checked ? (buildState.tabFiles[index]?.name || '') : '',
    Description: entry.description,
    Filename: entry.file.name,
    'File Size': formatBytes(entry.file.size),
  }));
  const sheet = XLSX.utils.json_to_sheet(rows);
  sheet['!cols'] = [{ wch: 8 }, { wch: 14 }, { wch: 26 }, { wch: 44 }, { wch: 50 }, { wch: 14 }];
  XLSX.utils.book_append_sheet(workbook, sheet, 'Manifest');

  const configSheet = XLSX.utils.aoa_to_sheet([
    ['PDF Binder Tool - Final Binder Order'],
    ['Output filename', els.buildOutputName.value || 'Rebuilt Binder.pdf'],
    ['Tab semantics', 'position'],
    ['Tabs inserted', els.insertBuildTabs.checked ? 'Yes' : 'No'],
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

function sanitizePdfFilename(value) {
  let name = String(value || 'section.pdf')
    .replace(/[<>:"/\\|?*\x00-\x1F]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
  if (!name.toLowerCase().endsWith('.pdf')) name += '.pdf';
  return name || 'section.pdf';
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
