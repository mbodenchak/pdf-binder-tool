import './styles.css';
import * as pdfjsLib from 'pdfjs-dist';
import pdfjsWorker from 'pdfjs-dist/build/pdf.worker.min.mjs?url';
import { PDFArray, PDFDocument, PDFHexString, PDFName } from 'pdf-lib';
import JSZip from 'jszip';
import * as XLSX from 'xlsx';

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfjsWorker;

async function createOutputPdfDocument() {
  // pdf-lib normally writes its own Creator/Producer metadata when a new
  // document is created. Disable that automatic metadata, while retaining
  // ordinary creation/modification timestamps for the generated file.
  const pdfDoc = await PDFDocument.create({ updateMetadata: false });
  const now = new Date();
  pdfDoc.setCreationDate(now);
  pdfDoc.setModificationDate(now);
  return pdfDoc;
}

const els = {
  modeTabs: [...document.querySelectorAll('.mode-tab')],
  splitMode: document.querySelector('#split-mode'),
  buildMode: document.querySelector('#build-mode'),
  insertMode: document.querySelector('#insert-mode'),
  reorderMode: document.querySelector('#reorder-mode'),

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
  bulkIndexFile: document.querySelector('#bulk-index-file'),
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
  manualBuildIndex: document.querySelector('#manual-build-index'),
  matchBuildIndex: document.querySelector('#match-build-index'),
  buildMatchDialog: document.querySelector('#build-match-dialog'),
  closeBuildMatch: document.querySelector('#close-build-match'),
  cancelBuildMatch: document.querySelector('#cancel-build-match'),
  rerunBuildMatch: document.querySelector('#rerun-build-match'),
  applyBuildMatch: document.querySelector('#apply-build-match'),
  buildMatchSummary: document.querySelector('#build-match-summary'),
  buildMatchBody: document.querySelector('#build-match-body'),
  insertBuildTabs: document.querySelector('#insert-build-tabs'),
  createBuildBookmarks: document.querySelector('#create-build-bookmarks'),
  buildManifest: document.querySelector('#build-manifest'),
  buildOutputName: document.querySelector('#build-output-name'),
  buildDownload: document.querySelector('#build-download'),
  buildStatus: document.querySelector('#build-status'),
  buildBody: document.querySelector('#build-body'),
  exportBuildXlsx: document.querySelector('#export-build-xlsx'),

  insertBinder: document.querySelector('#insert-binder'),
  insertFiles: document.querySelector('#insert-files'),
  insertTabs: document.querySelector('#insert-tabs'),
  clearInsertTabs: document.querySelector('#clear-insert-tabs'),
  insertIndex: document.querySelector('#insert-index'),
  clearInsertIndex: document.querySelector('#clear-insert-index'),
  insertIndexSummary: document.querySelector('#insert-index-summary'),
  insertIndexBody: document.querySelector('#insert-index-body'),
  manualInsertIndex: document.querySelector('#manual-insert-index'),
  insertCreateBookmarks: document.querySelector('#insert-create-bookmarks'),
  insertRefreshIndexLinks: document.querySelector('#insert-refresh-index-links'),
  insertOutputName: document.querySelector('#insert-output-name'),
  insertDownload: document.querySelector('#insert-download'),
  insertStatus: document.querySelector('#insert-status'),
  insertAnalysisSummary: document.querySelector('#insert-analysis-summary'),
  insertStructureNote: document.querySelector('#insert-structure-note'),
  insertOperationsBody: document.querySelector('#insert-operations-body'),
  insertPreviewBody: document.querySelector('#insert-preview-body'),

  reorderBinder: document.querySelector('#reorder-binder'),
  reorderIndex: document.querySelector('#reorder-index'),
  clearReorderIndex: document.querySelector('#clear-reorder-index'),
  resetReorderOrder: document.querySelector('#reset-reorder-order'),
  reorderCreateBookmarks: document.querySelector('#reorder-create-bookmarks'),
  reorderRefreshIndexLinks: document.querySelector('#reorder-refresh-index-links'),
  reorderOutputName: document.querySelector('#reorder-output-name'),
  reorderDownload: document.querySelector('#reorder-download'),
  reorderStatus: document.querySelector('#reorder-status'),
  reorderAnalysisSummary: document.querySelector('#reorder-analysis-summary'),
  reorderStructureNote: document.querySelector('#reorder-structure-note'),
  reorderIndexSummary: document.querySelector('#reorder-index-summary'),
  reorderIndexBody: document.querySelector('#reorder-index-body'),
  manualReorderIndex: document.querySelector('#manual-reorder-index'),
  reorderBody: document.querySelector('#reorder-body'),

  syncMode: document.querySelector('#sync-mode'),
  syncBinder: document.querySelector('#sync-binder'),
  syncIndex: document.querySelector('#sync-index'),
  syncFiles: document.querySelector('#sync-files'),
  clearSyncFiles: document.querySelector('#clear-sync-files'),
  syncTabs: document.querySelector('#sync-tabs'),
  clearSyncTabs: document.querySelector('#clear-sync-tabs'),
  syncCreateBookmarks: document.querySelector('#sync-create-bookmarks'),
  syncOutputName: document.querySelector('#sync-output-name'),
  syncDownload: document.querySelector('#sync-download'),
  syncStatus: document.querySelector('#sync-status'),
  syncAnalysisSummary: document.querySelector('#sync-analysis-summary'),
  syncStructureNote: document.querySelector('#sync-structure-note'),
  syncRematch: document.querySelector('#sync-rematch'),
  syncBody: document.querySelector('#sync-body'),
  syncOmissions: document.querySelector('#sync-omissions'),
  syncApproveOmissions: document.querySelector('#sync-approve-omissions'),
  syncOmissionsList: document.querySelector('#sync-omissions-list'),

  manualIndexDialog: document.querySelector('#manual-index-dialog'),
  closeManualIndex: document.querySelector('#close-manual-index'),
  cancelManualIndex: document.querySelector('#cancel-manual-index'),
  applyManualIndex: document.querySelector('#apply-manual-index'),
  manualIndexPrev: document.querySelector('#manual-index-prev'),
  manualIndexNext: document.querySelector('#manual-index-next'),
  manualIndexPageLabel: document.querySelector('#manual-index-page-label'),
  manualIndexDraw: document.querySelector('#manual-index-draw'),
  manualIndexReset: document.querySelector('#manual-index-reset'),
  manualIndexStatus: document.querySelector('#manual-index-status'),
  manualIndexCanvas: document.querySelector('#manual-index-canvas'),
  manualIndexOverlay: document.querySelector('#manual-index-overlay'),
  manualIndexMapBody: document.querySelector('#manual-index-map-body'),
  manualIndexSummary: document.querySelector('#manual-index-summary'),
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
  manualIndexMappings: null,
  matchAssignments: new Map(),
};

const insertState = {
  file: null,
  bytes: null,
  pageCount: 0,
  frontMatterPageIndexes: [],
  sourceTabs: [],
  sourceSections: [],
  indexRows: [],
  replacementIndexFile: null,
  replacementIndexInfo: null,
  replacementIndexLoading: false,
  manualIndexMappings: null,
  operations: [],
  extraTabs: [],
  analyzing: false,
  error: null,
};

const reorderState = {
  file: null,
  bytes: null,
  pageCount: 0,
  frontMatterPageIndexes: [],
  sourceTabs: [],
  sourceSections: [],
  indexRows: [],
  items: [],
  replacementIndexFile: null,
  replacementIndexInfo: null,
  replacementIndexLoading: false,
  manualIndexMappings: null,
  analyzing: false,
  error: null,
  draggedItemId: null,
};

const syncState = {
  file: null,
  bytes: null,
  pageCount: 0,
  frontMatterPageIndexes: [],
  sourceTabs: [],
  sourceSections: [],
  oldIndexRows: [],
  newIndexFile: null,
  newIndexInfo: null,
  indexLoading: false,
  newFiles: [],
  extraTabs: [],
  assignments: new Map(),
  approveOmissions: false,
  analyzing: false,
  error: null,
};

const manualIndexState = {
  mode: null,
  sourceKey: '',
  loadingTask: null,
  pdf: null,
  pageIndexes: [],
  pageCursor: 0,
  mappings: [],
  automaticMappings: [],
  viewport: null,
  textItems: [],
  drawMode: false,
  dragStart: null,
  drawElement: null,
  selectedId: null,
};

els.modeTabs.forEach((tab) => {
  tab.addEventListener('click', () => {
    const mode = tab.dataset.mode;
    els.modeTabs.forEach((item) => item.classList.toggle('active', item === tab));
    els.splitMode.classList.toggle('active', mode === 'split');
    els.buildMode.classList.toggle('active', mode === 'build');
    els.insertMode.classList.toggle('active', mode === 'insert');
    els.reorderMode.classList.toggle('active', mode === 'reorder');
    els.syncMode.classList.toggle('active', mode === 'sync');
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
els.bulkIndexFile.addEventListener('change', handleBulkIndexFile);
els.bulkRegenerateFilenames.addEventListener('change', renderBulkPastePreview);
els.applyBulkPaste.addEventListener('click', applyBulkPasteEntries);
els.cancelBulkPaste.addEventListener('click', closeBulkPasteDialog);
els.closeBulkPaste.addEventListener('click', closeBulkPasteDialog);
els.bulkPasteDialog.addEventListener('cancel', (event) => {
  event.preventDefault();
  closeBulkPasteDialog();
});

els.manualBuildIndex.addEventListener('click', () => openManualIndexDialog('build'));
els.matchBuildIndex.addEventListener('click', openBuildMatchDialog);
els.closeBuildMatch.addEventListener('click', closeBuildMatchDialog);
els.cancelBuildMatch.addEventListener('click', closeBuildMatchDialog);
els.rerunBuildMatch.addEventListener('click', () => { recomputeBuildMatchProposal(); renderBuildMatchDialog(); });
els.applyBuildMatch.addEventListener('click', applyBuildMatchOrder);
els.buildMatchDialog.addEventListener('cancel', (event) => { event.preventDefault(); closeBuildMatchDialog(); });
els.manualInsertIndex.addEventListener('click', () => openManualIndexDialog('insert'));
els.manualReorderIndex.addEventListener('click', () => openManualIndexDialog('reorder'));
els.closeManualIndex.addEventListener('click', closeManualIndexDialog);
els.cancelManualIndex.addEventListener('click', closeManualIndexDialog);
els.applyManualIndex.addEventListener('click', applyManualIndexMappings);
els.manualIndexPrev.addEventListener('click', () => changeManualIndexPage(-1));
els.manualIndexNext.addEventListener('click', () => changeManualIndexPage(1));
els.manualIndexDraw.addEventListener('click', toggleManualIndexDrawMode);
els.manualIndexReset.addEventListener('click', resetManualIndexMappingsToAutomatic);
els.manualIndexDialog.addEventListener('cancel', (event) => { event.preventDefault(); closeManualIndexDialog(); });
els.manualIndexOverlay.addEventListener('pointerdown', handleManualIndexPointerDown);
els.manualIndexOverlay.addEventListener('pointermove', handleManualIndexPointerMove);
els.manualIndexOverlay.addEventListener('pointerup', handleManualIndexPointerUp);
els.manualIndexOverlay.addEventListener('pointercancel', cancelManualIndexDraw);

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
  buildState.manualIndexMappings = null;
  buildState.matchAssignments = new Map();
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

els.insertBinder.addEventListener('change', async (event) => {
  const [file] = event.target.files;
  if (file) await loadInsertBinder(file);
  event.target.value = '';
});
els.insertFiles.addEventListener('change', (event) => {
  addInsertOperations([...event.target.files]);
  event.target.value = '';
});
els.insertTabs.addEventListener('change', async (event) => {
  await addInsertExtraTabs([...event.target.files]);
  event.target.value = '';
});
els.clearInsertTabs.addEventListener('click', () => {
  insertState.extraTabs = [];
  renderInsertWorkspace();
});
els.insertIndex.addEventListener('change', async (event) => {
  const [file] = event.target.files;
  if (file) await loadInsertReplacementIndex(file);
  event.target.value = '';
});
els.clearInsertIndex.addEventListener('click', () => {
  insertState.replacementIndexFile = null;
  insertState.replacementIndexInfo = null;
  insertState.replacementIndexLoading = false;
  insertState.manualIndexMappings = null;
  renderInsertWorkspace();
});
els.insertCreateBookmarks.addEventListener('change', renderInsertWorkspace);
els.insertRefreshIndexLinks.addEventListener('change', renderInsertWorkspace);
els.insertDownload.addEventListener('click', buildInsertedBinder);

els.reorderBinder.addEventListener('change', async (event) => {
  const [file] = event.target.files;
  if (file) await loadReorderBinder(file);
  event.target.value = '';
});
els.reorderIndex.addEventListener('change', async (event) => {
  const [file] = event.target.files;
  if (file) await loadReorderReplacementIndex(file);
  event.target.value = '';
});
els.clearReorderIndex.addEventListener('click', () => {
  reorderState.replacementIndexFile = null;
  reorderState.replacementIndexInfo = null;
  reorderState.replacementIndexLoading = false;
  reorderState.manualIndexMappings = null;
  renderReorderWorkspace();
});
els.resetReorderOrder.addEventListener('click', resetReorderOrder);
els.reorderCreateBookmarks.addEventListener('change', renderReorderWorkspace);
els.reorderRefreshIndexLinks.addEventListener('change', renderReorderWorkspace);
els.reorderDownload.addEventListener('click', buildReorderedBinder);

els.syncBinder.addEventListener('change', async (event) => {
  const [file] = event.target.files;
  if (file) await loadSyncBinder(file);
  event.target.value = '';
});
els.syncIndex.addEventListener('change', async (event) => {
  const [file] = event.target.files;
  if (file) await loadSyncIndex(file);
  event.target.value = '';
});
els.syncFiles.addEventListener('change', (event) => {
  addSyncFiles([...event.target.files]);
  event.target.value = '';
});
els.clearSyncFiles.addEventListener('click', () => {
  syncState.newFiles = [];
  recomputeSyncProposal({ preserveManual: false });
  renderSyncWorkspace();
});
els.syncTabs.addEventListener('change', async (event) => {
  await addSyncTabs([...event.target.files]);
  event.target.value = '';
});
els.clearSyncTabs.addEventListener('click', () => {
  syncState.extraTabs = [];
  renderSyncWorkspace();
});
els.syncRematch.addEventListener('click', () => {
  recomputeSyncProposal({ preserveManual: false });
  renderSyncWorkspace();
});
els.syncApproveOmissions.addEventListener('change', () => {
  syncState.approveOmissions = els.syncApproveOmissions.checked;
  renderSyncWorkspace();
});
els.syncCreateBookmarks.addEventListener('change', renderSyncWorkspace);
els.syncDownload.addEventListener('click', buildSyncedBinder);



function getManualIndexModeState(mode) {
  if (mode === 'build') return buildState;
  if (mode === 'insert') return insertState;
  if (mode === 'reorder') return reorderState;
  return null;
}

function manualIndexSourceKeyForFile(file, prefix = 'file') {
  if (!file) return '';
  return `${prefix}:${file.name}:${file.size}:${file.lastModified || 0}`;
}

function getManualIndexContext(mode) {
  if (mode === 'build') {
    if (!buildState.indexFile || buildState.indexLoading) return null;
    const pageCount = buildState.indexInfo?.pageCount || 0;
    return {
      mode,
      title: buildState.indexFile.name,
      sourceKey: manualIndexSourceKeyForFile(buildState.indexFile, 'build-index'),
      file: buildState.indexFile,
      data: null,
      pageIndexes: Array.from({ length: pageCount }, (_, index) => index),
      automaticRows: buildState.indexInfo?.rows || [],
      destinationLabels: buildState.entries.map((entry) => entry.description || entry.file.name),
    };
  }

  if (mode === 'insert') {
    const plan = getInsertPlan();
    if (!insertState.file || insertState.analyzing) return null;
    if (insertState.replacementIndexFile) {
      if (insertState.replacementIndexLoading) return null;
      const pageCount = insertState.replacementIndexInfo?.pageCount || 0;
      return {
        mode,
        title: insertState.replacementIndexFile.name,
        sourceKey: manualIndexSourceKeyForFile(insertState.replacementIndexFile, 'insert-replacement-index'),
        file: insertState.replacementIndexFile,
        data: null,
        pageIndexes: Array.from({ length: pageCount }, (_, index) => index),
        automaticRows: insertState.replacementIndexInfo?.rows || [],
        destinationLabels: plan.items.map((item) => item.description || `Position ${plan.items.indexOf(item) + 1}`),
      };
    }
    const pages = insertState.frontMatterPageIndexes || [];
    if (!insertState.bytes || !pages.length) return null;
    return {
      mode,
      title: `${insertState.file.name} — existing front matter`,
      sourceKey: `insert-existing:${insertState.file.name}:${insertState.file.size}:${insertState.file.lastModified || 0}`,
      file: null,
      data: insertState.bytes.slice(),
      pageIndexes: [...pages],
      automaticRows: insertState.indexRows || [],
      destinationLabels: plan.items.map((item) => item.description || `Position ${plan.items.indexOf(item) + 1}`),
    };
  }

  if (mode === 'reorder') {
    if (!reorderState.file || reorderState.analyzing) return null;
    if (reorderState.replacementIndexFile) {
      if (reorderState.replacementIndexLoading) return null;
      const pageCount = reorderState.replacementIndexInfo?.pageCount || 0;
      return {
        mode,
        title: reorderState.replacementIndexFile.name,
        sourceKey: manualIndexSourceKeyForFile(reorderState.replacementIndexFile, 'reorder-replacement-index'),
        file: reorderState.replacementIndexFile,
        data: null,
        pageIndexes: Array.from({ length: pageCount }, (_, index) => index),
        automaticRows: reorderState.replacementIndexInfo?.rows || [],
        destinationLabels: reorderState.items.map((item) => item.description || `Position ${reorderState.items.indexOf(item) + 1}`),
      };
    }
    const pages = reorderState.frontMatterPageIndexes || [];
    if (!reorderState.bytes || !pages.length) return null;
    return {
      mode,
      title: `${reorderState.file.name} — existing front matter`,
      sourceKey: `reorder-existing:${reorderState.file.name}:${reorderState.file.size}:${reorderState.file.lastModified || 0}`,
      file: null,
      data: reorderState.bytes.slice(),
      pageIndexes: [...pages],
      automaticRows: reorderState.indexRows || [],
      destinationLabels: reorderState.items.map((item) => item.description || `Position ${reorderState.items.indexOf(item) + 1}`),
    };
  }
  return null;
}

function mappingFromAutomaticIndexRow(row, destinationCount) {
  const destination = Number.isInteger(row.number) && row.number >= 1 && row.number <= destinationCount
    ? row.number
    : null;
  return {
    id: makeId(),
    pageIndex: row.pageIndex,
    rect: [...row.rect],
    description: row.description || `Detected row ${row.number || ''}`.trim(),
    date: row.date || '',
    type: destination ? 'link' : 'ignore',
    destinationPosition: destination,
    source: row.rectSource === 'existing-link' ? 'existing-link' : 'automatic',
    originalNumber: row.number || null,
  };
}

function cloneManualMappings(mappings) {
  return (mappings || []).map((mapping) => ({
    ...mapping,
    rect: Array.isArray(mapping.rect) ? [...mapping.rect] : null,
  }));
}

async function openManualIndexDialog(mode) {
  const context = getManualIndexContext(mode);
  if (!context || !context.pageIndexes.length) {
    const message = 'Load an index/front matter source before opening manual mapping.';
    if (mode === 'build') setBuildStatus(message, 'error');
    else if (mode === 'insert') setInsertStatus(message, 'error');
    else setReorderStatus(message, 'error');
    return;
  }

  await closeManualIndexPdf();
  manualIndexState.mode = mode;
  manualIndexState.sourceKey = context.sourceKey;
  manualIndexState.pageIndexes = [...context.pageIndexes];
  manualIndexState.pageCursor = 0;
  manualIndexState.drawMode = false;
  manualIndexState.dragStart = null;
  manualIndexState.selectedId = null;
  manualIndexState.viewport = null;
  manualIndexState.textItems = [];

  const automatic = context.automaticRows.map((row) => mappingFromAutomaticIndexRow(row, context.destinationLabels.length));
  manualIndexState.automaticMappings = cloneManualMappings(automatic);
  const stored = getManualIndexModeState(mode)?.manualIndexMappings;
  const hasStoredMapping = stored?.sourceKey === context.sourceKey;
  manualIndexState.mappings = hasStoredMapping
    ? cloneManualMappings(stored.mappings)
    : cloneManualMappings(automatic);

  els.manualIndexStatus.textContent = `Loading ${context.title}…`;
  els.manualIndexDraw.classList.remove('active');
  els.manualIndexOverlay.classList.remove('draw-mode');
  if (typeof els.manualIndexDialog.showModal === 'function') els.manualIndexDialog.showModal();
  else els.manualIndexDialog.setAttribute('open', '');

  try {
    let data;
    if (context.file) data = new Uint8Array(await context.file.arrayBuffer());
    else data = context.data;
    manualIndexState.loadingTask = pdfjsLib.getDocument({ data });
    manualIndexState.pdf = await manualIndexState.loadingTask.promise;
    if (!hasStoredMapping) await augmentManualMappingsWithRawLinks(context);
    await renderManualIndexPage();
    renderManualIndexMapTable();
  } catch (error) {
    console.error('Could not open manual index mapping source.', error);
    els.manualIndexStatus.textContent = `Could not open index: ${friendlyError(error)}`;
  }
}

async function closeManualIndexPdf() {
  if (manualIndexState.loadingTask) {
    try { await manualIndexState.loadingTask.destroy(); } catch (error) { console.warn('Could not release manual index PDF resources.', error); }
  }
  manualIndexState.loadingTask = null;
  manualIndexState.pdf = null;
}

function closeManualIndexDialog() {
  closeManualIndexPdf();
  cancelManualIndexDraw();
  if (typeof els.manualIndexDialog.close === 'function' && els.manualIndexDialog.open) els.manualIndexDialog.close();
  else els.manualIndexDialog.removeAttribute('open');
}

async function changeManualIndexPage(delta) {
  const next = manualIndexState.pageCursor + delta;
  if (next < 0 || next >= manualIndexState.pageIndexes.length) return;
  manualIndexState.pageCursor = next;
  manualIndexState.selectedId = null;
  cancelManualIndexDraw();
  await renderManualIndexPage();
  renderManualIndexMapTable();
}


function rectsSubstantiallyOverlap(a, b) {
  const ra = normalizePdfRect(a);
  const rb = normalizePdfRect(b);
  if (!ra || !rb) return false;
  const left = Math.max(ra[0], rb[0]);
  const bottom = Math.max(ra[1], rb[1]);
  const right = Math.min(ra[2], rb[2]);
  const top = Math.min(ra[3], rb[3]);
  if (right <= left || top <= bottom) return false;
  const intersection = (right - left) * (top - bottom);
  const areaA = Math.max(1, (ra[2] - ra[0]) * (ra[3] - ra[1]));
  const areaB = Math.max(1, (rb[2] - rb[0]) * (rb[3] - rb[1]));
  return intersection / Math.min(areaA, areaB) >= 0.45;
}

async function augmentManualMappingsWithRawLinks(context) {
  if (!manualIndexState.pdf) return;
  const pending = [];
  for (const pageIndex of manualIndexState.pageIndexes) {
    const page = await manualIndexState.pdf.getPage(pageIndex + 1);
    let annotations = [];
    try { annotations = await page.getAnnotations({ intent: 'any' }); } catch { annotations = []; }
    const linkRects = annotations
      .filter((annotation) => annotation?.subtype === 'Link' && Array.isArray(annotation.rect))
      .map((annotation) => normalizePdfRect(annotation.rect))
      .filter(Boolean);
    if (!linkRects.length) continue;
    const textContent = await page.getTextContent();
    const textItems = textContent.items.map(normalizeIndexTextItem).filter(Boolean);
    linkRects.forEach((rect) => {
      const duplicate = manualIndexState.mappings.some((mapping) => mapping.pageIndex === pageIndex && rectsSubstantiallyOverlap(mapping.rect, rect));
      if (duplicate) return;
      const text = joinIndexTextItems(textItems.filter((item) => {
        const cx = item.x + item.width / 2;
        const cy = item.y;
        return cx >= rect[0] && cx <= rect[2] && cy >= rect[1] && cy <= rect[3];
      })) || 'Existing link region';
      pending.push({ pageIndex, rect, description: text });
    });
  }
  pending.sort((a, b) => a.pageIndex - b.pageIndex || b.rect[3] - a.rect[3] || a.rect[0] - b.rect[0]);
  const used = new Set(manualIndexState.mappings.filter((mapping) => mapping.type === 'link').map((mapping) => mapping.destinationPosition).filter(Boolean));
  pending.forEach((entry) => {
    const firstUnused = context.destinationLabels.findIndex((_, index) => !used.has(index + 1));
    const destinationPosition = firstUnused >= 0 ? firstUnused + 1 : null;
    if (destinationPosition) used.add(destinationPosition);
    manualIndexState.mappings.push({
      id: makeId(),
      pageIndex: entry.pageIndex,
      rect: entry.rect,
      description: entry.description,
      date: '',
      type: 'link',
      destinationPosition,
      source: 'existing-link-unmapped',
      originalNumber: null,
    });
  });
}

async function renderManualIndexPage() {
  if (!manualIndexState.pdf || !manualIndexState.pageIndexes.length) return;
  const sourcePageIndex = manualIndexState.pageIndexes[manualIndexState.pageCursor];
  const page = await manualIndexState.pdf.getPage(sourcePageIndex + 1);
  const baseViewport = page.getViewport({ scale: 1 });
  const scale = Math.min(1.55, Math.max(0.8, 900 / Math.max(1, baseViewport.width)));
  const viewport = page.getViewport({ scale });
  manualIndexState.viewport = viewport;

  const canvas = els.manualIndexCanvas;
  const context = canvas.getContext('2d', { alpha: false });
  canvas.width = Math.ceil(viewport.width);
  canvas.height = Math.ceil(viewport.height);
  canvas.style.width = `${viewport.width}px`;
  canvas.style.height = `${viewport.height}px`;
  els.manualIndexPreview.style.width = `${viewport.width}px`;
  els.manualIndexPreview.style.height = `${viewport.height}px`;
  els.manualIndexOverlay.style.width = `${viewport.width}px`;
  els.manualIndexOverlay.style.height = `${viewport.height}px`;

  await page.render({ canvasContext: context, viewport }).promise;
  const textContent = await page.getTextContent();
  manualIndexState.textItems = textContent.items.map(normalizeIndexTextItem).filter(Boolean);

  els.manualIndexPageLabel.textContent = `Page ${manualIndexState.pageCursor + 1} of ${manualIndexState.pageIndexes.length} (source PDF page ${sourcePageIndex + 1})`;
  els.manualIndexPrev.disabled = manualIndexState.pageCursor === 0;
  els.manualIndexNext.disabled = manualIndexState.pageCursor === manualIndexState.pageIndexes.length - 1;
  els.manualIndexStatus.textContent = `${manualIndexState.mappings.filter((mapping) => mapping.pageIndex === sourcePageIndex).length} mapped region${manualIndexState.mappings.filter((mapping) => mapping.pageIndex === sourcePageIndex).length === 1 ? '' : 's'} on this page.`;
  renderManualIndexOverlay();
}

function manualRectToCss(rect) {
  const viewport = manualIndexState.viewport;
  if (!viewport || !rect) return null;
  const converted = viewport.convertToViewportRectangle(rect);
  const widthRatio = els.manualIndexOverlay.clientWidth / viewport.width || 1;
  const heightRatio = els.manualIndexOverlay.clientHeight / viewport.height || 1;
  const left = Math.min(converted[0], converted[2]) * widthRatio;
  const top = Math.min(converted[1], converted[3]) * heightRatio;
  const right = Math.max(converted[0], converted[2]) * widthRatio;
  const bottom = Math.max(converted[1], converted[3]) * heightRatio;
  return { left, top, width: Math.max(2, right - left), height: Math.max(2, bottom - top) };
}

function renderManualIndexOverlay() {
  els.manualIndexOverlay.innerHTML = '';
  const pageIndex = manualIndexState.pageIndexes[manualIndexState.pageCursor];
  manualIndexState.mappings.filter((mapping) => mapping.pageIndex === pageIndex).forEach((mapping) => {
    const box = manualRectToCss(mapping.rect);
    if (!box) return;
    const div = document.createElement('div');
    div.className = `manual-link-region ${mapping.type}${mapping.id === manualIndexState.selectedId ? ' selected' : ''}`;
    div.style.left = `${box.left}px`;
    div.style.top = `${box.top}px`;
    div.style.width = `${box.width}px`;
    div.style.height = `${box.height}px`;
    div.title = mapping.type === 'link'
      ? `${mapping.description || 'Link region'} → Position ${mapping.destinationPosition || '?'}`
      : `${mapping.type}: ${mapping.description || 'region'}`;
    div.addEventListener('click', (event) => {
      if (manualIndexState.drawMode) return;
      event.stopPropagation();
      manualIndexState.selectedId = mapping.id;
      renderManualIndexOverlay();
      renderManualIndexMapTable();
    });
    els.manualIndexOverlay.append(div);
  });
}

function renderManualIndexMapTable() {
  const context = getManualIndexContext(manualIndexState.mode);
  if (!context) return;
  const mappings = manualIndexState.mappings;
  if (!mappings.length) {
    els.manualIndexMapBody.innerHTML = '<tr><td colspan="5" class="empty-cell">No regions detected. Use Draw link region to create one.</td></tr>';
  } else {
    els.manualIndexMapBody.innerHTML = '';
    mappings.forEach((mapping) => {
      const tr = document.createElement('tr');
      if (mapping.id === manualIndexState.selectedId) tr.classList.add('selected');
      const typeOptions = ['link', 'heading', 'ignore'].map((type) => `<option value="${type}"${mapping.type === type ? ' selected' : ''}>${type === 'link' ? 'Link' : type === 'heading' ? 'Heading' : 'Ignore'}</option>`).join('');
      const destinationOptions = ['<option value="">Choose…</option>', ...context.destinationLabels.map((label, index) => `<option value="${index + 1}"${mapping.destinationPosition === index + 1 ? ' selected' : ''}>${index + 1} — ${escapeHtml(String(label || `Position ${index + 1}`).slice(0, 70))}</option>`)].join('');
      tr.innerHTML = `
        <td>${mapping.pageIndex + 1}</td>
        <td><span class="manual-map-description">${escapeHtml(mapping.description || '(drawn region)')}</span><span class="manual-map-source">${escapeHtml(mapping.source || 'manual')}</span></td>
        <td><select data-role="type">${typeOptions}</select></td>
        <td><select data-role="destination" ${mapping.type === 'link' ? '' : 'disabled'}>${destinationOptions}</select></td>
        <td><button type="button" data-role="remove" title="Remove region">×</button></td>`;
      tr.addEventListener('click', (event) => {
        if (event.target.closest('select,button')) return;
        manualIndexState.selectedId = mapping.id;
        const pageCursor = manualIndexState.pageIndexes.indexOf(mapping.pageIndex);
        if (pageCursor >= 0 && pageCursor !== manualIndexState.pageCursor) {
          manualIndexState.pageCursor = pageCursor;
          renderManualIndexPage();
        } else renderManualIndexOverlay();
        renderManualIndexMapTable();
      });
      tr.querySelector('[data-role="type"]').addEventListener('change', (event) => {
        mapping.type = event.target.value;
        if (mapping.type === 'link' && !mapping.destinationPosition) mapping.destinationPosition = 1;
        renderManualIndexOverlay();
        renderManualIndexMapTable();
      });
      tr.querySelector('[data-role="destination"]').addEventListener('change', (event) => {
        const value = Number.parseInt(event.target.value, 10);
        mapping.destinationPosition = Number.isInteger(value) ? value : null;
        renderManualIndexOverlay();
        updateManualIndexSummary();
      });
      tr.querySelector('[data-role="remove"]').addEventListener('click', () => {
        manualIndexState.mappings = manualIndexState.mappings.filter((candidate) => candidate.id !== mapping.id);
        if (manualIndexState.selectedId === mapping.id) manualIndexState.selectedId = null;
        renderManualIndexOverlay();
        renderManualIndexMapTable();
      });
      els.manualIndexMapBody.append(tr);
    });
  }
  updateManualIndexSummary();
}

function updateManualIndexSummary() {
  const context = getManualIndexContext(manualIndexState.mode);
  if (!context) return;
  const links = manualIndexState.mappings.filter((mapping) => mapping.type === 'link');
  const invalid = links.filter((mapping) => !Number.isInteger(mapping.destinationPosition) || mapping.destinationPosition < 1 || mapping.destinationPosition > context.destinationLabels.length || !normalizePdfRect(mapping.rect));
  const headings = manualIndexState.mappings.filter((mapping) => mapping.type === 'heading').length;
  const ignored = manualIndexState.mappings.filter((mapping) => mapping.type === 'ignore').length;
  els.manualIndexSummary.textContent = `${links.length} link${links.length === 1 ? '' : 's'}, ${headings} heading${headings === 1 ? '' : 's'}, ${ignored} ignored region${ignored === 1 ? '' : 's'}.${invalid.length ? ` ${invalid.length} link${invalid.length === 1 ? ' needs' : 's need'} a valid destination.` : ' Ready to apply.'}`;
  els.applyManualIndex.disabled = invalid.length > 0 || links.length === 0;
}

function toggleManualIndexDrawMode() {
  manualIndexState.drawMode = !manualIndexState.drawMode;
  els.manualIndexDraw.classList.toggle('active', manualIndexState.drawMode);
  els.manualIndexDraw.textContent = manualIndexState.drawMode ? 'Drawing… drag on page' : 'Draw link region';
  els.manualIndexOverlay.classList.toggle('draw-mode', manualIndexState.drawMode);
  els.manualIndexStatus.textContent = manualIndexState.drawMode ? 'Drag a rectangle over the index text you want clickable.' : 'Draw mode off.';
  cancelManualIndexDraw(false);
}

function resetManualIndexMappingsToAutomatic() {
  manualIndexState.mappings = cloneManualMappings(manualIndexState.automaticMappings);
  manualIndexState.selectedId = null;
  renderManualIndexOverlay();
  renderManualIndexMapTable();
  els.manualIndexStatus.textContent = manualIndexState.mappings.length ? 'Restored automatically detected regions.' : 'No automatic regions were detected; draw regions manually.';
}

function pointerPositionInManualOverlay(event) {
  const rect = els.manualIndexOverlay.getBoundingClientRect();
  return {
    x: Math.max(0, Math.min(rect.width, event.clientX - rect.left)),
    y: Math.max(0, Math.min(rect.height, event.clientY - rect.top)),
  };
}

function handleManualIndexPointerDown(event) {
  if (!manualIndexState.drawMode || !manualIndexState.viewport) return;
  event.preventDefault();
  const point = pointerPositionInManualOverlay(event);
  manualIndexState.dragStart = point;
  const div = document.createElement('div');
  div.className = 'manual-draw-region';
  div.style.left = `${point.x}px`;
  div.style.top = `${point.y}px`;
  div.style.width = '1px';
  div.style.height = '1px';
  manualIndexState.drawElement = div;
  els.manualIndexOverlay.append(div);
  try { els.manualIndexOverlay.setPointerCapture(event.pointerId); } catch {}
}

function handleManualIndexPointerMove(event) {
  if (!manualIndexState.drawMode || !manualIndexState.dragStart || !manualIndexState.drawElement) return;
  const point = pointerPositionInManualOverlay(event);
  const left = Math.min(point.x, manualIndexState.dragStart.x);
  const top = Math.min(point.y, manualIndexState.dragStart.y);
  manualIndexState.drawElement.style.left = `${left}px`;
  manualIndexState.drawElement.style.top = `${top}px`;
  manualIndexState.drawElement.style.width = `${Math.abs(point.x - manualIndexState.dragStart.x)}px`;
  manualIndexState.drawElement.style.height = `${Math.abs(point.y - manualIndexState.dragStart.y)}px`;
}

function handleManualIndexPointerUp(event) {
  if (!manualIndexState.drawMode || !manualIndexState.dragStart || !manualIndexState.viewport) return;
  const end = pointerPositionInManualOverlay(event);
  const start = manualIndexState.dragStart;
  const width = Math.abs(end.x - start.x);
  const height = Math.abs(end.y - start.y);
  cancelManualIndexDraw(false);
  if (width < 8 || height < 8) {
    els.manualIndexStatus.textContent = 'Draw a slightly larger rectangle.';
    return;
  }

  const overlayRect = els.manualIndexOverlay.getBoundingClientRect();
  const sx = manualIndexState.viewport.width / Math.max(1, overlayRect.width);
  const sy = manualIndexState.viewport.height / Math.max(1, overlayRect.height);
  const [pdfX1, pdfY1] = manualIndexState.viewport.convertToPdfPoint(start.x * sx, start.y * sy);
  const [pdfX2, pdfY2] = manualIndexState.viewport.convertToPdfPoint(end.x * sx, end.y * sy);
  const rect = normalizePdfRect([pdfX1, pdfY1, pdfX2, pdfY2]);
  if (!rect) return;

  const selectedText = manualIndexState.textItems.filter((item) => {
    const cx = item.x + item.width / 2;
    const cy = item.y;
    return cx >= rect[0] && cx <= rect[2] && cy >= rect[1] && cy <= rect[3];
  });
  const description = joinIndexTextItems(selectedText) || 'Manual link region';
  const used = new Set(manualIndexState.mappings.filter((mapping) => mapping.type === 'link').map((mapping) => mapping.destinationPosition));
  const context = getManualIndexContext(manualIndexState.mode);
  const firstUnused = context?.destinationLabels.findIndex((_, index) => !used.has(index + 1));
  const destinationPosition = firstUnused != null && firstUnused >= 0 ? firstUnused + 1 : 1;
  const mapping = {
    id: makeId(),
    pageIndex: manualIndexState.pageIndexes[manualIndexState.pageCursor],
    rect,
    description,
    date: '',
    type: 'link',
    destinationPosition,
    source: 'manual-draw',
    originalNumber: null,
  };
  manualIndexState.mappings.push(mapping);
  manualIndexState.selectedId = mapping.id;
  manualIndexState.drawMode = false;
  els.manualIndexDraw.classList.remove('active');
  els.manualIndexDraw.textContent = 'Draw link region';
  els.manualIndexOverlay.classList.remove('draw-mode');
  renderManualIndexOverlay();
  renderManualIndexMapTable();
  els.manualIndexStatus.textContent = `Added region: ${description.slice(0, 90)}${description.length > 90 ? '…' : ''}`;
}

function cancelManualIndexDraw(removeMode = true) {
  manualIndexState.dragStart = null;
  manualIndexState.drawElement?.remove();
  manualIndexState.drawElement = null;
  if (removeMode) {
    manualIndexState.drawMode = false;
    els.manualIndexDraw?.classList.remove('active');
    if (els.manualIndexDraw) els.manualIndexDraw.textContent = 'Draw link region';
    els.manualIndexOverlay?.classList.remove('draw-mode');
  }
}

function applyManualIndexMappings() {
  const context = getManualIndexContext(manualIndexState.mode);
  const state = getManualIndexModeState(manualIndexState.mode);
  if (!context || !state) return;
  const links = manualIndexState.mappings.filter((mapping) => mapping.type === 'link');
  if (!links.length) return;
  state.manualIndexMappings = {
    sourceKey: context.sourceKey,
    mappings: cloneManualMappings(manualIndexState.mappings),
  };
  const mode = manualIndexState.mode;
  closeManualIndexDialog();
  if (mode === 'build') { renderIndexPanel(); renderBuildTable(); updateBuildStatusSummary(); }
  else if (mode === 'insert') renderInsertWorkspace();
  else if (mode === 'reorder') renderReorderWorkspace();
}

function getManualMappingsForMode(mode) {
  const context = getManualIndexContext(mode);
  const stored = getManualIndexModeState(mode)?.manualIndexMappings;
  if (!context || !stored || stored.sourceKey !== context.sourceKey) return null;
  return stored.mappings || [];
}

function getManualActiveLinks(mode) {
  const mappings = getManualMappingsForMode(mode);
  if (!mappings) return null;
  return mappings
    .filter((mapping) => mapping.type === 'link')
    .map((mapping) => ({ ...mapping, destinationPosition: Number(mapping.destinationPosition) }));
}

function validateManualActiveLinks(mode, finalCount) {
  const links = getManualActiveLinks(mode);
  if (links == null) return null;
  if (!links.length) return { canRefresh: false, reason: 'manual mapping contains no active links.', manual: true };
  const invalid = links.find((link) => !Number.isInteger(link.destinationPosition) || link.destinationPosition < 1 || link.destinationPosition > finalCount || !normalizePdfRect(link.rect));
  if (invalid) return { canRefresh: false, reason: 'one or more manual links have an invalid destination or rectangle.', manual: true };
  return { canRefresh: true, reason: '', manual: true, linkCount: links.length };
}

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
  const source = htmlRows.length ? 'Word/Excel clipboard table' : rows.length ? 'plain text' : '';
  applyBulkPasteAnalysis(rows, source);
}

function applyBulkPasteAnalysis(rows, source) {
  const analysis = validateBulkPasteRows(rows, getSections());
  splitState.bulkPaste = {
    rows: analysis.rows,
    canApply: analysis.canApply,
    source,
    issues: analysis.issues,
  };
  renderBulkPastePreview();
}

async function handleBulkIndexFile(event) {
  const [file] = event.target.files;
  if (!file) return;

  try {
    const lower = file.name.toLowerCase();
    let workbook;
    if (lower.endsWith('.xlsx') || lower.endsWith('.xls')) {
      workbook = XLSX.read(await file.arrayBuffer(), { type: 'array' });
    } else if (lower.endsWith('.csv') || lower.endsWith('.tsv')) {
      const text = await file.text();
      workbook = XLSX.read(text, { type: 'string', FS: lower.endsWith('.tsv') ? '\t' : ',' });
    } else {
      throw new Error('Use an Excel (.xlsx/.xls), CSV, or TSV index file.');
    }

    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    if (!sheet) throw new Error('The index file does not contain a readable worksheet.');
    const cellRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false });
    const rows = parseTabularIndexRows(cellRows);
    if (!rows.length) throw new Error('No usable index rows were found in the first worksheet.');

    els.bulkPasteInput.value = '';
    applyBulkPasteAnalysis(rows, file.name);
  } catch (error) {
    console.error(error);
    splitState.bulkPaste = { rows: [], canApply: false, source: '', issues: [friendlyError(error)] };
    els.bulkPasteSummary.textContent = `Could not import index file: ${friendlyError(error)}`;
    els.bulkPasteSummary.className = 'bulk-paste-summary error';
    els.bulkPastePreview.innerHTML = '<tr><td colspan="5" class="empty-cell">No index entries imported.</td></tr>';
    els.applyBulkPaste.disabled = true;
  } finally {
    event.target.value = '';
  }
}

function parseWordTableHtml(html) {
  try {
    const doc = new DOMParser().parseFromString(html, 'text/html');
    const candidates = [...doc.querySelectorAll('table')].map((table) => {
      const cellRows = [...table.querySelectorAll('tr')].map((tr) => (
        [...tr.children]
          .filter((node) => ['TD', 'TH'].includes(node.tagName))
          .map(extractPastedCellText)
      ));
      return parseTabularIndexRows(cellRows);
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

function parseTabularIndexRows(cellRows) {
  const rows = [];
  let inferredNumber = 1;

  (cellRows || []).forEach((rawCells) => {
    const cells = (rawCells || []).map(normalizePastedCell);
    if (!cells.some(Boolean)) return;
    const parsed = rowFromCells(cells, inferredNumber);
    if (!parsed) return;
    rows.push(parsed);
    if (parsed.inferredNumber) inferredNumber += 1;
    else inferredNumber = Math.max(inferredNumber, parsed.number + 1);
  });

  return rows;
}

function rowFromCells(cells, inferredNumber = null) {
  if (cells.length < 2) return null;

  const normalizedHeaders = cells.map(normalizeHeader);
  const firstHeader = normalizedHeaders[0] || '';
  const obviousHeader = ['no', 'number', 'order', 'position', 'tab', 'description', 'filename', 'file', 'filepath', 'path', 'document', 'documentname'].includes(firstHeader)
    || normalizedHeaders.some((header) => ['description', 'filename', 'documentdate'].includes(header));
  if (obviousHeader && normalizedHeaders.some((header) => ['date', 'documentdate'].includes(header))) return null;

  const numberIndex = cells.findIndex((cell) => parsePastedRowNumber(cell) != null);
  if (numberIndex >= 0) {
    const number = parsePastedRowNumber(cells[numberIndex]);
    const trailing = cells.slice(numberIndex + 1);
    if (!trailing.length) return null;

    const description = normalizePastedCell(trailing[0]);
    const date = normalizePastedCell(trailing.length > 1 ? trailing[trailing.length - 1] : '');
    if (!description || ['description', 'filename', 'document'].includes(normalizeHeader(description))) return null;
    return { number, description, date, inferredNumber: false };
  }

  if (inferredNumber == null) return null;
  if (['no', 'number', 'order', 'position', 'tab', 'description', 'filename', 'file', 'filepath', 'path', 'document', 'documentname'].includes(firstHeader)) return null;

  const description = descriptionFromFilenameOrPath(cells[0]);
  const date = normalizePastedCell(cells.length > 1 ? cells[cells.length - 1] : '');
  if (!description) return null;
  return { number: inferredNumber, description, date, inferredNumber: true };
}

function descriptionFromFilenameOrPath(value) {
  let text = normalizePastedCell(value);
  if (!text) return '';
  text = text.replace(/^['"]+|['"]+$/g, '').trim();
  const pieces = text.split(/[\\/]/);
  let base = pieces[pieces.length - 1] || text;
  base = base.replace(/\.pdf$/i, '');
  base = base
    // Remove a simple ordering prefix such as "01 - " or "01_", but
    // do not mistake a dotted date prefix such as "26.05.15" for ordering.
    .replace(/^\s*\d+\s*(?:[-_^]\s*|\.(?!\d)\s*)/, '')
    // Office filename convention: caret stands in for a colon and underscore
    // stands in for a slash. Restore that punctuation in the Description only;
    // generated filenames are sanitized separately for cross-platform safety.
    .replace(/\^+/g, ':')
    .replace(/_+/g, '/')
    .replace(/\s+/g, ' ')
    .trim();
  return base;
}

function parsePlainIndexText(text) {
  const source = String(text || '').replace(/\u00a0/g, ' ');
  if (!source.trim()) return [];

  const lines = source.split(/\r?\n/).map((line) => line.trim()).filter(Boolean);

  // First choice for clipboard text from Word/Excel tables: tab-separated cells.
  const tabCellRows = lines
    .filter((line) => line.includes('\t'))
    .map((line) => line.split(/\t+/).map(normalizePastedCell));
  const tabRows = parseTabularIndexRows(tabCellRows);
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
      inferredNumber: false,
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
    inferredNumber: Boolean(row.inferredNumber),
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
    const inferred = rows.some((row) => row.inferredNumber);
    const numberingNote = inferred ? ' Positions were inferred from row order.' : ' Explicit numbering matches the current sections.';
    els.bulkPasteSummary.textContent = `${rows.length} entries detected from ${source || 'clipboard text'}.${numberingNote}`;
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
    const source = await PDFDocument.load(splitState.bytes.slice(), { updateMetadata: false });
    const zip = new JSZip();
    const documentsFolder = zip.folder('documents');
    const tabsFolder = els.preserveTabPdfs.checked ? zip.folder('tabs') : null;

    for (let i = 0; i < sections.length; i += 1) {
      const section = sections[i];
      setSplitStatus(`Creating document ${i + 1} of ${sections.length}: ${section.filename}`);

      const output = await createOutputPdfDocument();
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
        const tabOutput = await createOutputPdfDocument();
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
  buildState.matchAssignments = new Map();
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
  buildState.manualIndexMappings = null;
  buildState.matchAssignments = new Map();
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
      && (headerY == null || item.y < headerY - 2));

  // Numeric text can legitimately begin a Description (for example, "504 Plan").
  // A broad left-page cutoff alone can therefore create a fake index row. Infer the
  // actual number-column x track from the densest cluster of numeric candidates and
  // keep only candidates aligned with that track. This tolerates the few-point shift
  // between one- and two-digit right-aligned row numbers while rejecting numbers in
  // the Description column.
  if (numberItems.length >= 2) {
    const clusterTolerance = Math.max(12, pageWidth * 0.025);
    const candidatesByDensity = numberItems.map((candidate) => {
      const cluster = numberItems.filter((other) => Math.abs(other.item.x - candidate.item.x) <= clusterTolerance);
      const xs = cluster.map((entry) => entry.item.x).sort((a, b) => a - b);
      const center = xs[Math.floor(xs.length / 2)];
      return { candidate, cluster, center };
    }).sort((a, b) => b.cluster.length - a.cluster.length || a.center - b.center);

    const numberColumn = candidatesByDensity[0];
    const keepTolerance = Math.max(18, pageWidth * 0.035);
    numberItems = numberItems.filter((candidate) => Math.abs(candidate.item.x - numberColumn.center) <= keepTolerance);
  }

  numberItems.sort((a, b) => b.item.y - a.item.y || a.item.x - b.item.x);

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

    // For an index with no usable existing link annotation, size the clickable
    // rectangle to the actual Description text rather than the full table cell.
    // PDF.js text-item y coordinates are baselines, so add modest ascent/descent
    // allowances plus a small click-friendly padding around the union of all
    // Description fragments in this row.
    const horizontalPadding = 5;
    const verticalPadding = 2.5;
    const textLeft = descriptionItems.length
      ? Math.min(...descriptionItems.map((item) => item.x))
      : candidate.item.x + 24;
    const textRight = descriptionItems.length
      ? Math.max(...descriptionItems.map((item) => item.x + Math.max(item.width, 1)))
      : textLeft + 20;
    const textBottom = descriptionItems.length
      ? Math.min(...descriptionItems.map((item) => item.y - item.height * 0.28))
      : lower + 1;
    const textTop = descriptionItems.length
      ? Math.max(...descriptionItems.map((item) => item.y + item.height * 0.82))
      : upper - 1;

    const descriptionLeft = Math.max(pageX0, textLeft - horizontalPadding);
    const descriptionRight = Math.min(
      pageX1,
      dateLeft - 4,
      Math.max(descriptionLeft + 20, textRight + horizontalPadding),
    );
    const descriptionBottom = Math.max(lower + 0.5, textBottom - verticalPadding);
    const descriptionTop = Math.min(upper - 0.5, textTop + verticalPadding);
    let rect = normalizePdfRect([
      descriptionLeft,
      descriptionBottom,
      descriptionRight,
      descriptionTop,
    ]);
    let rectSource = 'text-bounds';

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


function automaticRowAsActiveLink(row) {
  return { ...row, destinationPosition: row.number, type: 'link' };
}

function getBuildActiveIndexLinks() {
  const manual = getManualActiveLinks('build');
  if (manual != null) return manual;
  return (buildState.indexInfo?.rows || []).map(automaticRowAsActiveLink);
}

function getInsertActiveIndexLinks() {
  const manual = getManualActiveLinks('insert');
  if (manual != null) return manual;
  return getInsertActiveIndexRows().map(automaticRowAsActiveLink);
}

function getReorderActiveIndexLinks() {
  const manual = getManualActiveLinks('reorder');
  if (manual != null) return manual;
  return getReorderActiveIndexRows().map(automaticRowAsActiveLink);
}

function getIndexLinkValidation() {
  if (!buildState.indexFile || !els.createIndexLinks.checked) return { ok: true, message: '' };
  if (buildState.indexLoading) return { ok: false, message: 'Index analysis is still running.' };
  if (buildState.indexInfo?.error) return { ok: false, message: `Could not analyze the index: ${buildState.indexInfo.error}` };

  const rows = buildState.indexInfo?.rows || [];
  const contentCount = buildState.entries.length;
  const manualValidation = validateManualActiveLinks('build', contentCount);
  if (manualValidation) return { ok: manualValidation.canRefresh, message: manualValidation.canRefresh ? '' : `Manual index mapping: ${manualValidation.reason}`, manual: true };
  if (!rows.length) return { ok: false, message: 'No numbered index rows were detected. Use Manual mapping for an atypical index, or turn off index linking.' };
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
  els.manualBuildIndex.disabled = !buildState.indexFile || buildState.indexLoading;
  const buildMatchValidation = getBuildMatchValidation();
  els.matchBuildIndex.disabled = !buildMatchValidation.ok;
  els.matchBuildIndex.title = buildMatchValidation.ok ? 'Propose a content-PDF order from the numbered index.' : buildMatchValidation.message;
  els.createIndexLinks.disabled = !buildState.indexFile || buildState.indexLoading;

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
  const manualLinks = getManualActiveLinks('build');
  const reusedNote = info.existingLinkRectsUsed
    ? ` ${info.existingLinkRectsUsed} existing link rectangle${info.existingLinkRectsUsed === 1 ? '' : 's'} reused for placement.`
    : '';
  const manualNote = manualLinks != null ? ` Manual mapping applied: ${manualLinks.length} active link${manualLinks.length === 1 ? '' : 's'}.` : '';
  els.indexDetectionSummary.textContent = `${buildState.indexFile.name} — ${info.pageCount} page${info.pageCount === 1 ? '' : 's'}, ${info.rows.length} numbered row${info.rows.length === 1 ? '' : 's'} detected.${reusedNote}${manualNote}`;
  els.indexDetectionSummary.className = `index-detection-summary${validation.ok || !els.createIndexLinks.checked ? ' success' : ' error'}`;

  const displayLinks = getBuildActiveIndexLinks();
  if (!displayLinks.length) {
    els.indexLinkBody.innerHTML = '<tr><td colspan="5" class="empty-cell">No automatic rows detected. Use Manual mapping to draw or assign atypical index entries.</td></tr>';
    return;
  }

  els.indexLinkBody.innerHTML = '';
  displayLinks.forEach((row, index) => {
    const position = row.destinationPosition;
    const entry = buildState.entries[position - 1];
    const destination = entry
      ? `Position ${position} — ${entry.description || entry.file.name}`
      : `No content PDF at position ${position}`;
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${row.originalNumber || row.number || `M${index + 1}`}</td>
      <td>${row.pageIndex + 1}</td>
      <td>${escapeHtml(row.description || '')}</td>
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

function getBuildMatchValidation() {
  if (!buildState.indexFile) return { ok: false, message: 'Add an index PDF first.' };
  if (buildState.indexLoading) return { ok: false, message: 'Index analysis is still running.' };
  if (buildState.indexInfo?.error) return { ok: false, message: `Could not analyze the index: ${buildState.indexInfo.error}` };
  if (!buildState.entries.length) return { ok: false, message: 'Add content PDFs first.' };
  const rows = buildState.indexInfo?.rows || [];
  const sequential = validateSequentialIndexRows(rows, buildState.entries.length);
  if (!sequential.ok) {
    return { ok: false, message: `Automatic matching needs one numbered index row per loaded PDF: ${sequential.reason}` };
  }
  return { ok: true, message: '' };
}

function canonicalBuildMatchDate(value) {
  const text = String(value || '').trim();
  let match = text.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{2,4})$/);
  if (match) {
    let year = Number(match[3]);
    if (year < 100) year += year >= 70 ? 1900 : 2000;
    return `${year}-${String(Number(match[1])).padStart(2, '0')}-${String(Number(match[2])).padStart(2, '0')}`;
  }
  match = text.match(/^(\d{4})[.\/-](\d{1,2})[.\/-](\d{1,2})$/);
  if (match) return `${match[1]}-${String(Number(match[2])).padStart(2, '0')}-${String(Number(match[3])).padStart(2, '0')}`;
  match = text.match(/^(\d{2})[.\/-](\d{1,2})[.\/-](\d{1,2})$/);
  if (match) {
    const year = Number(match[1]) + (Number(match[1]) >= 70 ? 1900 : 2000);
    return `${year}-${String(Number(match[2])).padStart(2, '0')}-${String(Number(match[3])).padStart(2, '0')}`;
  }
  return '';
}

function buildMatchFilenameInfo(filename) {
  let text = stripExtension(filename).replace(/^\s*\d+\s*[-_]\s*/, '').trim();
  let date = '';
  const datePrefix = text.match(/^(\d{2,4}[.\/-]\d{1,2}[.\/-]\d{1,2})(?:\s+|\s*[-_]\s*)(.+)$/);
  if (datePrefix) {
    date = canonicalBuildMatchDate(datePrefix[1]);
    text = datePrefix[2].trim();
  }
  text = text.replace(/\^+/g, ':').replace(/_+/g, '/').replace(/\s+/g, ' ').trim();
  return { description: text, date };
}

function buildMatchScore(entry, row) {
  const filenameInfo = buildMatchFilenameInfo(entry.file.name);
  const descriptions = [...new Set([
    String(entry.description || '').trim(),
    humanizeFilename(entry.file.name),
    filenameInfo.description,
  ].filter(Boolean))];
  let score = descriptions.reduce((best, description) => Math.max(best, syncSimilarity(description, row.description)), 0);
  const entryDate = canonicalBuildMatchDate(entry.date) || filenameInfo.date;
  const rowDate = canonicalBuildMatchDate(row.date);
  if (entryDate && rowDate) {
    score += entryDate === rowDate ? 0.12 : -0.04;
  }
  return Math.max(0, Math.min(1, score));
}

// Hungarian algorithm (minimum-cost assignment), used here with cost = 1 - similarity.
// It gives a globally one-to-one proposal instead of letting two similar rows claim the same PDF.
function maximizeBuildMatchAssignment(scoreMatrix) {
  const n = scoreMatrix.length;
  if (!n) return [];
  const m = scoreMatrix[0].length;
  if (n !== m) throw new Error('Content/index counts must match before automatic ordering.');
  const u = Array(n + 1).fill(0);
  const v = Array(m + 1).fill(0);
  const p = Array(m + 1).fill(0);
  const way = Array(m + 1).fill(0);

  for (let i = 1; i <= n; i += 1) {
    p[0] = i;
    let j0 = 0;
    const minv = Array(m + 1).fill(Infinity);
    const used = Array(m + 1).fill(false);
    do {
      used[j0] = true;
      const i0 = p[j0];
      let delta = Infinity;
      let j1 = 0;
      for (let j = 1; j <= m; j += 1) {
        if (used[j]) continue;
        const cur = (1 - scoreMatrix[i0 - 1][j - 1]) - u[i0] - v[j];
        if (cur < minv[j]) { minv[j] = cur; way[j] = j0; }
        if (minv[j] < delta) { delta = minv[j]; j1 = j; }
      }
      for (let j = 0; j <= m; j += 1) {
        if (used[j]) { u[p[j]] += delta; v[j] -= delta; }
        else minv[j] -= delta;
      }
      j0 = j1;
    } while (p[j0] !== 0);

    do {
      const j1 = way[j0];
      p[j0] = p[j1];
      j0 = j1;
    } while (j0 !== 0);
  }

  const assignment = Array(n).fill(-1);
  for (let j = 1; j <= m; j += 1) {
    if (p[j] > 0) assignment[p[j] - 1] = j - 1;
  }
  return assignment;
}

function recomputeBuildMatchProposal() {
  const validation = getBuildMatchValidation();
  buildState.matchAssignments = new Map();
  if (!validation.ok) return validation;
  const rows = [...buildState.indexInfo.rows].sort((a, b) => a.number - b.number);
  const scores = rows.map((row) => buildState.entries.map((entry) => buildMatchScore(entry, row)));
  const assignment = maximizeBuildMatchAssignment(scores);
  rows.forEach((row, rowIndex) => {
    const entryIndex = assignment[rowIndex];
    const entry = buildState.entries[entryIndex];
    buildState.matchAssignments.set(row.number, {
      entryId: entry.id,
      score: scores[rowIndex][entryIndex],
      manual: false,
    });
  });
  return { ok: true, message: '' };
}

function openBuildMatchDialog() {
  const validation = recomputeBuildMatchProposal();
  renderBuildMatchDialog();
  if (!validation.ok) {
    setBuildStatus(validation.message, 'error');
    return;
  }
  if (typeof els.buildMatchDialog.showModal === 'function') els.buildMatchDialog.showModal();
  else els.buildMatchDialog.setAttribute('open', '');
}

function closeBuildMatchDialog() {
  if (els.buildMatchDialog.open && typeof els.buildMatchDialog.close === 'function') els.buildMatchDialog.close();
  else els.buildMatchDialog.removeAttribute('open');
}

function buildMatchAssignmentIsValid() {
  const rows = buildState.indexInfo?.rows || [];
  if (buildState.matchAssignments.size !== rows.length) return false;
  const ids = [...buildState.matchAssignments.values()].map((assignment) => assignment.entryId);
  return ids.length === buildState.entries.length && new Set(ids).size === ids.length && ids.every((id) => buildState.entries.some((entry) => entry.id === id));
}

function renderBuildMatchDialog() {
  const validation = getBuildMatchValidation();
  if (!validation.ok) {
    els.buildMatchSummary.textContent = validation.message;
    els.buildMatchSummary.className = 'status error';
    els.buildMatchBody.innerHTML = `<tr><td colspan="5" class="empty-cell">${escapeHtml(validation.message)}</td></tr>`;
    els.applyBuildMatch.disabled = true;
    return;
  }

  const rows = [...buildState.indexInfo.rows].sort((a, b) => a.number - b.number);
  if (!buildState.matchAssignments.size) recomputeBuildMatchProposal();
  els.buildMatchBody.innerHTML = '';

  rows.forEach((row) => {
    const assignment = buildState.matchAssignments.get(row.number);
    const selectedEntry = buildState.entries.find((entry) => entry.id === assignment?.entryId);
    const confidence = syncConfidence(assignment?.score || 0, Boolean(assignment?.manual));
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${row.number}</td>
      <td>${escapeHtml(row.description || '')}</td>
      <td>${escapeHtml(row.date || '')}</td>
      <td><select class="build-match-source-select" aria-label="Content PDF for index position ${row.number}"></select></td>
      <td><span class="sync-confidence ${confidence.className}">${confidence.label}</span><span class="sync-score">${Math.round((assignment?.score || 0) * 100)}% similarity</span></td>
    `;
    const select = tr.querySelector('select');
    buildState.entries.forEach((entry) => {
      const option = document.createElement('option');
      option.value = entry.id;
      option.textContent = entry.file.name;
      option.selected = entry.id === selectedEntry?.id;
      select.append(option);
    });
    select.addEventListener('change', () => changeBuildMatchAssignment(row.number, select.value));
    els.buildMatchBody.append(tr);
  });

  const assignments = [...buildState.matchAssignments.values()];
  const strongGood = assignments.filter((a) => a.score >= 0.72 && !a.manual).length;
  const review = assignments.filter((a) => a.score < 0.72 && !a.manual).length;
  const manual = assignments.filter((a) => a.manual).length;
  els.buildMatchSummary.textContent = `${rows.length} index rows mapped one-to-one to ${buildState.entries.length} PDFs. ${strongGood} strong/good automatic match${strongGood === 1 ? '' : 'es'}${review ? `; ${review} row${review === 1 ? '' : 's'} should be reviewed` : ''}${manual ? `; ${manual} manual override${manual === 1 ? '' : 's'}` : ''}.`;
  els.buildMatchSummary.className = `status ${review ? '' : 'success'}`.trim();
  els.applyBuildMatch.disabled = !buildMatchAssignmentIsValid();
}

function changeBuildMatchAssignment(rowNumber, newEntryId) {
  const current = buildState.matchAssignments.get(rowNumber);
  if (!current || current.entryId === newEntryId) return;
  const conflictingRow = [...buildState.matchAssignments.entries()].find(([number, assignment]) => number !== rowNumber && assignment.entryId === newEntryId);
  const oldEntryId = current.entryId;
  const rowsByNumber = new Map((buildState.indexInfo?.rows || []).map((row) => [row.number, row]));
  const selectedEntry = buildState.entries.find((entry) => entry.id === newEntryId);
  const currentRow = rowsByNumber.get(rowNumber);
  buildState.matchAssignments.set(rowNumber, {
    entryId: newEntryId,
    score: selectedEntry && currentRow ? buildMatchScore(selectedEntry, currentRow) : 0,
    manual: true,
  });
  if (conflictingRow) {
    const [otherNumber] = conflictingRow;
    const oldEntry = buildState.entries.find((entry) => entry.id === oldEntryId);
    const otherRow = rowsByNumber.get(otherNumber);
    buildState.matchAssignments.set(otherNumber, {
      entryId: oldEntryId,
      score: oldEntry && otherRow ? buildMatchScore(oldEntry, otherRow) : 0,
      manual: true,
    });
  }
  renderBuildMatchDialog();
}

function applyBuildMatchOrder() {
  if (!buildMatchAssignmentIsValid()) return;
  const rows = [...buildState.indexInfo.rows].sort((a, b) => a.number - b.number);
  const byId = new Map(buildState.entries.map((entry) => [entry.id, entry]));
  buildState.entries = rows.map((row) => byId.get(buildState.matchAssignments.get(row.number).entryId)).filter(Boolean);
  const reviewed = [...buildState.matchAssignments.values()].filter((assignment) => assignment.manual || assignment.score < 0.72).length;
  closeBuildMatchDialog();
  renderBuildTable();
  updateBuildStatusSummary();
  setBuildStatus(`Applied index-based content order for ${buildState.entries.length} PDF${buildState.entries.length === 1 ? '' : 's'}${reviewed ? `; ${reviewed} mapping${reviewed === 1 ? '' : 's'} were manual or lower-confidence` : ''}. Review the Final binder order before building.`, reviewed ? '' : 'success');
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
      <td><input class="build-order-input" type="number" min="1" max="${buildState.entries.length}" value="${index + 1}" title="Type a new position and press Enter" aria-label="Content position" /></td>
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

    const orderInput = row.querySelector('.build-order-input');
    orderInput.addEventListener('focus', () => {
      try { orderInput.select(); } catch { /* Number inputs may not expose text selection in every browser. */ }
    });
    orderInput.addEventListener('change', () => moveBuildEntryToPosition(entry.id, orderInput.value));
    orderInput.addEventListener('keydown', (event) => {
      if (event.key === 'Enter') {
        event.preventDefault();
        moveBuildEntryToPosition(entry.id, orderInput.value);
      } else if (event.key === 'Escape') {
        orderInput.value = String(index + 1);
        orderInput.blur();
      }
    });

    row.querySelector('[data-action="description"]').addEventListener('input', (event) => {
      entry.description = event.target.value;
    });
    row.querySelector('[data-action="up"]').addEventListener('click', () => moveBuildEntry(index, -1));
    row.querySelector('[data-action="down"]').addEventListener('click', () => moveBuildEntry(index, 1));
    row.querySelector('[data-action="remove"]').addEventListener('click', () => {
      buildState.entries.splice(index, 1);
      buildState.matchAssignments = new Map();
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

function moveBuildEntryToPosition(entryId, requestedPosition) {
  const from = buildState.entries.findIndex((entry) => entry.id === entryId);
  if (from < 0 || !buildState.entries.length) return;

  const parsed = Number.parseInt(String(requestedPosition), 10);
  const position = Math.min(buildState.entries.length, Math.max(1, Number.isFinite(parsed) ? parsed : from + 1));
  const to = position - 1;
  if (from === to) {
    renderBuildTable();
    return;
  }

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
    setBuildStatus(`${contentCount} content PDF${contentCount === 1 ? '' : 's'} loaded. Positional tab insertion is off.${buildState.indexFile ? ' The index will be the first top-level bookmark named INDEX.' : ''}${indexNote}${duplicateNote}`, duplicateNames.length ? 'error' : '');
    return;
  }

  if (tabCount < contentCount) {
    const missing = contentCount - tabCount;
    setBuildStatus(`${contentCount} content PDF${contentCount === 1 ? '' : 's'} and ${tabCount} tab PDF${tabCount === 1 ? '' : 's'} loaded. Add ${missing} more tab PDF${missing === 1 ? '' : 's'} or turn off tab insertion.${buildState.indexFile ? ' The index will be the first top-level bookmark named INDEX.' : ''}${indexNote}${duplicateNote}`, 'error');
    return;
  }

  const extra = tabCount - contentCount;
  const extraNote = extra ? ` ${extra} extra tab PDF${extra === 1 ? '' : 's'} will not be used.` : '';
  const indexBookmarkNote = buildState.indexFile ? ' The index will be the first top-level bookmark named INDEX.' : '';
  const bookmarkNote = els.createBuildBookmarks.checked ? ` ${contentCount} positional tab bookmark${contentCount === 1 ? '' : 's'} will follow.` : '';
  setBuildStatus(`${contentCount} content PDF${contentCount === 1 ? '' : 's'} will be paired with the first ${contentCount} positional tab PDF${contentCount === 1 ? '' : 's'}.${indexBookmarkNote}${bookmarkNote}${indexNote}${extraNote}${duplicateNote}`, duplicateNames.length ? 'error' : 'success');
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
  const source = await PDFDocument.load(await file.arrayBuffer(), { updateMetadata: false });
  const pages = await output.copyPages(source, source.getPageIndices());
  pages.forEach((page) => output.addPage(page));
  return pages.length;
}

async function appendIndexPdfToOutput(output, file) {
  const source = await PDFDocument.load(await file.arrayBuffer(), { updateMetadata: false });
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
    const output = await createOutputPdfDocument();
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
      bookmarks.push({ title: 'INDEX', pageIndex: indexPageStart });
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
      const rows = getBuildActiveIndexLinks();
      setBuildStatus(`Creating ${rows.length} linked index entr${rows.length === 1 ? 'y' : 'ies'}…`);
      rows.forEach((row) => {
        const targetPageIndex = contentPageTargets[row.destinationPosition - 1];
        const sourcePage = output.getPage(indexPageStart + row.pageIndex);
        if (targetPageIndex == null || !sourcePage) {
          throw new Error(`Could not map index entry to final position ${row.destinationPosition}.`);
        }
        addInternalLinkAnnotation(output, sourcePage, row.rect, targetPageIndex);
        indexLinksCreated += 1;
      });
    }

    if (bookmarks.length) {
      setBuildStatus(`Creating ${bookmarks.length} top-level bookmark${bookmarks.length === 1 ? '' : 's'}…`);
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
    ['INDEX bookmark created', buildState.indexFile ? 'Yes' : 'No'],
    ['Index pages', buildState.indexInfo?.pageCount || ''],
    ['Index Description links created', buildState.indexFile && els.createIndexLinks.checked ? 'Yes' : 'No'],
    ['Detected index rows', buildState.indexInfo?.rows?.length || ''],
  ]);
  configSheet['!cols'] = [{ wch: 22 }, { wch: 70 }];
  XLSX.utils.book_append_sheet(workbook, configSheet, 'Binder Settings');

  const bytes = XLSX.write(workbook, { type: 'array', bookType: 'xlsx' });
  downloadBlob(new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), 'final-binder-order.xlsx');
}



// ---------------------------------------------------------------------------
// Insert Documents workflow (v0.10.1)
// ---------------------------------------------------------------------------

async function loadInsertBinder(file) {
  if (!isPdfFile(file)) {
    setInsertStatus('The existing binder must be a PDF file.', 'error');
    return;
  }

  insertState.file = file;
  insertState.bytes = null;
  insertState.pageCount = 0;
  insertState.frontMatterPageIndexes = [];
  insertState.sourceTabs = [];
  insertState.sourceSections = [];
  insertState.indexRows = [];
  insertState.manualIndexMappings = null;
  insertState.replacementIndexFile = null;
  insertState.replacementIndexInfo = null;
  insertState.replacementIndexLoading = false;
  insertState.operations = [];
  insertState.analyzing = true;
  insertState.error = null;
  renderInsertWorkspace();
  setInsertStatus(`Analyzing ${file.name} for positional TAB pages…`);

  let loadingTask = null;
  try {
    insertState.bytes = new Uint8Array(await file.arrayBuffer());
    loadingTask = pdfjsLib.getDocument({ data: insertState.bytes.slice() });
    const pdf = await loadingTask.promise;
    insertState.pageCount = pdf.numPages;

    const pattern = /\bTAB\s*(?:NO\.?\s*)?[:#\-–—]?\s*([A-Z]|\d{1,3})\b/i;
    const tabs = [];

    for (let pageIndex = 0; pageIndex < pdf.numPages; pageIndex += 1) {
      setInsertStatus(`Analyzing page ${pageIndex + 1} of ${pdf.numPages}…`);
      const page = await pdf.getPage(pageIndex + 1);
      const content = await page.getTextContent();
      const text = content.items.map((item) => item.str).join(' ').replace(/\s+/g, ' ').trim();
      const match = text.slice(0, 1800).match(pattern);
      if (match) {
        const token = String(match[1]).toUpperCase();
        tabs.push({
          pageIndex,
          label: `TAB ${token}`,
          detectedNumber: /^\d+$/.test(token) ? Number(token) : null,
        });
      }
    }

    if (!tabs.length) throw new Error('No TAB divider pages were detected. Use Split PDF / Build Binder for this document, or verify that the tab labels contain extractable text.');

    tabs.sort((a, b) => a.pageIndex - b.pageIndex);
    insertState.sourceTabs = tabs.map((tab, index) => ({ ...tab, position: index + 1 }));
    insertState.frontMatterPageIndexes = Array.from({ length: tabs[0].pageIndex }, (_, index) => index);
    insertState.sourceSections = tabs.map((tab, index) => {
      const nextTab = tabs[index + 1];
      const startPageIndex = tab.pageIndex + 1;
      const endPageIndex = nextTab ? nextTab.pageIndex - 1 : pdf.numPages - 1;
      const pageIndexes = endPageIndex >= startPageIndex
        ? Array.from({ length: endPageIndex - startPageIndex + 1 }, (_, offset) => startPageIndex + offset)
        : [];
      return {
        originalPosition: index + 1,
        tabPageIndex: tab.pageIndex,
        tabLabel: tab.label,
        pageIndexes,
      };
    });

    const rows = [];
    for (const pageIndex of insertState.frontMatterPageIndexes) {
      const page = await pdf.getPage(pageIndex + 1);
      const pageRows = await detectIndexRowsOnPage(page, pageIndex);
      rows.push(...pageRows);
    }
    rows.sort((a, b) => a.number - b.number || a.pageIndex - b.pageIndex || b.rect[3] - a.rect[3]);
    insertState.indexRows = rows;
  } catch (error) {
    console.error(error);
    insertState.error = friendlyError(error);
    setInsertStatus(`Could not analyze the binder: ${friendlyError(error)}`, 'error');
  } finally {
    if (loadingTask) {
      try {
        await loadingTask.destroy();
      } catch (cleanupError) {
        console.warn(`Could not fully release PDF.js resources for ${file.name}.`, cleanupError);
      }
    }
    insertState.analyzing = false;
    renderInsertWorkspace();
  }
}

async function loadInsertReplacementIndex(file) {
  if (!isPdfFile(file)) {
    setInsertStatus('The replacement index must be a PDF file.', 'error');
    return;
  }

  insertState.replacementIndexFile = file;
  insertState.replacementIndexInfo = null;
  insertState.manualIndexMappings = null;
  insertState.replacementIndexLoading = true;
  renderInsertWorkspace();
  setInsertStatus(`Analyzing replacement index ${file.name}…`);

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
    insertState.replacementIndexInfo = {
      pageCount: pdf.numPages,
      rows,
      existingLinkRectsUsed,
    };
  } catch (error) {
    console.error(`Could not analyze replacement index ${file.name}.`, error);
    insertState.replacementIndexInfo = {
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
    insertState.replacementIndexLoading = false;
    renderInsertWorkspace();
  }
}

function getInsertActiveIndexRows() {
  if (insertState.replacementIndexFile) return insertState.replacementIndexInfo?.rows || [];
  return insertState.indexRows || [];
}

function getInsertOldIndexPageIndexes() {
  const pages = (insertState.indexRows || []).map((row) => row.pageIndex);
  const stored = insertState.manualIndexMappings;
  if (stored?.sourceKey?.startsWith('insert-existing:')) pages.push(...stored.mappings.map((mapping) => mapping.pageIndex));
  return [...new Set(pages)].sort((a, b) => a - b);
}

function addInsertOperations(files) {
  const pdfs = files.filter(isPdfFile);
  if (!pdfs.length) return;

  let insertionCount = insertState.operations.filter((operation) => operation.action === 'insert').length;
  pdfs.forEach((file) => {
    insertionCount += 1;
    const baseCount = insertState.sourceSections.length;
    insertState.operations.push({
      id: makeId(),
      file,
      description: humanizeFilename(file.name),
      action: 'insert',
      position: Math.max(1, baseCount + insertionCount),
    });
  });
  renderInsertWorkspace();
}

async function addInsertExtraTabs(files) {
  const pdfs = files.filter(isPdfFile);
  if (!pdfs.length) return;

  const entries = pdfs.map((file) => ({
    id: makeId(),
    file,
    tabNumber: parseTabNumberFromText(file.name),
    bookmarkTitle: fallbackBookmarkTitle(file.name),
    bookmarkSource: 'filename',
    bookmarkStatus: 'reading',
  }));
  insertState.extraTabs.push(...entries);
  sortInsertExtraTabs();
  renderInsertWorkspace();

  await Promise.all(entries.map(readTabBookmarkTitle));
  entries.forEach((entry) => {
    if (entry.tabNumber == null) entry.tabNumber = parseTabNumberFromText(entry.bookmarkTitle);
  });
  sortInsertExtraTabs();
  renderInsertWorkspace();
}

function parseTabNumberFromText(value) {
  const text = stripExtension(String(value || ''));
  const tabMatch = text.match(/\bTAB\s*(?:NO\.?\s*)?[:#\-–—]?\s*(\d{1,3})\b/i);
  if (tabMatch) return Number(tabMatch[1]);
  const leading = text.match(/^\s*(\d{1,3})\b/);
  return leading ? Number(leading[1]) : null;
}

function sortInsertExtraTabs() {
  insertState.extraTabs.sort((a, b) => {
    if (a.tabNumber != null && b.tabNumber != null && a.tabNumber !== b.tabNumber) return a.tabNumber - b.tabNumber;
    if (a.tabNumber != null && b.tabNumber == null) return -1;
    if (a.tabNumber == null && b.tabNumber != null) return 1;
    return naturalCompare(a.file.name, b.file.name);
  });
}

function getInsertIndexDescription(position) {
  const matches = insertState.indexRows.filter((row) => row.number === position);
  return matches.length === 1 ? matches[0].description : '';
}

function getInsertPlan() {
  const baseItems = insertState.sourceSections.map((section) => ({
    kind: 'existing',
    originalPosition: section.originalPosition,
    sourceSection: section,
    description: getInsertIndexDescription(section.originalPosition) || `Existing ${section.tabLabel} content`,
    change: '',
  }));

  const issues = [];
  const operations = insertState.operations.map((operation) => ({
    ...operation,
    position: Number.parseInt(String(operation.position), 10),
  }));
  const inserts = operations.filter((operation) => operation.action === 'insert');
  const replacements = operations.filter((operation) => operation.action === 'replace');
  const finalCount = baseItems.length + inserts.length;

  const targetGroups = new Map();
  operations.forEach((operation) => {
    if (!Number.isInteger(operation.position)) {
      issues.push(`${operation.file.name}: enter a valid final position.`);
      return;
    }
    if (!targetGroups.has(operation.position)) targetGroups.set(operation.position, []);
    targetGroups.get(operation.position).push(operation);
  });
  targetGroups.forEach((group, position) => {
    if (group.length > 1) issues.push(`Final position ${position} is assigned to more than one new PDF.`);
  });

  inserts.forEach((operation) => {
    if (!Number.isInteger(operation.position) || operation.position < 1 || operation.position > finalCount) {
      issues.push(`${operation.file.name}: Insert & shift position must be between 1 and ${Math.max(1, finalCount)}.`);
    }
  });
  replacements.forEach((operation) => {
    if (!Number.isInteger(operation.position) || operation.position < 1 || operation.position > finalCount) {
      issues.push(`${operation.file.name}: Replace position must be between 1 and ${Math.max(1, finalCount)}.`);
    }
  });

  const items = [...baseItems];
  if (!issues.length) {
    [...inserts].sort((a, b) => a.position - b.position).forEach((operation) => {
      items.splice(operation.position - 1, 0, {
        kind: 'new',
        operation,
        description: operation.description || humanizeFilename(operation.file.name),
        change: 'insert',
      });
    });

    [...replacements].sort((a, b) => a.position - b.position).forEach((operation) => {
      const replaced = items[operation.position - 1];
      items[operation.position - 1] = {
        kind: 'new',
        operation,
        description: operation.description || humanizeFilename(operation.file.name),
        change: 'replace',
        replaced,
      };
    });
  }

  return { items, operations, inserts, replacements, finalCount, issues };
}

function getExtraTabForPosition(position) {
  const matches = insertState.extraTabs.filter((tab) => tab.tabNumber === position);
  if (matches.length === 1) return { tab: matches[0], duplicate: false };
  if (matches.length > 1) return { tab: matches[0], duplicate: true };
  return { tab: null, duplicate: false };
}

function getTabAssetForInsertPosition(position) {
  const sourceTab = insertState.sourceTabs[position - 1];
  if (sourceTab) return { kind: 'source', label: sourceTab.label, sourceTab };
  const extra = getExtraTabForPosition(position);
  if (extra.tab) return { kind: 'extra', label: extra.tab.bookmarkTitle || `TAB ${position}`, extraTab: extra.tab, duplicate: extra.duplicate };
  return { kind: 'missing', label: `TAB ${position}` };
}

function validateInsertPlan(plan = getInsertPlan()) {
  const errors = [...plan.issues];
  const warnings = [];

  if (insertState.analyzing) errors.push('Binder analysis is still running.');
  if (insertState.error) errors.push(insertState.error);
  if (!insertState.file) errors.push('Open an existing binder.');
  if (insertState.file && !insertState.sourceTabs.length && !insertState.analyzing) errors.push('No positional TAB pages were detected.');
  if (!insertState.operations.length) errors.push('Add at least one PDF to insert or replace.');
  if (insertState.replacementIndexLoading) errors.push('Replacement index analysis is still running.');
  if (insertState.replacementIndexInfo?.error) errors.push(`Could not analyze the replacement index: ${insertState.replacementIndexInfo.error}`);

  if (!plan.issues.length) {
    for (let position = 1; position <= plan.items.length; position += 1) {
      const asset = getTabAssetForInsertPosition(position);
      if (asset.kind === 'missing') errors.push(`TAB ${position} is required but no source or extra tab PDF is available.`);
      if (asset.duplicate) errors.push(`More than one extra tab PDF maps to TAB ${position}. Keep only one.`);
    }
  }

  const numericMismatches = insertState.sourceTabs.filter((tab) => tab.detectedNumber != null && tab.detectedNumber !== tab.position);
  if (numericMismatches.length) {
    warnings.push(`Detected ${numericMismatches.length} source tab label${numericMismatches.length === 1 ? '' : 's'} whose printed number does not match its positional order.`);
  }

  const indexCheck = getInsertIndexLinkState(plan.items.length);
  if (insertState.replacementIndexFile) {
    if (!insertState.replacementIndexLoading && !insertState.replacementIndexInfo?.error && !indexCheck.canRefresh) {
      errors.push(`Replacement index does not match the finished binder: ${indexCheck.reason}`);
    }
    if (!insertState.indexRows.length) {
      warnings.push('No old index pages were detected in the existing binder. The replacement index will be inserted at the end of the preserved front matter, immediately before TAB 1.');
    }
  } else if (insertState.indexRows.length && !indexCheck.canRefresh) {
    warnings.push(`The existing index has ${insertState.indexRows.length} detected row${insertState.indexRows.length === 1 ? '' : 's'}, but the finished binder has ${plan.items.length} position${plan.items.length === 1 ? '' : 's'}. Its old links will be removed, not guessed.`);
  } else if (indexCheck.canRefresh && plan.replacements.length) {
    warnings.push('Index links can be refreshed, but replacement documents may make the visible index description/date stale. Review the index text before final use.');
  }

  return { ok: errors.length === 0, errors, warnings, indexCheck };
}

function getInsertIndexLinkState(finalCount) {
  const rows = getInsertActiveIndexRows();
  const source = insertState.replacementIndexFile ? 'replacement' : 'existing';
  const manualValidation = validateManualActiveLinks('insert', finalCount);
  if (manualValidation) return { ...manualValidation, source };
  if (!rows.length || !finalCount) return { canRefresh: false, reason: 'No numbered index rows were detected. Use Manual mapping for an atypical index.', source };
  const counts = new Map();
  rows.forEach((row) => counts.set(row.number, (counts.get(row.number) || 0) + 1));
  if ([...counts.values()].some((count) => count !== 1)) return { canRefresh: false, reason: 'duplicate index row numbers were detected.', source };
  if (rows.length !== finalCount) return { canRefresh: false, reason: `${rows.length} numbered row${rows.length === 1 ? '' : 's'} were detected for ${finalCount} final position${finalCount === 1 ? '' : 's'}.`, source };
  for (let number = 1; number <= finalCount; number += 1) {
    if (!counts.has(number)) return { canRefresh: false, reason: `index row ${number} is missing.`, source };
  }
  return { canRefresh: true, reason: '', source };
}

function renderInsertIndexPanel(plan = getInsertPlan(), validation = validateInsertPlan(plan)) {
  els.clearInsertIndex.disabled = !insertState.replacementIndexFile;
  els.manualInsertIndex.disabled = insertState.analyzing || insertState.replacementIndexLoading || !insertState.file || (!(insertState.replacementIndexFile) && !insertState.frontMatterPageIndexes.length);

  if (insertState.replacementIndexLoading) {
    els.insertIndexSummary.textContent = `Analyzing ${insertState.replacementIndexFile?.name || 'replacement index'}…`;
    els.insertIndexSummary.className = 'index-detection-summary';
    els.insertIndexBody.innerHTML = '<tr><td colspan="5" class="empty-cell">Detecting numbered rows and Description link rectangles…</td></tr>';
    return;
  }

  if (insertState.replacementIndexFile) {
    const info = insertState.replacementIndexInfo || { rows: [], pageCount: 0, existingLinkRectsUsed: 0 };
    if (info.error) {
      els.insertIndexSummary.textContent = `Could not analyze ${insertState.replacementIndexFile.name}.`;
      els.insertIndexSummary.className = 'index-detection-summary error';
      els.insertIndexBody.innerHTML = `<tr><td colspan="5" class="empty-cell">${escapeHtml(info.error)}</td></tr>`;
      return;
    }
    const oldPages = getInsertOldIndexPageIndexes();
    const replacementNote = oldPages.length
      ? ` It will replace ${oldPages.length} detected old index page${oldPages.length === 1 ? '' : 's'}.`
      : ' No old index pages were detected; it will be inserted immediately before TAB 1.';
    const reusedNote = info.existingLinkRectsUsed
      ? ` ${info.existingLinkRectsUsed} existing link rectangle${info.existingLinkRectsUsed === 1 ? '' : 's'} will be reused for placement.`
      : '';
    const manualLinks = getManualActiveLinks('insert');
    const manualNote = manualLinks != null ? ` Manual mapping applied: ${manualLinks.length} active link${manualLinks.length === 1 ? '' : 's'}.` : '';
    els.insertIndexSummary.textContent = `${insertState.replacementIndexFile.name} — ${info.pageCount} page${info.pageCount === 1 ? '' : 's'}, ${info.rows.length} numbered row${info.rows.length === 1 ? '' : 's'} detected.${replacementNote}${reusedNote}${manualNote}`;
    els.insertIndexSummary.className = `index-detection-summary${validation.indexCheck.canRefresh ? ' success' : ' error'}`;
  } else if (insertState.indexRows.length) {
    const manualLinks = getManualActiveLinks('insert');
    const manualNote = manualLinks != null ? ` Manual mapping applied: ${manualLinks.length} active link${manualLinks.length === 1 ? '' : 's'}.` : '';
    els.insertIndexSummary.textContent = `Using existing index — ${insertState.indexRows.length} numbered row${insertState.indexRows.length === 1 ? '' : 's'} detected.${manualNote}`;
    els.insertIndexSummary.className = `index-detection-summary${validation.indexCheck.canRefresh ? ' success' : ''}`;
  } else {
    const manualLinks = getManualActiveLinks('insert');
    els.insertIndexSummary.textContent = manualLinks != null ? `Using existing front matter — manual mapping applied with ${manualLinks.length} active link${manualLinks.length === 1 ? '' : 's'}.` : 'Using existing front matter; no numbered index rows were detected. Use Manual mapping for an atypical index.';
    els.insertIndexSummary.className = 'index-detection-summary';
  }

  const rows = getInsertActiveIndexLinks();
  if (!rows.length) {
    els.insertIndexBody.innerHTML = '<tr><td colspan="5" class="empty-cell">No usable automatic rows detected. Use Manual mapping to draw or assign atypical index entries.</td></tr>';
    return;
  }

  els.insertIndexBody.innerHTML = '';
  rows.forEach((row, index) => {
    const position = row.destinationPosition;
    const item = plan.items[position - 1];
    const destination = item
      ? `Position ${position} — ${item.description || (item.kind === 'new' ? item.operation.file.name : item.sourceSection.tabLabel)}`
      : `No final content at position ${position}`;
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${row.originalNumber || row.number || `M${index + 1}`}</td>
      <td>${row.pageIndex + 1}</td>
      <td>${escapeHtml(row.description || '')}</td>
      <td>${escapeHtml(row.date || '')}</td>
      <td class="${item ? 'index-destination' : 'index-destination missing'}">${escapeHtml(destination)}</td>
    `;
    els.insertIndexBody.append(tr);
  });
}

function renderInsertWorkspace() {
  els.clearInsertTabs.disabled = insertState.extraTabs.length === 0;
  els.clearInsertIndex.disabled = !insertState.replacementIndexFile;

  if (!insertState.file) {
    els.insertAnalysisSummary.textContent = 'No binder loaded.';
    els.insertAnalysisSummary.className = 'index-detection-summary';
    els.insertStructureNote.textContent = insertState.extraTabs.length
      ? `${insertState.extraTabs.length} extra tab PDF${insertState.extraTabs.length === 1 ? '' : 's'} loaded; open a binder to use them.`
      : 'Load a binder to detect its tab positions.';
  } else if (insertState.analyzing) {
    els.insertAnalysisSummary.textContent = `Analyzing ${insertState.file.name}…`;
    els.insertAnalysisSummary.className = 'index-detection-summary';
    els.insertStructureNote.textContent = 'Reading page text to locate physical TAB dividers and inspect front-matter index rows.';
  } else if (insertState.error) {
    els.insertAnalysisSummary.textContent = `Could not analyze ${insertState.file.name}.`;
    els.insertAnalysisSummary.className = 'index-detection-summary error';
    els.insertStructureNote.textContent = insertState.error;
  } else {
    const indexNote = insertState.indexRows.length
      ? `${insertState.indexRows.length} numbered index row${insertState.indexRows.length === 1 ? '' : 's'} detected in front matter.`
      : 'No numbered index rows detected in front matter.';
    const extraNote = insertState.extraTabs.length
      ? ` ${insertState.extraTabs.length} extra tab PDF${insertState.extraTabs.length === 1 ? '' : 's'} loaded.`
      : '';
    els.insertAnalysisSummary.textContent = `${insertState.file.name} — ${insertState.pageCount} pages, ${insertState.sourceTabs.length} TAB position${insertState.sourceTabs.length === 1 ? '' : 's'}.`;
    els.insertAnalysisSummary.className = 'index-detection-summary success';
    const oldIndexPages = getInsertOldIndexPageIndexes();
    const frontMatterNote = insertState.replacementIndexFile && oldIndexPages.length
      ? `${Math.max(0, insertState.frontMatterPageIndexes.length - oldIndexPages.length)} non-index front-matter page${Math.max(0, insertState.frontMatterPageIndexes.length - oldIndexPages.length) === 1 ? '' : 's'} will be preserved; ${oldIndexPages.length} detected old index page${oldIndexPages.length === 1 ? '' : 's'} will be replaced.`
      : `${insertState.frontMatterPageIndexes.length} front-matter page${insertState.frontMatterPageIndexes.length === 1 ? '' : 's'} will be preserved.`;
    els.insertStructureNote.textContent = `${frontMatterNote} ${indexNote}${extraNote}`;
  }

  renderInsertOperationsTable();
  const plan = getInsertPlan();
  const validation = validateInsertPlan(plan);
  renderInsertIndexPanel(plan, validation);
  renderInsertPreviewTable(plan);

  els.insertRefreshIndexLinks.disabled = !validation.indexCheck.canRefresh || insertState.replacementIndexLoading;
  const refreshWillRun = validation.indexCheck.canRefresh && els.insertRefreshIndexLinks.checked;
  els.insertDownload.disabled = !validation.ok;

  if (insertState.analyzing) return;
  if (!insertState.file) {
    setInsertStatus('Open an existing tabbed binder to begin.');
    return;
  }
  if (validation.errors.length) {
    setInsertStatus(validation.errors[0], 'error');
    return;
  }

  const insertCount = plan.inserts.length;
  const replaceCount = plan.replacements.length;
  const changes = [];
  if (insertCount) changes.push(`${insertCount} insert-and-shift operation${insertCount === 1 ? '' : 's'}`);
  if (replaceCount) changes.push(`${replaceCount} replacement${replaceCount === 1 ? '' : 's'}`);
  const activeIndexRows = getInsertActiveIndexRows();
  const indexPhrase = insertState.replacementIndexFile
    ? refreshWillRun
      ? ' The replacement index will replace the detected old index pages and its Description links will be created for the final content pages.'
      : ' The replacement index will be used, but Description links are turned off.'
    : activeIndexRows.length
      ? refreshWillRun
        ? ' Existing index links will be regenerated to the final content pages.'
        : ' Existing detected index links will be removed because they cannot be safely mapped.'
      : '';
  const warningPhrase = validation.warnings.length ? ` Warning: ${validation.warnings[0]}` : '';
  setInsertStatus(`${plan.items.length} final position${plan.items.length === 1 ? '' : 's'} from ${changes.join(' and ')}.${indexPhrase}${warningPhrase}`, validation.warnings.length ? '' : 'success');
}

function renderInsertOperationsTable() {
  if (!insertState.operations.length) {
    els.insertOperationsBody.innerHTML = '<tr><td colspan="5" class="empty-cell">Add one or more PDFs to create insertion/replacement operations.</td></tr>';
    return;
  }

  const potentialFinalCount = insertState.sourceSections.length + insertState.operations.filter((operation) => operation.action === 'insert').length;
  els.insertOperationsBody.innerHTML = '';
  insertState.operations.forEach((operation) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><strong>${escapeHtml(operation.file.name)}</strong><span class="tab-bookmark">${formatBytes(operation.file.size)}</span></td>
      <td><input data-action="description" value="${escapeAttr(operation.description)}" /></td>
      <td><input class="insert-position-input" data-action="position" type="number" min="1" max="${Math.max(1, potentialFinalCount)}" value="${escapeAttr(operation.position)}" /></td>
      <td>
        <select data-action="operation">
          <option value="insert" ${operation.action === 'insert' ? 'selected' : ''}>Insert &amp; shift</option>
          <option value="replace" ${operation.action === 'replace' ? 'selected' : ''}>Replace</option>
        </select>
      </td>
      <td><button type="button" class="remove-file" data-action="remove" aria-label="Remove operation">×</button></td>
    `;
    tr.querySelector('[data-action="description"]').addEventListener('change', (event) => {
      operation.description = event.target.value;
      renderInsertWorkspace();
    });
    tr.querySelector('[data-action="position"]').addEventListener('change', (event) => {
      operation.position = Number.parseInt(event.target.value, 10);
      renderInsertWorkspace();
    });
    tr.querySelector('[data-action="operation"]').addEventListener('change', (event) => {
      operation.action = event.target.value;
      const plan = getInsertPlan();
      if (operation.action === 'replace' && operation.position > Math.max(1, plan.finalCount)) operation.position = Math.max(1, plan.finalCount);
      renderInsertWorkspace();
    });
    tr.querySelector('[data-action="remove"]').addEventListener('click', () => {
      insertState.operations = insertState.operations.filter((candidate) => candidate.id !== operation.id);
      renderInsertWorkspace();
    });
    els.insertOperationsBody.append(tr);
  });
}

function renderInsertPreviewTable(plan = getInsertPlan()) {
  if (!insertState.sourceSections.length) {
    els.insertPreviewBody.innerHTML = '<tr><td colspan="4" class="empty-cell">Open a binder to preview its existing positions.</td></tr>';
    return;
  }
  if (plan.issues.length) {
    els.insertPreviewBody.innerHTML = `<tr><td colspan="4" class="empty-cell">${escapeHtml(plan.issues[0])}</td></tr>`;
    return;
  }

  els.insertPreviewBody.innerHTML = '';
  plan.items.forEach((item, index) => {
    const position = index + 1;
    const asset = getTabAssetForInsertPosition(position);
    const tabText = asset.kind === 'source'
      ? `${asset.sourceTab.label} — binder page ${asset.sourceTab.pageIndex + 1}`
      : asset.kind === 'extra'
        ? `${asset.extraTab.file.name} — extra PDF${asset.duplicate ? ' (duplicate mapping)' : ''}`
        : `TAB ${position} — MISSING`;
    const sourceText = item.kind === 'existing'
      ? (() => {
          const pages = item.sourceSection.pageIndexes;
          const range = pages.length ? ` · source pages ${pages[0] + 1}${pages.length > 1 ? `–${pages[pages.length - 1] + 1}` : ''}` : ' · no content pages';
          return item.originalPosition === position
            ? `Existing position ${item.originalPosition}${range}`
            : `Existing position ${item.originalPosition} → shifted${range}`;
        })()
      : item.change === 'replace'
        ? `NEW — replaces final position ${position}`
        : `NEW — inserted at ${position}`;
    const contentText = item.kind === 'existing'
      ? item.description
      : `${item.description} (${item.operation.file.name})`;
    const tr = document.createElement('tr');
    if (item.kind === 'new') tr.classList.add(item.change === 'replace' ? 'insert-row-replace' : 'insert-row-new');
    tr.innerHTML = `
      <td><strong>${position}</strong></td>
      <td class="${asset.kind === 'missing' || asset.duplicate ? 'tab-missing' : ''}">${escapeHtml(tabText)}</td>
      <td>${escapeHtml(contentText)}</td>
      <td>${escapeHtml(sourceText)}</td>
    `;
    els.insertPreviewBody.append(tr);
  });
}

async function appendSourcePageIndexes(output, source, pageIndexes, { removeLinks = false } = {}) {
  if (!pageIndexes.length) return [];
  const pages = await output.copyPages(source, pageIndexes);
  const outputIndexes = [];
  pages.forEach((page) => {
    output.addPage(page);
    outputIndexes.push(output.getPageCount() - 1);
    if (removeLinks) removeLinkAnnotations(output, page);
  });
  return outputIndexes;
}

async function buildInsertedBinder() {
  const plan = getInsertPlan();
  const validation = validateInsertPlan(plan);
  if (!validation.ok) {
    renderInsertWorkspace();
    return;
  }

  els.insertDownload.disabled = true;
  try {
    const output = await createOutputPdfDocument();
    const source = await PDFDocument.load(insertState.bytes.slice(), { updateMetadata: false });
    const bookmarks = [];
    const contentPageTargets = [];
    const frontMatterMap = new Map();
    const oldIndexPageIndexes = getInsertOldIndexPageIndexes();
    const oldIndexPageSet = new Set(oldIndexPageIndexes);
    let replacementIndexStart = null;
    let replacementIndexPageCount = 0;
    let removedReplacementLinks = 0;
    let replacementInserted = false;

    const appendReplacementIndex = async () => {
      if (!insertState.replacementIndexFile || replacementInserted) return;
      replacementIndexStart = output.getPageCount();
      setInsertStatus(`Adding replacement index: ${insertState.replacementIndexFile.name}…`);
      const result = await appendIndexPdfToOutput(output, insertState.replacementIndexFile);
      replacementIndexPageCount = result.pageCount;
      removedReplacementLinks = result.removedLinks;
      replacementInserted = true;
    };

    if (insertState.frontMatterPageIndexes.length) {
      setInsertStatus(`Processing ${insertState.frontMatterPageIndexes.length} front-matter page${insertState.frontMatterPageIndexes.length === 1 ? '' : 's'}…`);
      for (const sourcePageIndex of insertState.frontMatterPageIndexes) {
        if (insertState.replacementIndexFile && oldIndexPageSet.has(sourcePageIndex)) {
          await appendReplacementIndex();
          continue;
        }

        const [copiedPage] = await output.copyPages(source, [sourcePageIndex]);
        output.addPage(copiedPage);
        const outputPageIndex = output.getPageCount() - 1;
        frontMatterMap.set(sourcePageIndex, outputPageIndex);
        if (!insertState.replacementIndexFile && oldIndexPageSet.has(sourcePageIndex)) {
          removeLinkAnnotations(output, copiedPage);
        }
      }
    }

    // If no old index pages were detectable, place the supplied replacement
    // after all preserved front matter and immediately before TAB 1.
    if (insertState.replacementIndexFile && !replacementInserted) {
      await appendReplacementIndex();
    }

    if (els.insertCreateBookmarks.checked) {
      if (insertState.replacementIndexFile && replacementIndexStart != null) {
        bookmarks.push({ title: 'INDEX', pageIndex: replacementIndexStart });
      } else {
        const indexPages = getInsertOldIndexPageIndexes();
        if (indexPages.length) {
          const indexOutputPage = frontMatterMap.get(indexPages[0]);
          if (indexOutputPage != null) bookmarks.push({ title: 'INDEX', pageIndex: indexOutputPage });
        }
      }
    }

    for (let i = 0; i < plan.items.length; i += 1) {
      const position = i + 1;
      const item = plan.items[i];
      const tabAsset = getTabAssetForInsertPosition(position);
      const tabPageIndex = output.getPageCount();

      if (tabAsset.kind === 'source') {
        setInsertStatus(`Adding positional TAB ${position} from the existing binder…`);
        await appendSourcePageIndexes(output, source, [tabAsset.sourceTab.pageIndex]);
      } else if (tabAsset.kind === 'extra') {
        setInsertStatus(`Adding extra positional TAB ${position}: ${tabAsset.extraTab.file.name}`);
        await appendPdfToOutput(output, tabAsset.extraTab.file);
      } else {
        throw new Error(`No tab page is available for final position ${position}.`);
      }

      if (els.insertCreateBookmarks.checked) {
        const title = tabAsset.kind === 'extra'
          ? (tabAsset.extraTab.bookmarkTitle || `TAB ${position}`)
          : (tabAsset.sourceTab.label || `TAB ${position}`);
        bookmarks.push({ title, pageIndex: tabPageIndex });
      }

      const contentPageIndex = output.getPageCount();
      contentPageTargets.push(contentPageIndex);
      if (item.kind === 'existing') {
        setInsertStatus(`Adding existing content for final position ${position}…`);
        await appendSourcePageIndexes(output, source, item.sourceSection.pageIndexes);
      } else {
        setInsertStatus(`Adding new content at final position ${position}: ${item.operation.file.name}`);
        await appendPdfToOutput(output, item.operation.file);
      }
    }

    const indexState = getInsertIndexLinkState(plan.items.length);
    const activeIndexRows = getInsertActiveIndexLinks();
    let refreshedLinks = 0;
    if (indexState.canRefresh && els.insertRefreshIndexLinks.checked) {
      const actionWord = insertState.replacementIndexFile ? 'Creating' : 'Refreshing';
      setInsertStatus(`${actionWord} ${activeIndexRows.length} index link${activeIndexRows.length === 1 ? '' : 's'}…`);
      activeIndexRows.forEach((row) => {
        const sourcePageIndex = insertState.replacementIndexFile
          ? (replacementIndexStart == null ? null : replacementIndexStart + row.pageIndex)
          : frontMatterMap.get(row.pageIndex);
        const targetPageIndex = contentPageTargets[row.destinationPosition - 1];
        if (sourcePageIndex == null || targetPageIndex == null) throw new Error(`Could not map index entry to final position ${row.destinationPosition}.`);
        addInternalLinkAnnotation(output, output.getPage(sourcePageIndex), row.rect, targetPageIndex);
        refreshedLinks += 1;
      });
    }

    if (bookmarks.length) {
      setInsertStatus(`Creating ${bookmarks.length} top-level bookmark${bookmarks.length === 1 ? '' : 's'}…`);
      addTopLevelBookmarks(output, bookmarks);
    }

    setInsertStatus('Saving edited binder…');
    const bytes = await output.save();
    const filename = sanitizePdfFilename(els.insertOutputName.value || 'Edited Binder.pdf');
    downloadBlob(new Blob([bytes], { type: 'application/pdf' }), filename);

    const insertCount = plan.inserts.length;
    const replaceCount = plan.replacements.length;
    const changeParts = [];
    if (insertCount) changeParts.push(`${insertCount} inserted`);
    if (replaceCount) changeParts.push(`${replaceCount} replaced`);

    let indexPhrase = '';
    if (insertState.replacementIndexFile) {
      indexPhrase = ` Replacement index applied (${replacementIndexPageCount} page${replacementIndexPageCount === 1 ? '' : 's'})`;
      if (refreshedLinks) indexPhrase += ` with ${refreshedLinks} linked entr${refreshedLinks === 1 ? 'y' : 'ies'}`;
      indexPhrase += '.';
      if (removedReplacementLinks) indexPhrase += ` ${removedReplacementLinks} stale link annotation${removedReplacementLinks === 1 ? ' was' : 's were'} replaced.`;
    } else if (insertState.indexRows.length) {
      indexPhrase = refreshedLinks
        ? ` ${refreshedLinks} existing index link${refreshedLinks === 1 ? '' : 's'} refreshed.`
        : ' Existing detected index links were removed because they were not safely refreshable.';
    }

    setInsertStatus(`Built ${filename}: ${changeParts.join(', ')}; ${plan.items.length} final position${plan.items.length === 1 ? '' : 's'}.${indexPhrase}`, 'success');
  } catch (error) {
    console.error(error);
    setInsertStatus(`Could not build the edited binder: ${friendlyError(error)}`, 'error');
  } finally {
    const currentValidation = validateInsertPlan(getInsertPlan());
    els.insertDownload.disabled = !currentValidation.ok;
  }
}

// ---------------------------------------------------------------------------
// Sync Binder to New Index workflow (v0.13.0)
// ---------------------------------------------------------------------------

async function loadSyncBinder(file) {
  if (!isPdfFile(file)) {
    setSyncStatus('The existing binder must be a PDF file.', 'error');
    return;
  }

  syncState.file = file;
  syncState.bytes = null;
  syncState.pageCount = 0;
  syncState.frontMatterPageIndexes = [];
  syncState.sourceTabs = [];
  syncState.sourceSections = [];
  syncState.oldIndexRows = [];
  syncState.assignments = new Map();
  syncState.approveOmissions = false;
  els.syncApproveOmissions.checked = false;
  syncState.analyzing = true;
  syncState.error = null;
  renderSyncWorkspace();
  setSyncStatus(`Analyzing ${file.name} for positional TAB pages and the old index…`);

  let loadingTask = null;
  try {
    syncState.bytes = new Uint8Array(await file.arrayBuffer());
    loadingTask = pdfjsLib.getDocument({ data: syncState.bytes.slice() });
    const pdf = await loadingTask.promise;
    syncState.pageCount = pdf.numPages;

    const pattern = /\bTAB\s*(?:NO\.?\s*)?[:#\-–—]?\s*([A-Z]|\d{1,3})\b/i;
    const tabs = [];
    for (let pageIndex = 0; pageIndex < pdf.numPages; pageIndex += 1) {
      setSyncStatus(`Analyzing old binder page ${pageIndex + 1} of ${pdf.numPages}…`);
      const page = await pdf.getPage(pageIndex + 1);
      const content = await page.getTextContent();
      const text = content.items.map((item) => item.str).join(' ').replace(/\s+/g, ' ').trim();
      const match = text.slice(0, 1800).match(pattern);
      if (match) {
        const token = String(match[1]).toUpperCase();
        tabs.push({
          pageIndex,
          label: `TAB ${token}`,
          detectedNumber: /^\d+$/.test(token) ? Number(token) : null,
        });
      }
    }
    if (!tabs.length) throw new Error('No TAB divider pages were detected in the old binder.');

    tabs.sort((a, b) => a.pageIndex - b.pageIndex);
    syncState.sourceTabs = tabs.map((tab, index) => ({ ...tab, position: index + 1 }));
    syncState.frontMatterPageIndexes = Array.from({ length: tabs[0].pageIndex }, (_, index) => index);
    syncState.sourceSections = tabs.map((tab, index) => {
      const nextTab = tabs[index + 1];
      const startPageIndex = tab.pageIndex + 1;
      const endPageIndex = nextTab ? nextTab.pageIndex - 1 : pdf.numPages - 1;
      const pageIndexes = endPageIndex >= startPageIndex
        ? Array.from({ length: endPageIndex - startPageIndex + 1 }, (_, offset) => startPageIndex + offset)
        : [];
      return {
        originalPosition: index + 1,
        tabPageIndex: tab.pageIndex,
        tabLabel: tab.label,
        pageIndexes,
      };
    });

    const rows = [];
    for (const pageIndex of syncState.frontMatterPageIndexes) {
      const page = await pdf.getPage(pageIndex + 1);
      const pageRows = await detectIndexRowsOnPage(page, pageIndex);
      rows.push(...pageRows);
    }
    rows.sort((a, b) => a.number - b.number || a.pageIndex - b.pageIndex || b.rect[3] - a.rect[3]);
    syncState.oldIndexRows = rows;
  } catch (error) {
    console.error(error);
    syncState.error = friendlyError(error);
    setSyncStatus(`Could not analyze the old binder: ${friendlyError(error)}`, 'error');
  } finally {
    if (loadingTask) {
      try { await loadingTask.destroy(); } catch (cleanupError) { console.warn('Could not fully release old-binder PDF.js resources.', cleanupError); }
    }
    syncState.analyzing = false;
    recomputeSyncProposal();
    renderSyncWorkspace();
  }
}

async function loadSyncIndex(file) {
  if (!isPdfFile(file)) {
    setSyncStatus('The new index must be a PDF file.', 'error');
    return;
  }
  syncState.newIndexFile = file;
  syncState.newIndexInfo = null;
  syncState.indexLoading = true;
  syncState.assignments = new Map();
  renderSyncWorkspace();
  setSyncStatus(`Analyzing new index ${file.name}…`);

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
    syncState.newIndexInfo = { pageCount: pdf.numPages, rows, existingLinkRectsUsed };
  } catch (error) {
    console.error(error);
    syncState.newIndexInfo = { pageCount: 0, rows: [], existingLinkRectsUsed: 0, error: friendlyError(error) };
  } finally {
    if (loadingTask) {
      try { await loadingTask.destroy(); } catch (cleanupError) { console.warn('Could not fully release new-index PDF.js resources.', cleanupError); }
    }
    syncState.indexLoading = false;
    recomputeSyncProposal();
    renderSyncWorkspace();
  }
}

function addSyncFiles(files) {
  const pdfs = files.filter(isPdfFile);
  pdfs.forEach((file) => {
    syncState.newFiles.push({
      id: makeId(),
      file,
      description: descriptionFromFilenameOrPath(file.name) || humanizeFilename(file.name),
    });
  });
  recomputeSyncProposal();
  renderSyncWorkspace();
}

async function addSyncTabs(files) {
  const pdfs = files.filter(isPdfFile);
  const entries = pdfs.map((file) => ({
    id: makeId(),
    file,
    tabNumber: parseTabNumberFromText(file.name),
    bookmarkTitle: fallbackBookmarkTitle(file.name),
    bookmarkSource: 'filename',
    bookmarkStatus: 'reading',
  }));
  syncState.extraTabs.push(...entries);
  sortSyncTabs();
  renderSyncWorkspace();
  await Promise.all(entries.map(readTabBookmarkTitle));
  entries.forEach((entry) => {
    if (entry.tabNumber == null) entry.tabNumber = parseTabNumberFromText(entry.bookmarkTitle);
  });
  sortSyncTabs();
  renderSyncWorkspace();
}

function sortSyncTabs() {
  syncState.extraTabs.sort((a, b) => {
    if (a.tabNumber != null && b.tabNumber != null && a.tabNumber !== b.tabNumber) return a.tabNumber - b.tabNumber;
    if (a.tabNumber != null && b.tabNumber == null) return -1;
    if (a.tabNumber == null && b.tabNumber != null) return 1;
    return naturalCompare(a.file.name, b.file.name);
  });
}

function validateSequentialIndexRows(rows, expectedCount = null) {
  if (!rows.length) return { ok: false, reason: 'no numbered index rows were detected.' };
  const counts = new Map();
  rows.forEach((row) => counts.set(row.number, (counts.get(row.number) || 0) + 1));
  if ([...counts.values()].some((count) => count !== 1)) return { ok: false, reason: 'duplicate index row numbers were detected.' };
  const count = expectedCount == null ? rows.length : expectedCount;
  if (rows.length !== count) return { ok: false, reason: `${rows.length} rows were detected for ${count} expected positions.` };
  for (let i = 1; i <= count; i += 1) {
    if (!counts.has(i)) return { ok: false, reason: `index row ${i} is missing.` };
  }
  return { ok: true, reason: '' };
}

function normalizeSyncText(value) {
  return String(value || '')
    .normalize('NFKD')
    .toLowerCase()
    .replace(/&/g, ' and ')
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function syncTokenSet(value) {
  return new Set(normalizeSyncText(value).split(' ').filter((token) => token.length > 1));
}

function syncBigramDice(a, b) {
  const x = normalizeSyncText(a).replace(/\s+/g, ' ');
  const y = normalizeSyncText(b).replace(/\s+/g, ' ');
  if (!x && !y) return 1;
  if (!x || !y) return 0;
  if (x === y) return 1;
  const grams = (text) => {
    const list = [];
    for (let i = 0; i < text.length - 1; i += 1) list.push(text.slice(i, i + 2));
    return list;
  };
  const aGrams = grams(x);
  const bGrams = grams(y);
  const counts = new Map();
  aGrams.forEach((gram) => counts.set(gram, (counts.get(gram) || 0) + 1));
  let overlap = 0;
  bGrams.forEach((gram) => {
    const count = counts.get(gram) || 0;
    if (count > 0) {
      overlap += 1;
      counts.set(gram, count - 1);
    }
  });
  return (2 * overlap) / Math.max(1, aGrams.length + bGrams.length);
}

function syncSimilarity(a, b, dateA = '', dateB = '') {
  const na = normalizeSyncText(a);
  const nb = normalizeSyncText(b);
  if (!na || !nb) return 0;
  if (na === nb) return 1;
  const setA = syncTokenSet(a);
  const setB = syncTokenSet(b);
  const intersection = [...setA].filter((token) => setB.has(token)).length;
  const union = new Set([...setA, ...setB]).size;
  const jaccard = union ? intersection / union : 0;
  const dice = syncBigramDice(a, b);
  let score = Math.max(dice * 0.72 + jaccard * 0.28, jaccard * 0.84 + dice * 0.16);
  const da = normalizeSyncText(dateA);
  const db = normalizeSyncText(dateB);
  if (da && db && da === db) score = Math.min(1, score + 0.12);
  return score;
}

function alignOldAndNewIndexRows(oldRows, newRows) {
  const m = oldRows.length;
  const n = newRows.length;
  const gap = -0.62;
  const dp = Array.from({ length: m + 1 }, () => Array(n + 1).fill(-Infinity));
  const prev = Array.from({ length: m + 1 }, () => Array(n + 1).fill(null));
  dp[0][0] = 0;
  for (let i = 1; i <= m; i += 1) { dp[i][0] = dp[i - 1][0] + gap; prev[i][0] = 'old-gap'; }
  for (let j = 1; j <= n; j += 1) { dp[0][j] = dp[0][j - 1] + gap; prev[0][j] = 'new-gap'; }

  for (let i = 1; i <= m; i += 1) {
    for (let j = 1; j <= n; j += 1) {
      const sim = syncSimilarity(oldRows[i - 1].description, newRows[j - 1].description, oldRows[i - 1].date, newRows[j - 1].date);
      const matchGain = sim >= 0.88 ? 2.8 + sim : sim >= 0.70 ? 1.75 + sim : sim >= 0.52 ? 0.75 + sim : sim >= 0.42 ? 0.05 + sim : -1.1;
      const choices = [
        { value: dp[i - 1][j - 1] + matchGain, action: 'match' },
        { value: dp[i - 1][j] + gap, action: 'old-gap' },
        { value: dp[i][j - 1] + gap, action: 'new-gap' },
      ].sort((a, b) => b.value - a.value);
      dp[i][j] = choices[0].value;
      prev[i][j] = choices[0].action;
    }
  }

  const matches = [];
  let i = m;
  let j = n;
  while (i > 0 || j > 0) {
    const action = prev[i][j];
    if (action === 'match') {
      const oldRow = oldRows[i - 1];
      const newRow = newRows[j - 1];
      const score = syncSimilarity(oldRow.description, newRow.description, oldRow.date, newRow.date);
      if (score >= 0.48) matches.push({ oldRow, newRow, score });
      i -= 1;
      j -= 1;
    } else if (action === 'old-gap') {
      i -= 1;
    } else {
      j -= 1;
    }
  }
  return matches.reverse();
}

function syncConfidence(score, manual = false) {
  if (manual) return { label: 'Manual', className: 'manual' };
  if (score >= 0.90) return { label: 'Strong', className: 'strong' };
  if (score >= 0.72) return { label: 'Good', className: 'good' };
  if (score >= 0.52) return { label: 'Review', className: 'review' };
  return { label: 'Uncertain', className: 'uncertain' };
}

function recomputeSyncProposal({ preserveManual = true } = {}) {
  const oldRows = syncState.oldIndexRows || [];
  const newRows = syncState.newIndexInfo?.rows || [];
  const oldValidation = validateSequentialIndexRows(oldRows, syncState.sourceSections.length);
  const newValidation = validateSequentialIndexRows(newRows);
  if (!oldValidation.ok || !newValidation.ok) {
    if (!preserveManual) syncState.assignments = new Map();
    return;
  }

  const manual = new Map();
  if (preserveManual) {
    syncState.assignments.forEach((value, key) => {
      if (value.manual) manual.set(key, value);
    });
  }

  const assignments = new Map(manual);
  const usedSources = new Set([...manual.values()].map((value) => value.sourceKey));
  const aligned = alignOldAndNewIndexRows(oldRows, newRows);
  aligned.forEach(({ oldRow, newRow, score }) => {
    if (assignments.has(newRow.number)) return;
    const sourceKey = `old:${oldRow.number}`;
    if (usedSources.has(sourceKey)) return;
    assignments.set(newRow.number, { sourceKey, sourceType: 'old', sourceId: oldRow.number, score, manual: false, reason: 'index sequence match' });
    usedSources.add(sourceKey);
  });

  const unmatchedRows = newRows.filter((row) => !assignments.has(row.number));
  const availableFiles = syncState.newFiles.filter((fileEntry) => !usedSources.has(`new:${fileEntry.id}`));
  const pairs = [];
  unmatchedRows.forEach((row) => {
    availableFiles.forEach((fileEntry) => {
      const score = syncSimilarity(fileEntry.description, row.description);
      pairs.push({ row, fileEntry, score });
    });
  });
  pairs.sort((a, b) => b.score - a.score);
  pairs.forEach(({ row, fileEntry, score }) => {
    if (assignments.has(row.number) || usedSources.has(`new:${fileEntry.id}`)) return;
    const forceSingle = unmatchedRows.length === 1 && availableFiles.length === 1;
    if (score < 0.40 && !forceSingle) return;
    const sourceKey = `new:${fileEntry.id}`;
    assignments.set(row.number, { sourceKey, sourceType: 'new', sourceId: fileEntry.id, score, manual: false, reason: 'new filename match' });
    usedSources.add(sourceKey);
  });

  syncState.assignments = assignments;
  syncState.approveOmissions = false;
  els.syncApproveOmissions.checked = false;
}

function syncOldRowForPosition(position) {
  return syncState.oldIndexRows.find((row) => row.number === position) || null;
}

function syncSourceOptions() {
  const old = syncState.sourceSections.map((section) => {
    const row = syncOldRowForPosition(section.originalPosition);
    return {
      key: `old:${section.originalPosition}`,
      label: `Existing ${section.originalPosition} — ${row?.description || section.tabLabel}`,
      type: 'old',
      id: section.originalPosition,
    };
  });
  const fresh = syncState.newFiles.map((entry) => ({
    key: `new:${entry.id}`,
    label: `New PDF — ${entry.file.name}`,
    type: 'new',
    id: entry.id,
  }));
  return { old, fresh };
}

function getSyncUsedSourceKeys() {
  return new Set([...syncState.assignments.values()].map((value) => value.sourceKey).filter(Boolean));
}

function getSyncOmittedOldPositions() {
  const used = getSyncUsedSourceKeys();
  return syncState.sourceSections.map((section) => section.originalPosition).filter((position) => !used.has(`old:${position}`));
}

function getSyncUnusedNewFiles() {
  const used = getSyncUsedSourceKeys();
  return syncState.newFiles.filter((entry) => !used.has(`new:${entry.id}`));
}

function getSyncExtraTabForPosition(position) {
  const matches = syncState.extraTabs.filter((tab) => tab.tabNumber === position);
  if (matches.length === 1) return { tab: matches[0], duplicate: false };
  if (matches.length > 1) return { tab: matches[0], duplicate: true };
  return { tab: null, duplicate: false };
}

function getSyncTabAsset(position) {
  const sourceTab = syncState.sourceTabs[position - 1];
  if (sourceTab) return { kind: 'source', sourceTab };
  const extra = getSyncExtraTabForPosition(position);
  if (extra.tab) return { kind: 'extra', extraTab: extra.tab, duplicate: extra.duplicate };
  return { kind: 'missing' };
}

function validateSyncPlan() {
  const errors = [];
  const warnings = [];
  if (!syncState.file) errors.push('Open the existing binder.');
  if (syncState.analyzing) errors.push('Old binder analysis is still running.');
  if (syncState.error) errors.push(syncState.error);
  if (!syncState.newIndexFile) errors.push('Add the new index PDF.');
  if (syncState.indexLoading) errors.push('New index analysis is still running.');
  if (syncState.newIndexInfo?.error) errors.push(`Could not analyze the new index: ${syncState.newIndexInfo.error}`);

  if (syncState.file && !syncState.analyzing && !syncState.error) {
    const oldCheck = validateSequentialIndexRows(syncState.oldIndexRows, syncState.sourceSections.length);
    if (!oldCheck.ok) errors.push(`The old binder index cannot be aligned automatically: ${oldCheck.reason}`);
  }
  const newRows = syncState.newIndexInfo?.rows || [];
  if (syncState.newIndexFile && !syncState.indexLoading && !syncState.newIndexInfo?.error) {
    const newCheck = validateSequentialIndexRows(newRows);
    if (!newCheck.ok) errors.push(`The new index cannot define final order automatically: ${newCheck.reason}`);
  }

  if (!errors.length) {
    const seen = new Map();
    newRows.forEach((row) => {
      const assignment = syncState.assignments.get(row.number);
      if (!assignment?.sourceKey) {
        errors.push(`Final position ${row.number} (${row.description}) is not assigned to a source document.`);
        return;
      }
      seen.set(assignment.sourceKey, (seen.get(assignment.sourceKey) || 0) + 1);
      const sourceExists = assignment.sourceType === 'old'
        ? syncState.sourceSections.some((section) => section.originalPosition === Number(assignment.sourceId))
        : syncState.newFiles.some((entry) => entry.id === assignment.sourceId);
      if (!sourceExists) errors.push(`Final position ${row.number} points to a source that is no longer loaded.`);
    });
    [...seen.entries()].filter(([, count]) => count > 1).forEach(([sourceKey]) => errors.push(`Source ${sourceKey} is assigned to more than one final position.`));

    for (let position = 1; position <= newRows.length; position += 1) {
      const tab = getSyncTabAsset(position);
      if (tab.kind === 'missing') errors.push(`TAB ${position} is required but no source or extra tab PDF is available.`);
      if (tab.duplicate) errors.push(`More than one extra tab PDF maps to TAB ${position}.`);
    }

    const unusedNew = getSyncUnusedNewFiles();
    if (unusedNew.length) errors.push(`${unusedNew.length} supplied new PDF${unusedNew.length === 1 ? ' is' : 's are'} not assigned to the new index.`);
    const omitted = getSyncOmittedOldPositions();
    if (omitted.length && !syncState.approveOmissions) errors.push(`${omitted.length} existing binder document${omitted.length === 1 ? '' : 's'} would be omitted. Review and approve the omission below.`);
    if (omitted.length) warnings.push(`Existing position${omitted.length === 1 ? '' : 's'} ${omitted.join(', ')} will not appear in the synced binder.`);

    const reviewAssignments = [...syncState.assignments.values()].filter((assignment) => !assignment.manual && assignment.score < 0.72);
    if (reviewAssignments.length) warnings.push(`${reviewAssignments.length} proposed match${reviewAssignments.length === 1 ? '' : 'es'} should be reviewed before building.`);
  }

  return { ok: errors.length === 0, errors, warnings };
}

function renderSyncWorkspace() {
  els.clearSyncFiles.disabled = syncState.newFiles.length === 0;
  els.clearSyncTabs.disabled = syncState.extraTabs.length === 0;
  els.syncRematch.disabled = syncState.analyzing || syncState.indexLoading || !syncState.file || !syncState.newIndexFile;
  const validation = validateSyncPlan();
  els.syncDownload.disabled = !validation.ok;

  if (!syncState.file) {
    els.syncAnalysisSummary.textContent = 'No binder loaded.';
    els.syncAnalysisSummary.className = 'index-detection-summary';
  } else if (syncState.analyzing) {
    els.syncAnalysisSummary.textContent = `Analyzing ${syncState.file.name}…`;
    els.syncAnalysisSummary.className = 'index-detection-summary';
  } else if (syncState.error) {
    els.syncAnalysisSummary.textContent = `Could not analyze ${syncState.file.name}.`;
    els.syncAnalysisSummary.className = 'index-detection-summary error';
  } else {
    els.syncAnalysisSummary.textContent = `${syncState.file.name} — ${syncState.sourceTabs.length} existing position${syncState.sourceTabs.length === 1 ? '' : 's'}, ${syncState.oldIndexRows.length} old index row${syncState.oldIndexRows.length === 1 ? '' : 's'}.`;
    els.syncAnalysisSummary.className = 'index-detection-summary success';
  }

  const newRows = syncState.newIndexInfo?.rows || [];
  const oldNote = syncState.oldIndexRows.length ? `${syncState.oldIndexRows.length} old index entries` : 'no usable old index entries';
  const newNote = syncState.indexLoading ? 'new index analysis running' : syncState.newIndexFile ? `${newRows.length} new index entries` : 'no new index loaded';
  const fileNote = `${syncState.newFiles.length} new PDF${syncState.newFiles.length === 1 ? '' : 's'} supplied`;
  const tabNote = syncState.extraTabs.length ? `; ${syncState.extraTabs.length} extra tab PDF${syncState.extraTabs.length === 1 ? '' : 's'} loaded` : '';
  els.syncStructureNote.textContent = `${oldNote}; ${newNote}; ${fileNote}${tabNote}. The new index row order defines the final binder order.`;

  renderSyncTable();
  renderSyncOmissions();

  if (validation.errors.length) setSyncStatus(validation.errors[0], 'error');
  else if (validation.warnings.length) setSyncStatus(`Ready to build. ${validation.warnings[0]}`);
  else if (syncState.file && syncState.newIndexFile) setSyncStatus(`Ready to build ${newRows.length} final position${newRows.length === 1 ? '' : 's'} from the proposed mapping.`, 'success');
  else setSyncStatus('Open an existing binder and a new index to begin.');
}

function renderSyncTable() {
  const rows = syncState.newIndexInfo?.rows || [];
  if (syncState.indexLoading) {
    els.syncBody.innerHTML = '<tr><td colspan="5" class="empty-cell">Analyzing the new index…</td></tr>';
    return;
  }
  if (!rows.length) {
    els.syncBody.innerHTML = '<tr><td colspan="5" class="empty-cell">Load an existing binder and a new index to generate a proposal.</td></tr>';
    return;
  }
  const options = syncSourceOptions();
  els.syncBody.innerHTML = '';
  rows.forEach((row) => {
    const assignment = syncState.assignments.get(row.number) || null;
    const tr = document.createElement('tr');
    const confidence = assignment ? syncConfidence(assignment.score || 0, assignment.manual) : { label: 'Unassigned', className: 'uncertain' };
    const select = document.createElement('select');
    select.className = 'sync-source-select';
    select.innerHTML = '<option value="">— Choose source —</option>';
    const oldGroup = document.createElement('optgroup');
    oldGroup.label = 'Existing binder';
    options.old.forEach((option) => {
      const el = document.createElement('option');
      el.value = option.key;
      el.textContent = option.label;
      if (assignment?.sourceKey === option.key) el.selected = true;
      oldGroup.append(el);
    });
    select.append(oldGroup);
    if (options.fresh.length) {
      const newGroup = document.createElement('optgroup');
      newGroup.label = 'New PDFs';
      options.fresh.forEach((option) => {
        const el = document.createElement('option');
        el.value = option.key;
        el.textContent = option.label;
        if (assignment?.sourceKey === option.key) el.selected = true;
        newGroup.append(el);
      });
      select.append(newGroup);
    }
    select.addEventListener('change', (event) => {
      const value = event.target.value;
      if (!value) syncState.assignments.delete(row.number);
      else if (value.startsWith('old:')) {
        const id = Number(value.slice(4));
        syncState.assignments.set(row.number, { sourceKey: value, sourceType: 'old', sourceId: id, score: 1, manual: true, reason: 'manual' });
      } else if (value.startsWith('new:')) {
        const id = value.slice(4);
        syncState.assignments.set(row.number, { sourceKey: value, sourceType: 'new', sourceId: id, score: 1, manual: true, reason: 'manual' });
      }
      syncState.approveOmissions = false;
      els.syncApproveOmissions.checked = false;
      renderSyncWorkspace();
    });

    const sourceCell = document.createElement('td');
    sourceCell.append(select);
    tr.innerHTML = `
      <td><strong>${row.number}</strong></td>
      <td><strong>${escapeHtml(row.description)}</strong></td>
      <td>${escapeHtml(row.date || '')}</td>
    `;
    tr.append(sourceCell);
    const confidenceCell = document.createElement('td');
    confidenceCell.innerHTML = `<span class="sync-confidence ${confidence.className}">${escapeHtml(confidence.label)}</span>${assignment && !assignment.manual ? `<span class="sync-score">${Math.round((assignment.score || 0) * 100)}%</span>` : ''}`;
    tr.append(confidenceCell);
    els.syncBody.append(tr);
  });
}

function renderSyncOmissions() {
  const omitted = getSyncOmittedOldPositions();
  els.syncOmissions.classList.toggle('is-hidden', omitted.length === 0);
  if (!omitted.length) {
    els.syncOmissionsList.textContent = '';
    return;
  }
  const labels = omitted.map((position) => {
    const row = syncOldRowForPosition(position);
    return `${position}: ${row?.description || `Existing TAB ${position} content`}`;
  });
  els.syncOmissionsList.textContent = `Would be omitted: ${labels.join(' · ')}`;
}

function getSyncOldIndexPageIndexes() {
  return [...new Set((syncState.oldIndexRows || []).map((row) => row.pageIndex))].sort((a, b) => a - b);
}

async function buildSyncedBinder() {
  const validation = validateSyncPlan();
  if (!validation.ok) {
    renderSyncWorkspace();
    return;
  }

  els.syncDownload.disabled = true;
  try {
    const output = await createOutputPdfDocument();
    const source = await PDFDocument.load(syncState.bytes.slice(), { updateMetadata: false });
    const newRows = syncState.newIndexInfo.rows;
    const oldIndexPages = new Set(getSyncOldIndexPageIndexes());
    let indexStart = null;
    let indexInserted = false;

    const appendNewIndex = async () => {
      if (indexInserted) return;
      indexStart = output.getPageCount();
      setSyncStatus(`Adding new index ${syncState.newIndexFile.name}…`);
      await appendIndexPdfToOutput(output, syncState.newIndexFile);
      indexInserted = true;
    };

    for (const sourcePageIndex of syncState.frontMatterPageIndexes) {
      if (oldIndexPages.has(sourcePageIndex)) {
        await appendNewIndex();
        continue;
      }
      await appendSourcePageIndexes(output, source, [sourcePageIndex]);
    }
    if (!indexInserted) await appendNewIndex();

    const bookmarks = [];
    if (els.syncCreateBookmarks.checked && indexStart != null) bookmarks.push({ title: 'INDEX', pageIndex: indexStart });
    const contentTargets = [];

    for (const row of newRows) {
      const position = row.number;
      const tabAsset = getSyncTabAsset(position);
      const tabPageIndex = output.getPageCount();
      if (tabAsset.kind === 'source') {
        setSyncStatus(`Adding positional TAB ${position} from the existing binder…`);
        await appendSourcePageIndexes(output, source, [tabAsset.sourceTab.pageIndex]);
      } else if (tabAsset.kind === 'extra') {
        setSyncStatus(`Adding extra positional TAB ${position}: ${tabAsset.extraTab.file.name}`);
        await appendPdfToOutput(output, tabAsset.extraTab.file);
      } else {
        throw new Error(`TAB ${position} is missing.`);
      }

      if (els.syncCreateBookmarks.checked) {
        const title = tabAsset.kind === 'extra'
          ? (tabAsset.extraTab.bookmarkTitle || `TAB ${position}`)
          : (tabAsset.sourceTab.label || `TAB ${position}`);
        bookmarks.push({ title, pageIndex: tabPageIndex });
      }

      contentTargets[position - 1] = output.getPageCount();
      const assignment = syncState.assignments.get(position);
      if (assignment.sourceType === 'old') {
        const section = syncState.sourceSections.find((candidate) => candidate.originalPosition === Number(assignment.sourceId));
        if (!section) throw new Error(`Existing source for final position ${position} is missing.`);
        setSyncStatus(`Adding existing content to final position ${position}…`);
        await appendSourcePageIndexes(output, source, section.pageIndexes);
      } else {
        const fileEntry = syncState.newFiles.find((candidate) => candidate.id === assignment.sourceId);
        if (!fileEntry) throw new Error(`New PDF for final position ${position} is missing.`);
        setSyncStatus(`Adding ${fileEntry.file.name} at final position ${position}…`);
        await appendPdfToOutput(output, fileEntry.file);
      }
    }

    setSyncStatus(`Creating ${newRows.length} index link${newRows.length === 1 ? '' : 's'}…`);
    newRows.forEach((row) => {
      const sourcePageIndex = indexStart + row.pageIndex;
      const targetPageIndex = contentTargets[row.number - 1];
      addInternalLinkAnnotation(output, output.getPage(sourcePageIndex), row.rect, targetPageIndex);
    });

    if (bookmarks.length) {
      setSyncStatus(`Creating ${bookmarks.length} top-level bookmark${bookmarks.length === 1 ? '' : 's'}…`);
      addTopLevelBookmarks(output, bookmarks);
    }

    setSyncStatus('Saving synced binder…');
    const bytes = await output.save();
    const filename = sanitizePdfFilename(els.syncOutputName.value || 'Synced Binder.pdf');
    downloadBlob(new Blob([bytes], { type: 'application/pdf' }), filename);
    const existingUsed = [...syncState.assignments.values()].filter((assignment) => assignment.sourceType === 'old').length;
    const newUsed = [...syncState.assignments.values()].filter((assignment) => assignment.sourceType === 'new').length;
    const omitted = getSyncOmittedOldPositions().length;
    setSyncStatus(`Built ${filename}: ${newRows.length} final positions, ${existingUsed} existing documents reused, ${newUsed} new PDFs inserted${omitted ? `, ${omitted} old document${omitted === 1 ? '' : 's'} omitted` : ''}.`, 'success');
  } catch (error) {
    console.error(error);
    setSyncStatus(`Could not build the synced binder: ${friendlyError(error)}`, 'error');
  } finally {
    els.syncDownload.disabled = !validateSyncPlan().ok;
  }
}


// ---------------------------------------------------------------------------
// Reorder Binder workflow (v0.11.0)
// ---------------------------------------------------------------------------

async function loadReorderBinder(file) {
  if (!isPdfFile(file)) {
    setReorderStatus('The existing binder must be a PDF file.', 'error');
    return;
  }

  reorderState.file = file;
  reorderState.bytes = null;
  reorderState.pageCount = 0;
  reorderState.frontMatterPageIndexes = [];
  reorderState.sourceTabs = [];
  reorderState.sourceSections = [];
  reorderState.indexRows = [];
  reorderState.manualIndexMappings = null;
  reorderState.items = [];
  reorderState.replacementIndexFile = null;
  reorderState.replacementIndexInfo = null;
  reorderState.replacementIndexLoading = false;
  reorderState.analyzing = true;
  reorderState.error = null;
  reorderState.draggedItemId = null;
  renderReorderWorkspace();
  setReorderStatus(`Analyzing ${file.name} for positional TAB pages…`);

  let loadingTask = null;
  try {
    reorderState.bytes = new Uint8Array(await file.arrayBuffer());
    loadingTask = pdfjsLib.getDocument({ data: reorderState.bytes.slice() });
    const pdf = await loadingTask.promise;
    reorderState.pageCount = pdf.numPages;

    const pattern = /\bTAB\s*(?:NO\.?\s*)?[:#\-–—]?\s*([A-Z]|\d{1,3})\b/i;
    const tabs = [];

    for (let pageIndex = 0; pageIndex < pdf.numPages; pageIndex += 1) {
      setReorderStatus(`Analyzing page ${pageIndex + 1} of ${pdf.numPages}…`);
      const page = await pdf.getPage(pageIndex + 1);
      const content = await page.getTextContent();
      const text = content.items.map((item) => item.str).join(' ').replace(/\s+/g, ' ').trim();
      const match = text.slice(0, 1800).match(pattern);
      if (match) {
        const token = String(match[1]).toUpperCase();
        tabs.push({
          pageIndex,
          label: `TAB ${token}`,
          detectedNumber: /^\d+$/.test(token) ? Number(token) : null,
        });
      }
    }

    if (!tabs.length) throw new Error('No TAB divider pages were detected. Use Split PDF / Build Binder for this document, or verify that the tab labels contain extractable text.');

    tabs.sort((a, b) => a.pageIndex - b.pageIndex);
    reorderState.sourceTabs = tabs.map((tab, index) => ({ ...tab, position: index + 1 }));
    reorderState.frontMatterPageIndexes = Array.from({ length: tabs[0].pageIndex }, (_, index) => index);
    reorderState.sourceSections = tabs.map((tab, index) => {
      const nextTab = tabs[index + 1];
      const startPageIndex = tab.pageIndex + 1;
      const endPageIndex = nextTab ? nextTab.pageIndex - 1 : pdf.numPages - 1;
      const pageIndexes = endPageIndex >= startPageIndex
        ? Array.from({ length: endPageIndex - startPageIndex + 1 }, (_, offset) => startPageIndex + offset)
        : [];
      return {
        originalPosition: index + 1,
        tabPageIndex: tab.pageIndex,
        tabLabel: tab.label,
        pageIndexes,
      };
    });

    const rows = [];
    for (const pageIndex of reorderState.frontMatterPageIndexes) {
      const page = await pdf.getPage(pageIndex + 1);
      const pageRows = await detectIndexRowsOnPage(page, pageIndex);
      rows.push(...pageRows);
    }
    rows.sort((a, b) => a.number - b.number || a.pageIndex - b.pageIndex || b.rect[3] - a.rect[3]);
    reorderState.indexRows = rows;

    reorderState.items = reorderState.sourceSections.map((section) => {
      const matches = rows.filter((row) => row.number === section.originalPosition);
      return {
        id: makeId(),
        originalPosition: section.originalPosition,
        sourceSection: section,
        description: matches.length === 1 && matches[0].description
          ? matches[0].description
          : `Existing ${section.tabLabel} content`,
      };
    });
  } catch (error) {
    console.error(error);
    reorderState.error = friendlyError(error);
    setReorderStatus(`Could not analyze the binder: ${friendlyError(error)}`, 'error');
  } finally {
    if (loadingTask) {
      try {
        await loadingTask.destroy();
      } catch (cleanupError) {
        console.warn(`Could not fully release PDF.js resources for ${file.name}.`, cleanupError);
      }
    }
    reorderState.analyzing = false;
    renderReorderWorkspace();
  }
}

async function loadReorderReplacementIndex(file) {
  if (!isPdfFile(file)) {
    setReorderStatus('The replacement index must be a PDF file.', 'error');
    return;
  }

  reorderState.replacementIndexFile = file;
  reorderState.replacementIndexInfo = null;
  reorderState.manualIndexMappings = null;
  reorderState.replacementIndexLoading = true;
  renderReorderWorkspace();
  setReorderStatus(`Analyzing replacement index ${file.name}…`);

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
    reorderState.replacementIndexInfo = {
      pageCount: pdf.numPages,
      rows,
      existingLinkRectsUsed,
    };
  } catch (error) {
    console.error(`Could not analyze replacement index ${file.name}.`, error);
    reorderState.replacementIndexInfo = {
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
    reorderState.replacementIndexLoading = false;
    renderReorderWorkspace();
  }
}

function getReorderActiveIndexRows() {
  if (reorderState.replacementIndexFile) return reorderState.replacementIndexInfo?.rows || [];
  return reorderState.indexRows || [];
}

function getReorderOldIndexPageIndexes() {
  const pages = (reorderState.indexRows || []).map((row) => row.pageIndex);
  const stored = reorderState.manualIndexMappings;
  if (stored?.sourceKey?.startsWith('reorder-existing:')) pages.push(...stored.mappings.map((mapping) => mapping.pageIndex));
  return [...new Set(pages)].sort((a, b) => a - b);
}

function getReorderIndexLinkState() {
  const rows = getReorderActiveIndexRows();
  const finalCount = reorderState.items.length;
  const source = reorderState.replacementIndexFile ? 'replacement' : 'existing';
  const manualValidation = validateManualActiveLinks('reorder', finalCount);
  if (manualValidation) return { ...manualValidation, source };
  if (!rows.length || !finalCount) return { canRefresh: false, reason: 'No numbered index rows were detected. Use Manual mapping for an atypical index.', source };
  const counts = new Map();
  rows.forEach((row) => counts.set(row.number, (counts.get(row.number) || 0) + 1));
  if ([...counts.values()].some((count) => count !== 1)) return { canRefresh: false, reason: 'duplicate index row numbers were detected.', source };
  if (rows.length !== finalCount) return { canRefresh: false, reason: `${rows.length} numbered row${rows.length === 1 ? '' : 's'} were detected for ${finalCount} final position${finalCount === 1 ? '' : 's'}.`, source };
  for (let number = 1; number <= finalCount; number += 1) {
    if (!counts.has(number)) return { canRefresh: false, reason: `index row ${number} is missing.`, source };
  }
  return { canRefresh: true, reason: '', source };
}

function isReorderChanged() {
  return reorderState.items.some((item, index) => item.originalPosition !== index + 1);
}

function validateReorderPlan() {
  const errors = [];
  const warnings = [];

  if (reorderState.analyzing) errors.push('Binder analysis is still running.');
  if (reorderState.error) errors.push(reorderState.error);
  if (!reorderState.file) errors.push('Open an existing binder.');
  if (reorderState.file && !reorderState.sourceTabs.length && !reorderState.analyzing) errors.push('No positional TAB pages were detected.');
  if (reorderState.replacementIndexLoading) errors.push('Replacement index analysis is still running.');
  if (reorderState.replacementIndexInfo?.error) errors.push(`Could not analyze the replacement index: ${reorderState.replacementIndexInfo.error}`);

  const numericMismatches = reorderState.sourceTabs.filter((tab) => tab.detectedNumber != null && tab.detectedNumber !== tab.position);
  if (numericMismatches.length) {
    warnings.push(`Detected ${numericMismatches.length} source tab label${numericMismatches.length === 1 ? '' : 's'} whose printed number does not match its positional order.`);
  }

  const indexCheck = getReorderIndexLinkState();
  if (reorderState.replacementIndexFile) {
    if (!reorderState.replacementIndexLoading && !reorderState.replacementIndexInfo?.error && !indexCheck.canRefresh) {
      errors.push(`Replacement index does not match the binder: ${indexCheck.reason}`);
    }
    if (!reorderState.indexRows.length) {
      warnings.push('No old index pages were detected in the existing binder. The replacement index will be inserted at the end of the preserved front matter, immediately before TAB 1.');
    }
  } else if (reorderState.indexRows.length && !indexCheck.canRefresh) {
    warnings.push(`The existing index cannot be safely relinked: ${indexCheck.reason}`);
  } else if (isReorderChanged() && reorderState.indexRows.length) {
    warnings.push('The existing index text remains in its original order. Its links can be refreshed, but the visible descriptions/dates will be stale unless you supply an updated replacement index.');
  }

  return { ok: errors.length === 0, errors, warnings, indexCheck };
}

function renderReorderWorkspace() {
  els.clearReorderIndex.disabled = !reorderState.replacementIndexFile;
  els.manualReorderIndex.disabled = reorderState.analyzing || reorderState.replacementIndexLoading || !reorderState.file || (!(reorderState.replacementIndexFile) && !reorderState.frontMatterPageIndexes.length);
  els.resetReorderOrder.disabled = !reorderState.items.length || !isReorderChanged();

  if (!reorderState.file) {
    els.reorderAnalysisSummary.textContent = 'No binder loaded.';
    els.reorderAnalysisSummary.className = 'index-detection-summary';
    els.reorderStructureNote.textContent = 'Load a binder to detect its tab positions.';
  } else if (reorderState.analyzing) {
    els.reorderAnalysisSummary.textContent = `Analyzing ${reorderState.file.name}…`;
    els.reorderAnalysisSummary.className = 'index-detection-summary';
    els.reorderStructureNote.textContent = 'Reading page text to locate physical TAB dividers and inspect front-matter index rows.';
  } else if (reorderState.error) {
    els.reorderAnalysisSummary.textContent = `Could not analyze ${reorderState.file.name}.`;
    els.reorderAnalysisSummary.className = 'index-detection-summary error';
    els.reorderStructureNote.textContent = reorderState.error;
  } else {
    const indexNote = reorderState.indexRows.length
      ? `${reorderState.indexRows.length} numbered index row${reorderState.indexRows.length === 1 ? '' : 's'} detected in front matter.`
      : 'No numbered index rows detected in front matter.';
    els.reorderAnalysisSummary.textContent = `${reorderState.file.name} — ${reorderState.pageCount} pages, ${reorderState.sourceTabs.length} TAB position${reorderState.sourceTabs.length === 1 ? '' : 's'}.`;
    els.reorderAnalysisSummary.className = 'index-detection-summary success';
    const oldIndexPages = getReorderOldIndexPageIndexes();
    const frontMatterNote = reorderState.replacementIndexFile && oldIndexPages.length
      ? `${Math.max(0, reorderState.frontMatterPageIndexes.length - oldIndexPages.length)} non-index front-matter page${Math.max(0, reorderState.frontMatterPageIndexes.length - oldIndexPages.length) === 1 ? '' : 's'} will be preserved; ${oldIndexPages.length} detected old index page${oldIndexPages.length === 1 ? '' : 's'} will be replaced.`
      : `${reorderState.frontMatterPageIndexes.length} front-matter page${reorderState.frontMatterPageIndexes.length === 1 ? '' : 's'} will be preserved.`;
    els.reorderStructureNote.textContent = `${frontMatterNote} ${indexNote}`;
  }

  const validation = validateReorderPlan();
  renderReorderIndexPanel(validation);
  renderReorderTable();

  els.reorderRefreshIndexLinks.disabled = !validation.indexCheck.canRefresh || reorderState.replacementIndexLoading;
  els.reorderDownload.disabled = !validation.ok;

  if (reorderState.analyzing) return;
  if (!reorderState.file) {
    setReorderStatus('Open an existing tabbed binder to begin.');
    return;
  }
  if (validation.errors.length) {
    setReorderStatus(validation.errors[0], 'error');
    return;
  }

  const changed = isReorderChanged();
  const refreshWillRun = validation.indexCheck.canRefresh && els.reorderRefreshIndexLinks.checked;
  const indexPhrase = reorderState.replacementIndexFile
    ? refreshWillRun
      ? ' The replacement index will replace the detected old index pages and its Description links will be created for the reordered content.'
      : ' The replacement index will be used, but Description links are turned off.'
    : reorderState.indexRows.length
      ? refreshWillRun
        ? ' Existing index links will be regenerated to the reordered content pages.'
        : ' Existing detected index links will be removed.'
      : '';
  const warningPhrase = validation.warnings.length ? ` Warning: ${validation.warnings[0]}` : '';
  setReorderStatus(`${reorderState.items.length} position${reorderState.items.length === 1 ? '' : 's'} loaded; order is ${changed ? 'modified' : 'unchanged'}.${indexPhrase}${warningPhrase}`, validation.warnings.length ? '' : 'success');
}

function renderReorderIndexPanel(validation = validateReorderPlan()) {
  if (reorderState.replacementIndexLoading) {
    els.reorderIndexSummary.textContent = `Analyzing ${reorderState.replacementIndexFile?.name || 'replacement index'}…`;
    els.reorderIndexSummary.className = 'index-detection-summary';
    els.reorderIndexBody.innerHTML = '<tr><td colspan="5" class="empty-cell">Detecting numbered rows and Description link rectangles…</td></tr>';
    return;
  }

  if (reorderState.replacementIndexFile) {
    const info = reorderState.replacementIndexInfo || { rows: [], pageCount: 0, existingLinkRectsUsed: 0 };
    if (info.error) {
      els.reorderIndexSummary.textContent = `Could not analyze ${reorderState.replacementIndexFile.name}.`;
      els.reorderIndexSummary.className = 'index-detection-summary error';
      els.reorderIndexBody.innerHTML = `<tr><td colspan="5" class="empty-cell">${escapeHtml(info.error)}</td></tr>`;
      return;
    }
    const oldPages = getReorderOldIndexPageIndexes();
    const replacementNote = oldPages.length
      ? ` It will replace ${oldPages.length} detected old index page${oldPages.length === 1 ? '' : 's'}.`
      : ' No old index pages were detected; it will be inserted immediately before TAB 1.';
    const reusedNote = info.existingLinkRectsUsed
      ? ` ${info.existingLinkRectsUsed} existing link rectangle${info.existingLinkRectsUsed === 1 ? '' : 's'} will be reused for placement.`
      : '';
    const manualLinks = getManualActiveLinks('reorder');
    const manualNote = manualLinks != null ? ` Manual mapping applied: ${manualLinks.length} active link${manualLinks.length === 1 ? '' : 's'}.` : '';
    els.reorderIndexSummary.textContent = `${reorderState.replacementIndexFile.name} — ${info.pageCount} page${info.pageCount === 1 ? '' : 's'}, ${info.rows.length} numbered row${info.rows.length === 1 ? '' : 's'} detected.${replacementNote}${reusedNote}${manualNote}`;
    els.reorderIndexSummary.className = `index-detection-summary${validation.indexCheck.canRefresh ? ' success' : ' error'}`;
  } else if (reorderState.indexRows.length) {
    const manualLinks = getManualActiveLinks('reorder');
    const manualNote = manualLinks != null ? ` Manual mapping applied: ${manualLinks.length} active link${manualLinks.length === 1 ? '' : 's'}.` : '';
    els.reorderIndexSummary.textContent = `Using existing index — ${reorderState.indexRows.length} numbered row${reorderState.indexRows.length === 1 ? '' : 's'} detected.${manualNote}`;
    els.reorderIndexSummary.className = `index-detection-summary${validation.indexCheck.canRefresh ? ' success' : ''}`;
  } else {
    const manualLinks = getManualActiveLinks('reorder');
    els.reorderIndexSummary.textContent = manualLinks != null ? `Using existing front matter — manual mapping applied with ${manualLinks.length} active link${manualLinks.length === 1 ? '' : 's'}.` : 'Using existing front matter; no numbered index rows were detected. Use Manual mapping for an atypical index.';
    els.reorderIndexSummary.className = 'index-detection-summary';
  }

  const rows = getReorderActiveIndexLinks();
  if (!rows.length) {
    els.reorderIndexBody.innerHTML = '<tr><td colspan="5" class="empty-cell">No usable automatic rows detected. Use Manual mapping to draw or assign atypical index entries.</td></tr>';
    return;
  }

  els.reorderIndexBody.innerHTML = '';
  rows.forEach((row, index) => {
    const position = row.destinationPosition;
    const item = reorderState.items[position - 1];
    const destination = item
      ? `Position ${position} — ${item.description}`
      : `No final content at position ${position}`;
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${row.originalNumber || row.number || `M${index + 1}`}</td>
      <td>${row.pageIndex + 1}</td>
      <td>${escapeHtml(row.description || '')}</td>
      <td>${escapeHtml(row.date || '')}</td>
      <td class="${item ? 'index-destination' : 'index-destination missing'}">${escapeHtml(destination)}</td>
    `;
    els.reorderIndexBody.append(tr);
  });
}

function renderReorderTable() {
  if (!reorderState.items.length) {
    els.reorderBody.innerHTML = '<tr><td colspan="6" class="empty-cell">Open a binder to load its existing tabbed content.</td></tr>';
    return;
  }

  els.reorderBody.innerHTML = '';
  reorderState.items.forEach((item, index) => {
    const position = index + 1;
    const sourceTab = reorderState.sourceTabs[index];
    const pages = item.sourceSection.pageIndexes;
    const pageRange = pages.length
      ? `source pages ${pages[0] + 1}${pages.length > 1 ? `–${pages[pages.length - 1] + 1}` : ''}`
      : 'no content pages';
    const moved = item.originalPosition !== position;
    const row = document.createElement('tr');
    row.dataset.itemId = item.id;
    if (moved) row.classList.add('reorder-moved');
    row.innerHTML = `
      <td><span class="drag-handle" draggable="true" title="Drag to reorder" aria-label="Drag to reorder">☰</span></td>
      <td><input class="reorder-position-input" type="number" min="1" max="${reorderState.items.length}" value="${position}" title="Type a new final position and press Enter" aria-label="Final position" /></td>
      <td><strong>${escapeHtml(sourceTab?.label || `TAB ${position}`)}</strong><span class="tab-bookmark">binder page ${(sourceTab?.pageIndex ?? 0) + 1}</span></td>
      <td>${escapeHtml(item.description)}</td>
      <td>Original ${item.originalPosition}<span class="tab-bookmark">${escapeHtml(pageRange)}</span></td>
      <td>
        <div class="order-buttons">
          <button type="button" data-action="up" ${index === 0 ? 'disabled' : ''} aria-label="Move content up">↑</button>
          <button type="button" data-action="down" ${index === reorderState.items.length - 1 ? 'disabled' : ''} aria-label="Move content down">↓</button>
        </div>
      </td>
    `;

    const positionInput = row.querySelector('.reorder-position-input');
    positionInput.addEventListener('focus', () => {
      try { positionInput.select(); } catch { /* Number inputs may not expose selection in every browser. */ }
    });
    positionInput.addEventListener('change', () => moveReorderItemToPosition(item.id, positionInput.value));
    positionInput.addEventListener('keydown', (event) => {
      if (event.key === 'Enter') {
        event.preventDefault();
        moveReorderItemToPosition(item.id, positionInput.value);
      } else if (event.key === 'Escape') {
        positionInput.value = String(position);
        positionInput.blur();
      }
    });
    row.querySelector('[data-action="up"]').addEventListener('click', () => moveReorderItem(index, -1));
    row.querySelector('[data-action="down"]').addEventListener('click', () => moveReorderItem(index, 1));
    attachReorderDragEvents(row, item.id);
    els.reorderBody.append(row);
  });
}

function attachReorderDragEvents(row, itemId) {
  const handle = row.querySelector('.drag-handle');
  handle.addEventListener('dragstart', (event) => {
    reorderState.draggedItemId = itemId;
    row.classList.add('dragging');
    event.dataTransfer.effectAllowed = 'move';
    event.dataTransfer.setData('text/plain', itemId);
  });
  handle.addEventListener('dragend', () => {
    reorderState.draggedItemId = null;
    [...els.reorderBody.querySelectorAll('tr')].forEach((item) => item.classList.remove('dragging', 'drag-over'));
  });

  row.addEventListener('dragover', (event) => {
    if (!reorderState.draggedItemId || reorderState.draggedItemId === itemId) return;
    event.preventDefault();
    event.dataTransfer.dropEffect = 'move';
    row.classList.add('drag-over');
  });
  row.addEventListener('dragleave', () => row.classList.remove('drag-over'));
  row.addEventListener('drop', (event) => {
    event.preventDefault();
    row.classList.remove('drag-over');
    const draggedId = reorderState.draggedItemId || event.dataTransfer.getData('text/plain');
    reorderReorderItem(draggedId, itemId);
  });
}

function reorderReorderItem(draggedId, targetId) {
  const from = reorderState.items.findIndex((item) => item.id === draggedId);
  const to = reorderState.items.findIndex((item) => item.id === targetId);
  if (from < 0 || to < 0 || from === to) return;
  const [item] = reorderState.items.splice(from, 1);
  reorderState.items.splice(to, 0, item);
  renderReorderWorkspace();
}

function moveReorderItemToPosition(itemId, requestedPosition) {
  const from = reorderState.items.findIndex((item) => item.id === itemId);
  if (from < 0 || !reorderState.items.length) return;
  const parsed = Number.parseInt(String(requestedPosition), 10);
  const position = Math.min(reorderState.items.length, Math.max(1, Number.isFinite(parsed) ? parsed : from + 1));
  const to = position - 1;
  if (from === to) {
    renderReorderWorkspace();
    return;
  }
  const [item] = reorderState.items.splice(from, 1);
  reorderState.items.splice(to, 0, item);
  renderReorderWorkspace();
}

function moveReorderItem(index, delta) {
  const target = index + delta;
  if (target < 0 || target >= reorderState.items.length) return;
  [reorderState.items[index], reorderState.items[target]] = [reorderState.items[target], reorderState.items[index]];
  renderReorderWorkspace();
}

function resetReorderOrder() {
  reorderState.items.sort((a, b) => a.originalPosition - b.originalPosition);
  renderReorderWorkspace();
}

async function buildReorderedBinder() {
  const validation = validateReorderPlan();
  if (!validation.ok) {
    renderReorderWorkspace();
    return;
  }

  els.reorderDownload.disabled = true;
  try {
    const output = await createOutputPdfDocument();
    const source = await PDFDocument.load(reorderState.bytes.slice(), { updateMetadata: false });
    const bookmarks = [];
    const contentPageTargets = [];
    const frontMatterMap = new Map();
    const oldIndexPageIndexes = getReorderOldIndexPageIndexes();
    const oldIndexPageSet = new Set(oldIndexPageIndexes);
    let replacementIndexStart = null;
    let replacementIndexPageCount = 0;
    let replacementInserted = false;

    const appendReplacementIndex = async () => {
      if (!reorderState.replacementIndexFile || replacementInserted) return;
      replacementIndexStart = output.getPageCount();
      setReorderStatus(`Adding replacement index: ${reorderState.replacementIndexFile.name}…`);
      const result = await appendIndexPdfToOutput(output, reorderState.replacementIndexFile);
      replacementIndexPageCount = result.pageCount;
      replacementInserted = true;
    };

    if (reorderState.frontMatterPageIndexes.length) {
      setReorderStatus(`Processing ${reorderState.frontMatterPageIndexes.length} front-matter page${reorderState.frontMatterPageIndexes.length === 1 ? '' : 's'}…`);
      for (const sourcePageIndex of reorderState.frontMatterPageIndexes) {
        if (reorderState.replacementIndexFile && oldIndexPageSet.has(sourcePageIndex)) {
          await appendReplacementIndex();
          continue;
        }
        const [copiedPage] = await output.copyPages(source, [sourcePageIndex]);
        output.addPage(copiedPage);
        const outputPageIndex = output.getPageCount() - 1;
        frontMatterMap.set(sourcePageIndex, outputPageIndex);
        if (!reorderState.replacementIndexFile && oldIndexPageSet.has(sourcePageIndex)) {
          removeLinkAnnotations(output, copiedPage);
        }
      }
    }

    if (reorderState.replacementIndexFile && !replacementInserted) {
      await appendReplacementIndex();
    }

    if (els.reorderCreateBookmarks.checked) {
      if (reorderState.replacementIndexFile && replacementIndexStart != null) {
        bookmarks.push({ title: 'INDEX', pageIndex: replacementIndexStart });
      } else {
        const indexPages = getReorderOldIndexPageIndexes();
        if (indexPages.length) {
          const indexOutputPage = frontMatterMap.get(indexPages[0]);
          if (indexOutputPage != null) bookmarks.push({ title: 'INDEX', pageIndex: indexOutputPage });
        }
      }
    }

    for (let i = 0; i < reorderState.items.length; i += 1) {
      const position = i + 1;
      const item = reorderState.items[i];
      const sourceTab = reorderState.sourceTabs[i];
      const tabPageIndex = output.getPageCount();

      setReorderStatus(`Adding positional TAB ${position}…`);
      await appendSourcePageIndexes(output, source, [sourceTab.pageIndex]);

      if (els.reorderCreateBookmarks.checked) {
        bookmarks.push({ title: sourceTab.label || `TAB ${position}`, pageIndex: tabPageIndex });
      }

      const contentPageIndex = output.getPageCount();
      contentPageTargets.push(contentPageIndex);
      setReorderStatus(`Adding content from original position ${item.originalPosition} at final position ${position}…`);
      await appendSourcePageIndexes(output, source, item.sourceSection.pageIndexes);
    }

    const indexState = getReorderIndexLinkState();
    const activeIndexRows = getReorderActiveIndexLinks();
    let refreshedLinks = 0;
    if (indexState.canRefresh && els.reorderRefreshIndexLinks.checked) {
      const actionWord = reorderState.replacementIndexFile ? 'Creating' : 'Refreshing';
      setReorderStatus(`${actionWord} ${activeIndexRows.length} index link${activeIndexRows.length === 1 ? '' : 's'}…`);
      activeIndexRows.forEach((row) => {
        const sourcePageIndex = reorderState.replacementIndexFile
          ? (replacementIndexStart == null ? null : replacementIndexStart + row.pageIndex)
          : frontMatterMap.get(row.pageIndex);
        const targetPageIndex = contentPageTargets[row.destinationPosition - 1];
        if (sourcePageIndex == null || targetPageIndex == null) throw new Error(`Could not map index entry to final position ${row.destinationPosition}.`);
        addInternalLinkAnnotation(output, output.getPage(sourcePageIndex), row.rect, targetPageIndex);
        refreshedLinks += 1;
      });
    }

    if (bookmarks.length) {
      setReorderStatus(`Creating ${bookmarks.length} top-level bookmark${bookmarks.length === 1 ? '' : 's'}…`);
      addTopLevelBookmarks(output, bookmarks);
    }

    setReorderStatus('Saving reordered binder…');
    const bytes = await output.save();
    const filename = sanitizePdfFilename(els.reorderOutputName.value || 'Reordered Binder.pdf');
    downloadBlob(new Blob([bytes], { type: 'application/pdf' }), filename);

    const movedCount = reorderState.items.filter((item, index) => item.originalPosition !== index + 1).length;
    let indexPhrase = '';
    if (reorderState.replacementIndexFile) {
      indexPhrase = ` Replacement index applied (${replacementIndexPageCount} page${replacementIndexPageCount === 1 ? '' : 's'})`;
      if (refreshedLinks) indexPhrase += ` with ${refreshedLinks} linked entr${refreshedLinks === 1 ? 'y' : 'ies'}`;
      indexPhrase += '.';
    } else if (refreshedLinks) {
      indexPhrase = ` ${refreshedLinks} existing index entr${refreshedLinks === 1 ? 'y was' : 'ies were'} relinked to the reordered content.`;
    }
    setReorderStatus(`Reordered binder downloaded. ${movedCount} content block${movedCount === 1 ? '' : 's'} changed position.${indexPhrase}`, 'success');
  } catch (error) {
    console.error(error);
    setReorderStatus(`Could not build the reordered binder: ${friendlyError(error)}`, 'error');
  } finally {
    const currentValidation = validateReorderPlan();
    els.reorderDownload.disabled = !currentValidation.ok;
  }
}

function setSyncStatus(message, type = '') {
  els.syncStatus.textContent = message;
  els.syncStatus.className = `status ${type}`.trim();
}

function setReorderStatus(message, type = '') {
  els.reorderStatus.textContent = message;
  els.reorderStatus.className = `status ${type}`.trim();
}

function setInsertStatus(message, type = '') {
  els.insertStatus.textContent = message;
  els.insertStatus.className = `status ${type}`.trim();
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
