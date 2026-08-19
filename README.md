# PDF Binder Tool — browser-only MVP v0.7.2

A client-side PDF workspace for splitting combined PDFs, extracting physical tab divider pages, editing a document manifest, rebuilding binders with positional tabs, and adding clickable links from a supplied index PDF to the final document contents.

The intended deployment is a static website. Users do not need Python, Node, or any desktop installation. The PDF files selected by the user are read and modified in browser memory; this project does not include a server-side upload endpoint.

## Current workflows

### Split PDF

1. Open a local PDF.
2. Review pages in the thumbnail rail and preview pane.
3. Create divisions in either of two ways:
   - **Split after page** — makes a normal document boundary and retains the selected page.
   - **Mark as tab divider** — excludes that physical divider page from the content PDFs.
4. Optionally use **Detect tab pages** to find text such as `TAB 1` or `TAB A`.
5. Name the resulting documents in the current-section editor or manifest table.
6. Download a ZIP containing:
   - `documents/` — clean split content PDFs
   - `tabs/` — separately preserved divider PDFs, when enabled
   - `binder-manifest.json`
   - `binder-manifest.csv`
   - `binder-manifest.xlsx`

Tabbed-binder behavior intentionally omits unmarked front matter before the first tab divider. If leading pages should be retained as documents, make a normal split cut in that leading material.

### Manifest editor

Version 0.5 added a more complete editor layer:

- Edit Label, Description, and Output Filename.
- The preview pane shows the manifest fields for the currently selected output section.
- Automatic filenames use the configurable template:

  `{order:02} - {description}.pdf`

- Supported filename tokens:
  - `{order}`
  - `{order:02}`
  - `{label}`
  - `{description}`
  - `{startPage}`
  - `{endPage}`
  - `{original}`
- Auto-generated filenames renumber when section order changes because a new split is inserted.
- Manually editing an individual filename locks that filename. **Apply template to all** clears those manual locks.
- Validation warns about blank descriptions, filename sanitization, unusually large sections, duplicate content filenames, and duplicate preserved-tab filenames. Duplicate filenames block splitting because they would collide inside the ZIP.

### Excel round trip

The app can export the split manifest as `.xlsx` and import edited `.xlsx` or `.json` manifests.

When a source PDF is already open, spreadsheet import updates the current sections by **Order** and applies Label, Description, and Filename. Spreadsheet edits to page/boundary columns do not alter source-PDF split points. This is deliberate: page extraction remains controlled by the visual markers in the app.

The workbook includes a `Manifest` sheet and an `Instructions` sheet.

### Build Binder

1. Add content PDFs.
2. Optionally add supplied tab PDFs.
3. Tabs are positional assets: first tab = position 1, second tab = position 2, etc.
4. Reorder content PDFs with the arrow buttons or drag handles.
5. Edit the content description for future index/manifest use.
6. Import a JSON or Excel manifest to restore filename order and descriptions when filenames match.
7. Export the final binder order to Excel if desired.
8. Build and download the merged binder PDF.

If positional tab insertion is enabled, the app requires at least as many tab PDFs as content PDFs. Extra tab PDFs are allowed and ignored.

### Tab bookmarks (v0.6)

When supplied positional tabs are loaded, the app uses PDF.js to inspect each tab PDF for an existing document-outline/bookmark title. The Build Binder table shows the title it found.

- If an existing bookmark title is readable, that title is reused.
- If the tab PDF has no readable bookmark, the filename without `.pdf` is used as the fallback title.
- **Create bookmarks for positional tabs** is enabled by default when tab insertion is active.
- On rebuild, the app creates a fresh top-level bookmark for each inserted positional tab, pointing to the exact tab page in the newly merged binder.
- The original bookmark object/destination is not copied, because its page reference belongs to the source PDF; only the useful title is reused.
- The final-order Excel export includes a `Bookmark Title` column and records whether tab bookmarks were enabled.

This bookmark creation uses pdf-lib's low-level PDF outline objects because pdf-lib does not expose a high-level bookmark API.

### Supplied index PDF + internal links (v0.7.1)

The Build Binder screen can now accept a separate index PDF that is inserted at the very front of the final binder. The intended index layout is the office's regular three-column Word/PDF table: numbered row, Description, and Date.

- PDF.js extracts positioned text from every index page and looks for numbered table rows.
- If the supplied index already contains link annotations, their rectangles are reused when they line up with detected rows; otherwise the app estimates the Description-cell rectangle from the table text/column positions.
- The mapping is positional: index row 1 -> content PDF at binder position 1, row 2 -> position 2, etc.
- The link destination is the **first page of the content PDF**, not the green positional tab immediately before it.
- Any existing Link annotations on the copied index pages are removed before new links are written. This avoids carrying stale destinations from an older binder into the rebuilt binder.
- For this first version, linked-index creation requires exactly one detected row numbered 1 through N for each of the N loaded content PDFs. If detection/counts do not match, Build is disabled until index linking is turned off or the inputs are corrected.
- Multi-page indexes are supported as long as the numbered rows can be extracted as text. Scanned/image-only index pages are not yet supported because this version does not include OCR.

The Index & Links table previews the detected Description/Date text and shows which current binder position each index row will target. Reordering content PDFs updates those destinations automatically.

## Libraries

- PDF.js — browser PDF rendering and text extraction
- pdf-lib — PDF page copying, splitting, and merging
- JSZip — ZIP output
- SheetJS (`xlsx`) — Excel import/export
- Vite — development and static production build

## Development

```bash
npm install
npm run dev
```

Production build:

```bash
npm run build
```

The static site is produced in `dist/`. `vite.config.js` uses a relative base so the build can be hosted under a subdirectory.

## Manifest schema

The current JSON manifest identifies itself as:

```json
{
  "schema": "pdf-binder-tool/5",
  "tabSemantics": "position"
}
```

Each section records its source-page range, boundary type, label, description, filename, and whether the filename came from the template or was manually set.

## Suggested next increments

- Manual correction tools for index-row/link rectangles when automatic table detection is imperfect.
- Automatic index-page generation from the final binder order and calculated pagination.
- Optional front-matter preservation as a named component rather than relying on a leading normal cut.
- Page-count calculation for loaded build PDFs so the final-order view can preview final binder pagination before assembly.
- Optional saved tab-library handling using the browser File System Access API where supported, while retaining download-based fallback behavior.



## v0.7.2 — Fit Page destinations

Generated navigation now uses the PDF `/Fit` destination type:

- Clicking a generated index Description link opens the first content page after its positional tab using **Fit Page**.
- Clicking a generated positional TAB bookmark opens that tab page using **Fit Page**.
- This controls the destination page zoom so the whole page fits in the viewer window; it does not force the PDF application into full-screen mode or hide its panels/toolbars.

No split, merge, ordering, index-detection, or bookmark-title behavior changed in this release.

## v0.7.1 fix

- Updated PDF.js cleanup to call `PDFDocumentLoadingTask.destroy()` rather than `PDFDocumentProxy.destroy()`.
- Cleanup failures are now isolated from index/bookmark analysis, so a cleanup API mismatch cannot falsely report that index detection failed.
