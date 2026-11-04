# PDF Binder Tool — browser-only MVP v0.14.0


## v0.14.0 — Match loose content PDFs to a supplied index

Build Binder now includes an optional **Match PDFs to index** assistant. This is intended for building a binder from loose PDFs when the supplied index already defines the desired final order.

- Load the content PDFs and the finished index PDF as usual.
- Click **Match PDFs to index** in the Index & links panel.
- The app compares each numbered index Description/Date against the loaded PDF filenames and current descriptions.
- Matching is globally one-to-one: the proposal cannot assign the same PDF to two index rows.
- Common date-prefixed filenames such as `26.05.15 Email ...` are recognized for matching; the date prefix is compared separately when possible rather than being treated as part of the document title.
- Proposed matches are labeled Strong, Good, Review, or Uncertain.
- Every row has a PDF dropdown for manual correction. Choosing a PDF that is already assigned elsewhere swaps the two assignments so the mapping stays one-to-one.
- Nothing is reordered until **Apply order** is clicked. The Final binder order remains editable afterward with drag-and-drop, arrows, or direct position entry.
- The assistant currently requires exactly one sequential numbered index row per loaded content PDF. Atypical indexes should continue to use Manual mapping for links and manual content ordering.

This feature changes only content order. It does not alter the index artwork, the supplied tab PDFs, or the source content PDFs.

## v0.13.1 index-number detection fix

Automatic index parsing now infers the actual number-column alignment before accepting numeric row markers. This prevents numeric text at the beginning of a Description, such as `504 Plan`, from being misidentified as an extra index row.


## v0.13.1 — Sync Binder to New Index

This release adds a fifth top-level workflow, **Sync to New Index**, for updating an existing tabbed binder from a newly revised index and a set of new PDFs. The new index is treated as the intended final order; the old binder supplies existing content; newly supplied PDFs fill rows that do not align to old content.

- Open the existing binder, then load the new index PDF and any newly added documents.
- The app compares the old and new numbered index descriptions with sequence-aware fuzzy matching, so large blocks shifted by inserted rows remain aligned instead of being treated as unrelated documents. Date matches are used as an additional signal when available.
- Unmatched new-index rows are compared to the filenames of the supplied new PDFs. Existing filename conventions (`^` for `:` and `_` for `/`) are normalized for matching.
- Every proposed row remains editable through a Source dropdown. Automatic matches are labeled Strong, Good, Review, or Uncertain; manual overrides are labeled Manual.
- Each source PDF/content block can be used only once. Build is blocked for unassigned final rows, duplicate source assignments, unused newly supplied PDFs, missing positional tabs, or an index whose numbered rows are not sequential and unique.
- Existing binder documents that do not appear in the new index are never silently deleted. The app lists them and requires an explicit omission approval before Build is enabled.
- The new index replaces detected old index pages while unrelated front matter is preserved. Its Description links are regenerated to the first content page after each final positional tab, using Fit Page.
- INDEX is regenerated as the first bookmark followed by positional TAB bookmarks.
- If the new index creates positions beyond the original binder's tab count, extra tab PDFs are required just as in Insert Documents.

### Typical sync workflow

1. Open the old completed binder.
2. Add the updated index PDF.
3. Add the PDFs for newly inserted index entries.
4. Add any extra positional tab PDFs required by the larger final binder.
5. Review the proposed mapping. Correct any Review/Uncertain row with the Source dropdown.
6. If the new index intentionally removes old documents, review the omission list and approve it.
7. Build the synced binder.

This first Sync implementation intentionally requires parseable sequential numbered rows in both the old and new indexes because those descriptions are the evidence used for automatic structural comparison. Atypical/manual index mapping remains available in the existing Build, Insert, and Reorder workflows.



## v0.12.0 — Manual mapping for atypical indexes

This release adds a fallback index-link editor for indexes that do not fit the usual sequential No. / Description / Date table. Automatic parsing remains the default. Use **Manual mapping** only when the index structure is unusual or the automatic destinations need correction.

Manual mapping is available in **Build Binder**, **Insert Documents**, and **Reorder Binder**. It can:

- render the actual index/front-matter PDF page in a mapping dialog;
- start from automatically detected rows and, when available, raw existing link rectangles even when the numbered-row parser cannot understand the index;
- reassign a detected entry to any final tab position;
- mark a detected region as **Heading** or **Ignore** so no PDF link is created there;
- draw a new clickable rectangle directly over index text when no row was detected;
- extract the text inside a newly drawn rectangle when PDF text is available;
- allow non-sequential mappings and multiple entries pointing to the same tab position;
- preserve the normal Fit Page destination behavior;
- use manually identified existing index pages when creating the top-level `INDEX` bookmark.

The automatic parser and existing-link rectangle reuse behavior are unchanged when manual mapping is not applied. To return to the automatic interpretation while the mapper is open, use **Reset to automatic regions**.

### Manual mapping workflow

1. Load the index or existing binder as usual.
2. Click **Manual mapping** in the relevant Index panel.
3. Review the highlighted regions.
4. Change a region's Type to Link, Heading, or Ignore and choose the desired final position for links.
5. If an entry was not detected, click **Draw link region** and drag over the desired index text.
6. Click **Apply manual mapping**.
7. Build the binder normally.

Manual mapping is intentionally permissive. It does not require rows 1 through N or a one-to-one relationship between index entries and tabs. The only hard requirement is that every active Link region point to a valid final binder position.

## v0.11.1 — Text-sized index link geometry

- Existing index link rectangles are still reused exactly when the supplied index already contains usable link annotations.
- When a supplied/replacement index has no usable links, generated Description links now use the actual detected Description text footprint rather than approximately filling the entire middle table cell.
- Multi-line descriptions use the union of all detected Description text fragments, with small click-friendly padding. The generated rectangle is clamped to the detected row and kept out of the Date column.
- This is intentionally a narrow, reversible experiment; PDF splitting, tab positioning, bookmarks, Fit Page destinations, and binder ordering are unchanged.

## v0.11.0 — Reorder Binder workflow

A fourth top-level mode, **Reorder Binder**, supports rearranging the existing document blocks inside an already assembled tabbed binder without first splitting them to disk.

- Open an existing tabbed binder; the app detects physical TAB pages, front matter, content ranges, and numbered index rows using the same analysis used by Insert Documents.
- Existing content blocks can be reordered by drag-and-drop, ↑/↓ controls, or by typing a new final position. Moving original position 40 to final position 13 shifts the intervening content automatically.
- TAB pages remain positional assets. Moving the content originally under TAB 40 to final position 13 pairs that content with the existing physical TAB 13 page; the old TAB 40 page does not travel with it.
- Original positions and source page ranges remain visible in the reorder table for verification. A **Reset order** button restores the source order.
- The number of positions does not change, so no extra tab PDFs are required.
- An updated index PDF can be supplied with **Replace index PDF**. Its rows must reconcile exactly with positions 1 through N; detected old index pages are replaced while unrelated front matter is preserved.
- Existing or replacement Description-cell links are regenerated to the first content page after each final positional tab and use Fit Page destinations.
- INDEX remains the first bookmark when an index is detected/supplied, followed by fresh positional TAB bookmarks.
- If the original index is kept after content has been reordered, the app warns that its visible descriptions/dates are still in the old order even though its internal links can be refreshed.



## v0.10.1 — Replacement index in Insert Documents

Insert Documents can now accept an updated index PDF directly.

- Use **Replace index PDF** after defining the insertion/replacement operations.
- The replacement index is analyzed with the same numbered-row and Description-cell detection used by Build Binder.
- Its numbered rows must reconcile exactly with the finished binder positions (1 through N); a mismatch blocks the build rather than creating a plausible but incorrect index.
- When old index pages are detected in the existing binder, those pages are removed and the replacement index is inserted at the same front-matter location. Other front matter is preserved.
- If no old index pages are detected, the replacement index is inserted after the preserved front matter and immediately before TAB 1, with a warning.
- Existing link annotations in the supplied replacement index are removed and recreated against the finished binder. Each Description row points to the first content page after its positional tab and uses Fit Page.
- When bookmark rebuilding is enabled, the replacement index becomes the first top-level bookmark named `INDEX`; positional TAB bookmarks follow it.
- **Use existing index** clears the replacement and returns to the original v0.10.0 behavior.

## v0.9.1 filename-placeholder cleanup

- Imported filename/path descriptions now interpret the office-safe placeholders `^` → `:` and `_` → `/`.
- Generated PDF filenames remain Windows/macOS-safe because the existing filename sanitizer converts those restored description characters when producing a filename.
- Dotted date prefixes such as `26.05.15` are no longer mistaken for removable numeric ordering prefixes.

## v0.9.0 index/order workflow update

- A supplied Build Binder index PDF now becomes the **first top-level bookmark**, always titled `INDEX`, pointing to the first index page with Fit Page. Positional TAB bookmarks follow it when tab bookmark generation is enabled.
- The Split PDF bulk metadata dialog is now **Paste/import index** and accepts Word or Excel clipboard tables.
- Two-column `Filename/Description | Date` tables are supported without an explicit number column. When numbering is absent, row order is treated as positions 1 through N and must still reconcile exactly with the current output-section count before Apply is enabled.
- Two-column filename/path input is cleaned for the Description preview: surrounding quotes and folder paths are removed, `.pdf` is removed, `^` is restored to `:` and `_` is restored to `/`, and a simple leading numeric ordering prefix is removed. Dotted date prefixes such as `26.05.15` are preserved. The preview remains the verification step before applying anything.
- The same bulk metadata dialog can directly import `.xlsx`, `.xls`, `.csv`, or `.tsv`; the first worksheet/table is used.
- In Build Binder, the Order column is editable. Type a destination position (for example change item 40 to `13`) and press Enter or leave the field; that PDF moves to position 13 and the intervening PDFs shift automatically. Drag-and-drop and arrow controls remain available.


## v0.8.2 metadata cleanup

- Generated split PDFs, preserved tab PDFs, and rebuilt binders are created with pdf-lib automatic metadata updates disabled.
- The output no longer intentionally includes pdf-lib in the PDF `Creator` or `Producer` fields.
- Creation and modification timestamps are still written to generated PDFs.
- Source PDFs are loaded with metadata updates disabled as well; source metadata is not used to brand the generated binder.

## v0.8.1 usability update

- Added **Next tab** in the Split PDF preview. It jumps to the next marked/detected tab divider and scrolls that thumbnail into view.
- Compacted the desktop header, toolbar, panel headings, and current-section editor.
- Made the split workspace use the available browser viewport height, so the PDF preview grows on larger screens instead of staying fixed at 500px.
- The current section's Label, Description, Date, and Filename fit on one row on wider desktop screens.


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

- Edit Label, Description, Date, and Output Filename.
- The preview pane shows the manifest fields for the currently selected output section.
- Automatic filenames use the configurable template:

  `{order:02} - {description}.pdf`

- Supported filename tokens:
  - `{order}`
  - `{order:02}`
  - `{label}`
  - `{description}`
  - `{date}`
  - `{startPage}`
  - `{endPage}`
  - `{original}`
- Auto-generated filenames renumber when section order changes because a new split is inserted.
- Manually editing an individual filename locks that filename. **Apply template to all** clears those manual locks.
- Validation warns about blank descriptions, filename sanitization, unusually large sections, duplicate content filenames, and duplicate preserved-tab filenames. Duplicate filenames block splitting because they would collide inside the ZIP.
- Auto-generated filenames are normalized for conservative Windows/macOS compatibility and capped at 120 characters including `.pdf`. Descriptions themselves are never shortened.

### Excel round trip

The app can export the split manifest as `.xlsx` and import edited `.xlsx` or `.json` manifests.

When a source PDF is already open, spreadsheet import updates the current sections by **Order** and applies Label, Description, Date, and Filename. Spreadsheet edits to page/boundary columns do not alter source-PDF split points. This is deliberate: page extraction remains controlled by the visual markers in the app.

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
- If the supplied index already contains link annotations, their rectangles are reused when they line up with detected rows; otherwise the app sizes the clickable rectangle to the detected Description text footprint, with small click-friendly padding.
- The mapping is positional: index row 1 -> content PDF at binder position 1, row 2 -> position 2, etc.
- The link destination is the **first page of the content PDF**, not the green positional tab immediately before it.
- Any existing Link annotations on the copied index pages are removed before new links are written. This avoids carrying stale destinations from an older binder into the rebuilt binder.
- For this first version, linked-index creation requires exactly one detected row numbered 1 through N for each of the N loaded content PDFs. If detection/counts do not match, Build is disabled until index linking is turned off or the inputs are corrected.
- Multi-page indexes are supported as long as the numbered rows can be extracted as text. Scanned/image-only index pages are not yet supported because this version does not include OCR.

The Index & Links table previews the detected Description/Date text and shows which current binder position each index row will target. Reordering content PDFs updates those destinations automatically.


## Paste/import index metadata (expanded in v0.9.0)

The Split PDF manifest includes **Paste/import index** for bulk naming. It supports the office's three-column Word index (`No.`, `Description`, `Date`) and simpler two-column Excel lists (`Filename/Description`, `Date`).

- Copy table rows in Word or Excel, open **Paste/import index**, click the paste box, and use Ctrl+V/Cmd+V; or import an Excel/CSV/TSV file directly.
- The browser reads the clipboard's HTML table structure when Word supplies it. If that is unavailable, it falls back to tab-delimited or numbered plain text.
- Explicit numbering is validated when present. For two-column lists without numbers, positions 1 through N are inferred from row order. In either case, the preview must reconcile exactly with the N current output sections; missing, duplicate, extra, or blank-description rows block **Apply entries**.
- `Description` and `Date` populate the corresponding manifest fields. Positional tab labels remain `TAB 1`, `TAB 2`, etc.
- By default, filenames are regenerated from the current filename template after the descriptions are applied. This can be unchecked if existing filenames should remain untouched.
- The preview shows the proposed filename for each row and flags filenames that required cross-platform character replacement, Windows reserved-name handling, or truncation.
- The app uses a conservative 120-character filename limit (including `.pdf`) even though common filesystems may permit longer names, reducing problems with longer folder paths, ZIP extraction, sync folders, and document-management systems.
- Pasted descriptions are never truncated or sanitized; only filenames are modified.

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
  "schema": "pdf-binder-tool/6",
  "tabSemantics": "position"
}
```

Each section records its source-page range, boundary type, label, description, date, filename, and whether the filename came from the template or was manually set.

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

## v0.10.0 — Insert Documents workflow

A third top-level mode, **Insert Documents**, supports editing an already assembled tabbed binder without first downloading all of its component PDFs.

1. Open the existing binder. The app scans for physical `TAB 1`, `TAB 2`, etc. divider pages using PDF.js text extraction.
2. Pages before the first detected TAB are preserved as front matter. Each TAB page remains a positional asset; the pages following it through the next TAB become that position's existing content.
3. Add one or more new content PDFs. For each PDF, choose a **final position** and one of two actions:
   - **Replace** — replaces the content occupying that final position while retaining that position's tab.
   - **Insert & shift** — creates a new position and shifts later existing content down.
4. Multiple operations are resolved against the finished binder, not sequentially against changing intermediate positions. Duplicate target positions are blocked.
5. Existing source tab pages are reused positionally. If insert-and-shift operations make the finished binder longer than the original, add extra tab PDFs named/mapped by position, such as `TAB 21.pdf`.
6. The Proposed Final Binder table shows the resolved order, source page ranges for retained content, whether a row is new/replaced/shifted, and the tab asset that will be used.
7. On rebuild, the app recreates top-level TAB bookmarks using Fit Page. If numbered index rows are detected in preserved front matter, an INDEX bookmark is also recreated at the detected index page.

### Index handling in Insert Documents

Existing internal index links are treated conservatively because edits can invalidate their destinations.

- Link annotations on detected existing index pages are removed during rebuild rather than copied blindly.
- If the existing index still has exactly rows 1 through N for the finished binder, **Create/refresh index links** can recreate the Description-cell links so they point to the first content page after each final positional tab.
- For changed indexes, use **Replace index PDF**. The replacement is validated against the finished binder, substitutes for the detected old index pages, receives fresh internal links, and becomes the `INDEX` bookmark.
- If an Insert & shift operation changes the number of positions, supplying a replacement index with the new row count is the recommended workflow.
- Replacement-only edits can still reuse the existing index when its visible Description/Date text remains accurate.

This first Insert Documents version intentionally requires extractable TAB labels. If a binder's physical tabs are image-only or cannot be reliably detected, use the existing Split PDF / Build Binder workflow where divider pages can be marked manually.
