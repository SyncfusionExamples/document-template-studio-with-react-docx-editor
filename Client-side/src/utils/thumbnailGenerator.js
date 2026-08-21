import { DocumentEditorContainer } from '@syncfusion/ej2-documenteditor';
import { DOCUMENT_EDITOR_SERVICE_URL } from '../data/sampleTemplates.js';

// ---------------------------------------------------------------------------
// Thumbnail generation for flows where the editor is NOT yet mounted.
//
// This module only owns the "no live editor exists yet" flow —
//   - `generateThumbnailFromFile(file)`  (upload flow: fired at the
//      moment the user confirms an upload, before the editor opens)
//   - `generateThumbnailFromUrl(url)`    (dashboard lazy-load: fired
//      for every built-in template that needs a thumbnail, again before
//      anyone opened the editor for it)
//
// Capture-from-live-editor (used by TemplateViewer's publish + save
// flows) lives next to the editor itself, in TemplateViewer.jsx, as a
// small file-local helper. That keeps each helper operating against
// exactly one DocumentEditor instance and avoids spinning up a second
// hidden editor container we're going to throw away.
//
// The export step (`documentEditor.exportAsImage(1, 'image/png')` then
// read `image.src` after `onload`) follows the Syncfusion reference
// sample:
//   1. Load the .docx into the offscreen editor.
//   2. Hike `printDevicePixelRatio = 2` for crisp HiDPI captures.
//   3. Wait a tick for the editor's layout to settle (`setTimeout`
//      matches the reference sample).
//   4. Resolve the data URI from the returned `HTMLImageElement`'s
//      `.src` once its bitmap has decoded.
// ---------------------------------------------------------------------------

// Export page 1 of the given editor as a `data:image/png;base64,...` data
// URI. Reject (not resolve-empty) because every caller here goes
// through the offscreen-editor path that *must* succeed — empty output is
// a real failure.
function exportPageOneAsDataUri(editor, { settleMs = 500 } = {}) {
  if (!editor || typeof editor.exportAsImage !== 'function') {
    return Promise.reject(new Error('DocumentEditor is not ready'));
  }
  editor.documentEditorSettings.printDevicePixelRatio = 2;

  return new Promise((resolve, reject) => {
    let img;
    setTimeout(() => {
      try {
        img = editor.exportAsImage(1, 'image/png');
      } catch (err) {
        reject(err);
        return;
      }
      if (!img || !img.src) {
        reject(new Error('exportAsImage returned no bitmap'));
        return;
      }
      img.onload = () => resolve(img.src);
      img.onerror = (e) => reject(e);
      if (img.complete) resolve(img.src);
    }, settleMs);
  });
}

// Build an off-screen DocumentEditorContainer, wait for its `created`
// event (fires once the inner DocumentEditor is ready), and resolve with
// it. The host MUST be attached to the DOM and given a real size for the
// editor to lay pages out and for `exportAsImage` to produce real
// bitmaps — we push it off-screen with left:-99999px + visibility:hidden
// so it never paints to the user while still rendering internally.
function createOffscreenContainer() {
  const host = document.createElement('div');
  host.style.position = 'fixed';
  host.style.left = '-99999px';
  host.style.top = '0';
  host.style.width = '820px';
  host.style.height = '600px';
  host.style.visibility = 'hidden';
  document.body.appendChild(host);

  const container = new DocumentEditorContainer({
    height: '600px',
    width: '820px',
    enableToolbar: false,
    showPropertiesPane: false,
    // serviceUrl is still required by the container even though we
    // open() documents from pre-fetched SFDT (via the same-origin dev
    // proxy). Using the same centralized URL the rest of the studio
    // uses keeps the thumbnail helper aligned with the live editor.
    serviceUrl: DOCUMENT_EDITOR_SERVICE_URL,
  });

  return new Promise((resolve, reject) => {
    let settled = false;
    container.created = () => {
      if (settled) return;
      settled = true;
      resolve({ host, container });
    };
    container.appendTo(host);

    // Safety net — if `created` never fires after a moment, reject so
    // the caller doesn't hang forever. If the editor is already
    // present, resolve.
    setTimeout(() => {
      if (settled) return;
      settled = true;
      if (container.documentEditor) {
        resolve({ host, container });
      } else {
        reject(new Error('DocumentEditor failed to initialise in time'));
      }
    }, 4000);
  });
}

// Load a .docx (either a File or a fetch-able URL) into the editor. The
// SFDT is produced via the same-origin dev proxy (see studioStorage's
// fetchSfdtFromDocx) — calling the Syncfusion Import service directly
// from the browser is blocked by CORS, so we route through the proxy
// instead.
async function importDocxIntoEditor(editor, { file, url }) {
  const { fetchSfdtFromDocx } = await import('./studioStorage.js');
  const sfdt = await fetchSfdtFromDocx({ file, url });
  editor.open(sfdt);
  // Give the editor a tick to lay out the pages before we capture them.
  await new Promise((r) => setTimeout(r, 200));
}

// Shared inner for the dashboard-side routes: import the docx into the
// offscreen container, render page 1 as a PNG data URI via Syncfusion's
// `exportAsImage`, and resolve with `{ thumbnailDataUri, pageCount }`.
// Tears the host down before resolving so we never leak a fixed-position
// container onto the page.
async function generateFromOffscreen({ file, url }) {
  let host;
  let container;
  try {
    ({ host, container } = await createOffscreenContainer());
    const editor = container.documentEditor;
    if (!editor) throw new Error('Offscreen editor did not expose documentEditor');
    await importDocxIntoEditor(editor, { file, url });
    const thumbnailDataUri = await exportPageOneAsDataUri(editor);
    return { thumbnailDataUri, pageCount: editor.pageCount };
  } finally {
    try { if (container) container.destroy(); } catch { /* noop */ }
    if (host?.parentNode) host.parentNode.removeChild(host);
  }
}

// Load a .docx File into a hidden DocumentEditor and resolve with the
// captured page-1 thumbnail + page count.
export async function generateThumbnailFromFile(file) {
  return generateFromOffscreen({ file });
}

// Fetch a .docx by URL (used to generate thumbnails for the built-in
// sample templates too — requirement: no hardcoded SVG thumbnails),
// render page 1 as a PNG data URI, and resolve with page count.
export async function generateThumbnailFromUrl(url) {
  return generateFromOffscreen({ url });
}
