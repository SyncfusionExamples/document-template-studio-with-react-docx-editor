import { DocumentEditorContainer } from '@syncfusion/ej2-documenteditor';
import { DOCUMENT_EDITOR_SERVICE_URL } from '../data/sampleTemplates.js';

// Render a single page of the loaded DocumentEditor into an HTMLImageElement
// (PNG data URI). Only page 1 is captured — that's all a dashboard card
// thumbnail needs.
function exportPageAsImage(editor, pageNumber, format = 'image/png') {
  // Crisp capture — the help doc sets printDevicePixelRatio to 2.
  editor.documentEditorSettings.printDevicePixelRatio = 2;
  const img = editor.exportAsImage(pageNumber, format);
  return new Promise((resolve, reject) => {
    img.onload = () => resolve(img);
    img.onerror = (e) => reject(e);
  });
}

// Scale a single page image onto a canvas and return a PNG data URI
// (base64) suitable for <img src=...> or storage as a string in JSON.
function composeThumbnail(img, maxWidth = 240) {
  if (!img) return '';
  const ow = parseInt(String(img.style.width || img.width).replace('px', ''), 10) || img.naturalWidth || img.width;
  const oh = parseInt(String(img.style.height || img.height).replace('px', ''), 10) || img.naturalHeight || img.height;
  const ratio = maxWidth / ow;
  const w = Math.round(ow * ratio);
  const h = Math.round(oh * ratio);

  const canvas = document.createElement('canvas');
  canvas.width = w;
  canvas.height = h;
  const ctx = canvas.getContext('2d');
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  ctx.drawImage(img, 0, 0, w, h);
  return canvas.toDataURL('image/png');
}

// Build an off-screen DocumentEditorContainer, wait for its `created` event
// (fires once the inner DocumentEditor is ready), and resolve with it.
// The host MUST be attached to the DOM and given a real size for the editor
// to lay pages out and for exportAsImage to produce real bitmaps — we push
// it off-screen with left:-99999px + visibility:hidden so it never paints
// to the user while still rendering internally.
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
    // serviceUrl is still required by the container even though we open()
    // documents from pre-fetched SFDT (via the same-origin proxy). The
    // public Syncfusion service URL works fine for this offscreen helper.
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

    // Safety net — if `created` never fires after a moment, reject so the
    // caller doesn't hang forever. If the editor is already present, resolve.
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
// fetchSfdtFromDocx) — calling the Syncfusion Import service directly from
// the browser is blocked by CORS, so we route through the proxy instead.
async function importDocxIntoEditor(editor, { file, url }) {
  const { fetchSfdtFromDocx } = await import('./studioStorage.js');
  const sfdt = await fetchSfdtFromDocx({ file, url });
  editor.open(sfdt);
  // Give the editor a tick to lay out the pages before we capture them.
  await new Promise((r) => setTimeout(r, 200));
}

// Load a .docx File into a hidden DocumentEditor, render page 1 as an
// image, and resolve with { thumbnailDataUri, pageCount }.
export async function generateThumbnailFromFile(file, { onProgress } = {}) {
  let host;
  let container;
  try {
    ({ host, container } = await createOffscreenContainer());
    const editor = container.documentEditor;
    await importDocxIntoEditor(editor, { file });
    if (onProgress) onProgress('rendering');
    const img = await exportPageAsImage(editor, 1, 'image/png');
    return { thumbnailDataUri: composeThumbnail(img, 240), pageCount: editor.pageCount };
  } finally {
    try { if (container) container.destroy(); } catch { /* noop */ }
    if (host?.parentNode) host.parentNode.removeChild(host);
  }
}

// Fetch a .docx by URL (used to generate thumbnails for the built-in sample
// templates too — requirement: no hardcoded SVG thumbnails), render page 1
// as an image, and resolve with { thumbnailDataUri, pageCount }.
export async function generateThumbnailFromUrl(url, { onProgress } = {}) {
  let host;
  let container;
  try {
    ({ host, container } = await createOffscreenContainer());
    const editor = container.documentEditor;
    await importDocxIntoEditor(editor, { url });
    if (onProgress) onProgress('rendering');
    const img = await exportPageAsImage(editor, 1, 'image/png');
    return { thumbnailDataUri: composeThumbnail(img, 240), pageCount: editor.pageCount };
  } finally {
    try { if (container) container.destroy(); } catch { /* noop */ }
    if (host?.parentNode) host.parentNode.removeChild(host);
  }
}

// Re-generate a thumbnail from an EXISTING editor instance — useful after the
// user edits a template in the viewer and saves. Reuses the same exportAsImage
// pipeline without re-loading the docx. Only page 1 is captured.
export async function generateThumbnailFromEditor(editor, { onProgress } = {}) {
  if (!editor) return { thumbnailDataUri: '', pageCount: 0 };
  if (onProgress) onProgress('rendering');
  const img = await exportPageAsImage(editor, 1, 'image/png');
  return { thumbnailDataUri: composeThumbnail(img, 240), pageCount: editor.pageCount };
}
