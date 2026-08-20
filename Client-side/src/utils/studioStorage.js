// studioStorage.js
// Client-side API for talking to the Vite dev middleware
// (vite-plugin-studio-templates.js). Handles:
//   - uploading a new .docx file (+ thumbnail) so it lands in src/data/
//   - saving an edited .docx back to the data folder (requirement 3)
//   - removing a template's files (requirement 2)
//   - listing templates already persisted on disk (built-in samples + uploads)
//
// All endpoints are dev-only (the plugin is registered with apply:'serve').
// In a production build these calls would 404 — but the sample app is meant
// to be run with `npm run dev`, exactly like the shipped README instructs.

const API = '/studio-api';

// Build a FormData payload for upload / save.
// `thumbnailDataUri` is optional ("data:image/png;base64,...").
function buildFormData({ name, type, description, fieldKeys, id, docxFile, thumbnailDataUri }) {
  const fd = new FormData();
  if (id) fd.append('id', id);
  if (name) fd.append('name', name);
  if (type) fd.append('type', type);
  if (description) fd.append('description', description);
  if (fieldKeys) fd.append('fieldKeys', JSON.stringify(fieldKeys));
  if (docxFile) fd.append('docx', docxFile, docxFile.name || 'template.docx');
  if (thumbnailDataUri) {
    // Pass the data URI as a pseudo-file so the multipart parser treats it
    // as a "file" part. The server strips the base64 prefix and writes bytes.
    const blob = new Blob([thumbnailDataUri], { type: 'text/plain' });
    fd.append('thumbnail', blob, 'thumbnail.txt');
  }
  return fd;
}

// POST a new template to the dev server. Returns the saved template object
// (including the docxUrl + thumbnailUrl that should be used on the client).
export async function uploadTemplate({ file, name, type = 'General', description = '', fieldKeys = [], thumbnailDataUri = '' }) {
  const fd = buildFormData({
    name, type, description, fieldKeys,
    docxFile: file,
    thumbnailDataUri,
  });
  const res = await fetch(`${API}/upload`, { method: 'POST', body: fd });
  if (!res.ok) throw new Error(`Upload failed (${res.status})`);
  const json = await res.json();
  if (!json.ok) throw new Error(json.error || 'Upload failed');
  return json.template;
}

// Import a .docx into SFDT via the same-origin dev proxy. Calling the
// Syncfusion Import service directly from the browser is blocked by CORS,
// so the Vite dev middleware forwards the request server-side and returns
// the SFDT text. The browser then calls DocumentEditor.open(sfdt) with it.
//
// Accepts either a File or a URL (for built-in templates served by Vite).
export async function fetchSfdtFromDocx({ file, url, name }) {
  let blob;
  let filename = 'template.docx';
  // baseName is the document name WITHOUT the .docx extension. The Vite
  // dev plugin forwards it to the backend's Import endpoint as a
  // multipart `FileName` field, and the Save endpoint uses the same
  // name to overwrite <name>.docx in wwwroot/Templates/.
  let baseName = (name || '').replace(/\.docx$/i, '').trim();
  if (file) {
    blob = file;
    filename = file.name || filename;
    if (!baseName) baseName = filename.replace(/\.docx$/i, '');
  } else if (url) {
    // `cache: 'no-store'` is critical for the Save→Reopen round-trip:
    // after Save overwrites <slug>.docx on the server, the browser's
    // HTTP cache (and Vite's proxy cache) would otherwise hand back the
    // pre-save bytes, and the reopened editor would look unedited.
    // We also append a cache-busting query so any intermediate cache
    // (Vite's dev proxy in particular) sees it as a fresh resource.
    const cacheBustUrl = url + (url.includes('?') ? '&' : '?') + 't=' + Date.now().toString(36);
    const res = await fetch(cacheBustUrl, { cache: 'no-store' });
    blob = await res.blob();
    // Derive a filename from the URL tail (use the original url so the
    // import form field carries the real .docx name, not the cache-bust).
    try { filename = decodeURIComponent(url.split('/').pop()) || filename; } catch { /* keep default */ }
    if (!baseName) baseName = filename.replace(/\.docx$/i, '');
  } else {
    throw new Error('fetchSfdtFromDocx: file or url required');
  }
  const fd = new FormData();
  fd.append('docx', blob, filename);
  if (baseName) fd.append('FileName', baseName);
  const res = await fetch(`${API}/import`, { method: 'POST', body: fd });
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err.error || `Import failed (${res.status})`);
  }
  return res.text();
}

// POST an edited template back to the dev server so the .docx (and any
// refreshed thumbnail) in src/data/ stays in sync with what is on screen.
// NOTE: the editor's Save button now bypasses this helper and calls
// saveTemplateToServer() below, which goes straight to the backend's
// DocumentEditorController.Save endpoint (no vite-plugin involvement).
// Kept here for backward compatibility (upload flow still relies on it
// for the initial .docx write).
export async function saveTemplate({ id, name, docxFile, thumbnailDataUri }) {
  const fd = buildFormData({ id, name, docxFile, thumbnailDataUri });
  const res = await fetch(`${API}/save`, { method: 'POST', body: fd });
  if (!res.ok) throw new Error(`Save failed (${res.status})`);
  const json = await res.json();
  if (!json.ok) throw new Error(json.error || 'Save failed');
  return json;
}

// ---------------------------------------------------------------------------
// Save via the backend's DocumentEditorController.Save endpoint.
//
// The editor's "Save Template" button serializes the live document to
// Syncfusion's SFDT JSON and POSTs it to
//   http://localhost:5212/api/documenteditor/Save
// with a SaveParameter body of:
//   {
//     Content: <SFDT JSON string>,
//     FileName: <document name WITHOUT the .docx extension>,
//     Format:   'Docx'
//   }
//
// The backend writes/overwrites the .docx in Server-side/wwwroot/Templates/
// using FileMode.OpenOrCreate, so calling Save again with the same
// FileName replaces the existing file (no "Save as new file" behaviour).
//
// This deliberately bypasses the Vite dev plugin so the editor hits the
// authoritative save path (same one a production app would call).
// ---------------------------------------------------------------------------
const DOC_EDITOR_SAVE_URL = 'http://localhost:5212/api/DocumentEditor/Save';

function stripDocxExtension(name) {
  if (!name) return 'Document';
  // The backend's Save() expects the FileName it sees in the SaveParameter
  // — it derives the format via Path.GetExtension. We send the name WITHOUT
  // an extension and pin Format: 'Docx' so the backend never sees
  // ambiguous extensions and the overwrite always targets <name>.docx.
  return String(name).replace(/\.docx$/i, '').trim() || 'Document';
}

export async function saveTemplateToServer({ sfdtContent, documentName, format = 'Docx' }) {
  if (typeof sfdtContent !== 'string' || sfdtContent.length === 0) {
    throw new Error('saveTemplateToServer: sfdtContent is required');
  }
  const fileName = stripDocxExtension(documentName);
  const payload = {
    Content: sfdtContent,
    FileName: fileName,
    Format: format,
  };
  const res = await fetch(DOC_EDITOR_SAVE_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  });
  if (!res.ok) {
    let detail = '';
    try { detail = await res.text(); } catch { /* ignore */ }
    throw new Error(
      `Save failed (${res.status})${detail ? `: ${detail.slice(0, 200)}` : ''}`,
    );
  }
  return { ok: true, fileName, format };
}

// ---------------------------------------------------------------------------
// Mail Merge preview via the backend's DocumentEditorController.MailMerge
// endpoint.
//
// Mirrors the official Syncfusion online "Mail Merge" sample flow:
//   1. Caller exports the live DocumentEditor to a .docx BLOB via
//      `documentEditor.saveAsBlob('Docx')`.
//   2. Caller reads the blob as a Data URL (FileReader.readAsDataURL) so
//      the result is a base64 string prefixed with the data: URL scheme.
//   3. We POST `{ fileName, documentData, mailMergeData }` (ExportData on
//      the .NET side) to
//        http://localhost:5212/api/DocumentEditor/MailMerge
//      where:
//        fileName      = container.documentEditor.documentName + ".docx"
//        documentData  = the base64 (data: URL) string from FileReader
//        mailMergeData = JSON.stringify(userInputJsonObject)
//
// The backend merges the JSON data into the Word doc, then re-serialises
// the merged Word document back to Syncfusion's SFDT JSON, which we return
// to the caller so the editor can `open()` it to display the merged preview.
//
// NOTE: `documentData` is intentionally sent as the raw Data URL string
// (e.g. "data:application/vnd...;base64,...."); the backend's MailMerge
// controller strips the "data:...;base64," prefix before
// Convert.FromBase64String. The fileName MUST end with ".docx" because
// the backend uses it as the Word doc filename hint.
// ---------------------------------------------------------------------------
const DOC_EDITOR_MAILMERGE_URL = 'http://localhost:5212/api/DocumentEditor/MailMerge';

export async function mailMergePreview({ fileName, documentData, mailMergeData }) {
  if (typeof documentData !== 'string' || documentData.length === 0) {
    throw new Error('mailMergePreview: documentData (base64) is required');
  }
  if (typeof mailMergeData !== 'string' || mailMergeData.length === 0) {
    throw new Error('mailMergePreview: mailMergeData (JSON string) is required');
  }
  const safeFileName = (fileName || 'Document.docx').endsWith('.docx')
    ? (fileName || 'Document.docx')
    : `${(fileName || 'Document').replace(/\.docx$/i, '')}.docx`;
  const payload = {
    fileName: safeFileName,
    documentData,
    mailMergeData,
  };
  const res = await fetch(DOC_EDITOR_MAILMERGE_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json;charset=UTF-8' },
    body: JSON.stringify(payload),
  });
  if (res.status === 200) {
    // Backend returns the merged document as a SFDT JSON string body.
    const sfdt = await res.text();
    if (!sfdt) throw new Error('Mail merge returned an empty document.');
    return sfdt;
  }
  // Non-200: surface a readable error to the caller.
  let detail = '';
  try { detail = await res.text(); } catch { /* ignore */ }
  throw new Error(
    `Mail merge failed (${res.status})${detail ? `: ${detail.slice(0, 240)}` : ''}`,
  );
}

// Helper: read a Blob (from documentEditor.saveAsBlob('Docx')) as a base64
// Data URL string (the exact format the .NET MailMerge endpoint expects).
// Mirrors the FileReader.readAsDataURL pattern in the Syncfusion sample.
export function readBlobAsDataUrl(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(reader.error || new Error('FileReader failed'));
    reader.readAsDataURL(blob);
  });
}

// DELETE a template's .docx file from the server's wwwroot/Templates folder.
// The matching metadata entry is NOT removed here — the client manages its
// own single templates.json collection (see saveTemplatesCatalog) and is
// responsible for removing the entry from there before / after calling this.
export async function deleteTemplateFiles(id) {
  const fd = new FormData();
  fd.append('id', id);
  const res = await fetch(`${API}/delete`, { method: 'POST', body: fd });
  if (!res.ok) throw new Error(`Delete failed (${res.status})`);
  return res.json();
}

// Persist the client's single templates.json collection back to disk
// (Client-side/src/data/templates.json) so any change (upload, save,
// delete, custom-field add, ...) survives Vite dev server restarts.
// `catalog` is the full array of template metadata entries.
export async function saveTemplatesCatalog(catalog) {
  const res = await fetch(`${API}/catalog`, {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(catalog),
  });
  if (!res.ok) throw new Error(`Catalog save failed (${res.status})`);
  return res.json();
}

// ---------------------------------------------------------------------------
// Hidden built-in templates.
//
// The built-in .docx templates ship with the app (loaded from
// src/data/templates.json on every page mount), so deleting one in the UI
// only removes it from the current session's state. To make the deletion
// survive a reload / service restart, we keep a list of "hidden" template
// ids in window.localStorage. On startup App.jsx filters the catalog
// against this list, and the delete handler adds the id to it.
//
// Clearing localStorage (or calling clearHiddenBuiltInTemplates()) restores
// the built-ins.
// ---------------------------------------------------------------------------
const HIDDEN_BUILTINS_KEY = 'studio.hiddenBuiltInTemplates';

function readHiddenBuiltIns() {
  try {
    const raw = window.localStorage.getItem(HIDDEN_BUILTINS_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) return [];
    return parsed.filter((v) => typeof v === 'string');
  } catch {
    // Corrupted JSON or localStorage disabled (e.g. privacy mode) — treat
    // as if nothing has been hidden so the user can still see all built-ins.
    return [];
  }
}

function writeHiddenBuiltIns(ids) {
  try {
    window.localStorage.setItem(HIDDEN_BUILTINS_KEY, JSON.stringify(ids));
  } catch {
    // Best-effort: if storage is full or unavailable the deletion will
    // only last the current session, which is the same as the old
    // behaviour. We intentionally don't throw — the user already saw
    // the row disappear in the UI.
  }
}

export function getHiddenBuiltInIds() {
  return readHiddenBuiltIns();
}

export function hideBuiltInTemplate(id) {
  const current = readHiddenBuiltIns();
  if (current.includes(id)) return current;
  const next = [...current, id];
  writeHiddenBuiltIns(next);
  return next;
}

export function unhideBuiltInTemplate(id) {
  const current = readHiddenBuiltIns();
  if (!current.includes(id)) return current;
  const next = current.filter((x) => x !== id);
  writeHiddenBuiltIns(next);
  return next;
}

export function clearHiddenBuiltInTemplates() {
  writeHiddenBuiltIns([]);
}

// Read a File (the <input type=file> selection) and return its bytes as a
// Uint8Array. Used by the thumbnail generator to load the DOCX into a
// temporary DocumentEditor for rendering.
export async function readFileBytes(file) {
  const buf = await file.arrayBuffer();
  return new Uint8Array(buf);
}

// Generate a unique id on the client that matches the server-side slug rule
// (see vite-plugin-studio-templates.js). Useful when the studio wants to
// know the id ahead of the round-trip (e.g. for optimistic UI).
export function makeTemplateId(name) {
  const slug = (name || 'template')
    .replace(/\.[^.]+$/, '')
    .replace(/[^A-Za-z0-9-_]+/g, '_')
    .replace(/^_+|_+$/g, '') || 'template';
  return `tpl-${slug}-${Date.now().toString(36)}`;
}

// POST a custom merge field to the dev middleware so it lands in:
//   - scope "common"  -> src/data/user-templates/common-merge-fields.json
//   - scope "template" -> the matching template's entry inside the single
//                        templates.json collection (appended to fieldKeys;
//                        if no entry exists yet, one is created from the
//                        template fields the client supplies here)
// A field is now just its `FieldName` — there is no separate label, group,
// sample, repeating-block flag, columns, or sampleRows metadata. The
// server stores fields purely as keys; downstream the chip just renders
// the key and the editor inserts the MERGEFIELD using the key.
// Returns the server's JSON payload { ok, scope, key, field, ... }.
export async function addCustomMergeField({
  scope,            // 'template' | 'common'
  templateId,       // required when scope === 'template'
  templateName,     // optional, used to bootstrap a new entry if missing
  templateType,     // optional, used to bootstrap a new entry if missing
  templateDescription, // optional
  key,              // the FieldName (camelCase identifier)
}) {
  const fd = new FormData();
  fd.append('scope', scope);
  if (templateId) fd.append('templateId', templateId);
  if (templateName) fd.append('templateName', templateName);
  if (templateType) fd.append('templateType', templateType);
  if (templateDescription) fd.append('templateDescription', templateDescription);
  fd.append('key', key);
  const res = await fetch(`${API}/mergefield`, { method: 'POST', body: fd }).catch((err) => {
    // Network-level failure (dev server down, CORS, etc.). The dev plugin
    // is only registered when Vite runs with `apply: 'serve'`, so this
    // path is normal in production builds.
    throw new Error(
      `Failed to reach the studio dev server (${API}/mergefield). ` +
      'Make sure `npm run dev` is running from the Client-side folder. ' +
      `Underlying error: ${err.message || err}`,
    );
  });
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err.error || `Add field failed (${res.status})`);
  }
  const json = await res.json();
  if (!json.ok) throw new Error(json.error || 'Add field failed');
  return json;
}

// GET the common (global) custom merge-field catalog persisted on disk.
// Returns {} when the file does not exist or in production builds.
export async function fetchCommonMergeFields() {
  try {
    const res = await fetch(`${API}/common-fields`);
    if (!res.ok) return {};
    const json = await res.json();
    return json.fields || {};
  } catch {
    return {};
  }
}
