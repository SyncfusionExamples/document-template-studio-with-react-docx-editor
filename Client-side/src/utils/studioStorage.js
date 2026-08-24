
// The .NET backend URL is sourced from `DOCUMENT_EDITOR_BASE_URL` defined
// in ../data/sampleTemplates.js so the host/port lives in one place. The
// dev-only /studio-api endpoints fall through cleanly in a production
// build because the plugin is registered with `apply: 'serve'`.

import { DOCUMENT_EDITOR_BASE_URL } from '../data/sampleTemplates.js';

const API = '/studio-api';

// The Vite dev plugin is still used for development-only catalog and
// custom-field operations, but DOCX Import now goes directly from the
// browser to the authoritative ASP.NET Core DocumentEditor controller.
// No upload filesystem write is performed by Vite.
//
// ASP.NET Core DocumentEditor Import endpoints. `Import` accepts a
// multipart upload from the browser (upload flow); `ImportFileURL`
// accepts a JSON `{ fileUrl }` body and the .NET service pulls the
// .docx from its own wwwroot/Templates/ folder server-side, which is
// the only path that works when the React app and the .NET service
// are on different machines (no same-origin proxy is available).
const DOC_EDITOR_IMPORT_URL = `${DOCUMENT_EDITOR_BASE_URL}/api/DocumentEditor/Import`;
const DOC_EDITOR_IMPORT_FILE_URL = `${DOCUMENT_EDITOR_BASE_URL}/api/DocumentEditor/ImportFileURL`;

// Import a .docx into SFDT by talking to the authoritative ASP.NET
// Core DocumentEditor controller. Two call shapes are accepted:
//   - { file }    (upload flow):
//       The browser `File` is posted directly as multipart/form-data
//       to `POST /api/DocumentEditor/Import`. No FileReader is
//       involved and no Vite proxy is required.
//   - { url }     (open-existing flow):
//       The absolute `${DOCUMENT_EDITOR_BASE_URL}/Templates/<slug>.docx`
//       URL is forwarded as `{ fileUrl }` in a JSON body to
//       `POST /api/DocumentEditor/ImportFileURL`. The .NET service
//       downloads the .docx from its own static-file path and returns
//       the SFDT directly. This collapses what would otherwise be a
//       browser fetch → server POST into a single same-origin round
//       trip and is the only path that works when the React app runs
//       on a different machine than ASP.NET Core.
export async function fetchSfdtFromDocx({ file, url, name }) {
  if (file) {
    // Upload flow: post the browser File as multipart/form-data.
    const baseName = (name || file.name || 'template').replace(/\.docx$/i, '').trim();
    const fd = new FormData();
    fd.append('docx', file, file.name || 'template.docx');
    if (baseName) fd.append('FileName', baseName);

    const res = await fetch(DOC_EDITOR_IMPORT_URL, {
      method: 'POST',
      body: fd,
    });

    if (!res.ok) {
      let detail = '';
      try { detail = await res.text(); } catch { /* ignore */ }
      throw new Error(
        `Import failed (${res.status})${detail ? `: ${detail.slice(0, 240)}` : ''}`,
      );
    }
    const responseText = await res.text();
    if (!responseText) throw new Error('Import returned an empty document.');
    return responseText;
  }

  if (url) {
    // Open-existing flow: let the .NET service resolve the URL from
    // its own wwwroot/Templates/ folder. The .NET controller's
    // `FileUrlInfo` body only carries `fileUrl`, so the request shape
    // is just `{ fileUrl }` — no `FileName` is needed (the Save round
    // trip derives its own slug from `template.docxUrl` at write time).
    //
    // Cache-busting is appended to the URL we hand to the server: the
    // .NET service uses WebClient.DownloadData underneath, which is
    // process-local and not cached, but the query string also dodges
    // any front-of-house proxy cache between the React origin and the
    // .NET service (a real concern in a multi-machine deployment).
    //const cacheBustUrl = url + (url.includes('?') ? '&' : '?') + 't=' + Date.now().toString(36);

    const res = await fetch(DOC_EDITOR_IMPORT_FILE_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json;charset=UTF-8' },
      body: JSON.stringify({ fileUrl: url }),
    });

    if (!res.ok) {
      let detail = '';
      try { detail = await res.text(); } catch { /* ignore */ }
      throw new Error(
        `ImportFileURL failed (${res.status})${detail ? `: ${detail.slice(0, 240)}` : ''}`,
      );
    }
    const responseText = await res.text();
    if (!responseText) throw new Error('ImportFileURL returned an empty document.');
    return responseText;
  }

  throw new Error('fetchSfdtFromDocx: file or url required');
}


const DOC_EDITOR_SAVE_URL = `${DOCUMENT_EDITOR_BASE_URL}/api/DocumentEditor/Save`;

function stripDocxExtension(name) {
  if (!name) return 'Document';
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

const DOC_EDITOR_MAILMERGE_URL = `${DOCUMENT_EDITOR_BASE_URL}/api/DocumentEditor/MailMerge`;

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
// Clearing localStorage restores the built-ins.
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
