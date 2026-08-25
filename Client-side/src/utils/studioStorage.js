// The .NET backend URL is sourced from `DOCUMENT_EDITOR_BASE_URL` defined
// in ../data/sampleTemplates.js so the host/port lives in one place. The
// previous Vite dev plugin endpoints (/studio-api/...) are gone: all
// metadata operations now go directly to the ASP.NET Core
// StudioController at /api/studio/...

import { DOCUMENT_EDITOR_BASE_URL } from '../data/sampleTemplates.js';

// ASP.NET Core DocumentEditor Import endpoints. `Import` accepts a
// multipart upload from the browser (upload flow); `ImportFileURL`
// accepts a JSON `{ fileUrl }` body and the .NET service pulls the
// .docx from its own wwwroot/Templates/ folder server-side.
const DOC_EDITOR_IMPORT_URL = `${DOCUMENT_EDITOR_BASE_URL}/api/DocumentEditor/Import`;
const DOC_EDITOR_IMPORT_FILE_URL = `${DOCUMENT_EDITOR_BASE_URL}/api/DocumentEditor/ImportFileURL`;
const DOC_EDITOR_SAVE_URL = `${DOCUMENT_EDITOR_BASE_URL}/api/DocumentEditor/Save`;
const DOC_EDITOR_MAILMERGE_URL = `${DOCUMENT_EDITOR_BASE_URL}/api/DocumentEditor/MailMerge`;

// Server-side catalog + common-fields endpoints. The catalog JSON file
// is maintained by the ASP.NET Core service in wwwroot/Data/ and the
// .docx files live next to it under wwwroot/Templates/.
const STUDIO_API = `${DOCUMENT_EDITOR_BASE_URL}/api/studio`;

// Import a .docx into SFDT by talking to the authoritative ASP.NET
// Core DocumentEditor controller. Two call shapes are accepted:
//   - { file }    (upload flow): browser File -> multipart -> Import
//   - { url }     (open-existing flow): absolute docxUrl ->
//     JSON { fileUrl } -> ImportFileURL
export async function fetchSfdtFromDocx({ file, url, name }) {
  if (file) {
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
    const sfdt = await res.text();
    if (!sfdt) throw new Error('Mail merge returned an empty document.');
    return sfdt;
  }
  let detail = '';
  try { detail = await res.text(); } catch { /* ignore */ }
  throw new Error(
    `Mail merge failed (${res.status})${detail ? `: ${detail.slice(0, 240)}` : ''}`,
  );
}

// Read a Blob (from documentEditor.saveAsBlob('Docx')) as a base64 Data
// URL string. Mirrors the FileReader.readAsDataURL pattern in the
// Syncfusion sample.
export function readBlobAsDataUrl(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(reader.error || new Error('FileReader failed'));
    reader.readAsDataURL(blob);
  });
}

// ---------------------------------------------------------------------------
// Server-side catalog + common-fields API.
//
// The catalog file is now maintained by the ASP.NET Core service in
// wwwroot/Data/templates.json (and common-merge-fields.json next to it).
// The client fetches it on startup, persists changes back via PUT, and
// delegates file uploads to POST /api/studio/upload so the .docx + the
// catalog entry are written in one server-side transaction.
// ---------------------------------------------------------------------------

// Fetch the full template catalog. Returns [] when the server is down
// or the file is missing.
export async function fetchTemplatesCatalog() {
  try {
    const res = await fetch(`${STUDIO_API}/catalog`);
    if (!res.ok) return [];
    const json = await res.json();
    if (!Array.isArray(json)) return [];
    return json;
  } catch {
    return [];
  }
}

// Persist the client's template collection back to disk. The server
// is authoritative — it writes wwwroot/Data/templates.json verbatim.
export async function saveTemplatesCatalog(catalog) {
  const res = await fetch(`${STUDIO_API}/catalog`, {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(catalog),
  });
  if (!res.ok) throw new Error(`Catalog save failed (${res.status})`);
  return res.json();
}

// Upload a new .docx template to the server. The server writes the
// file to wwwroot/Templates/<slug>.docx AND appends a new entry to
// wwwroot/Data/templates.json. Returns { entry, docxFileName }.
export async function uploadTemplate({ file, name, type, description }) {
  const fd = new FormData();
  fd.append('file', file, file.name || 'template.docx');
  if (name) fd.append('name', name);
  if (type) fd.append('type', type);
  if (description) fd.append('description', description);
  const res = await fetch(`${STUDIO_API}/upload`, {
    method: 'POST',
    body: fd,
  });
  if (!res.ok) {
    let detail = '';
    try { detail = await res.text(); } catch { /* ignore */ }
    throw new Error(
      `Upload failed (${res.status})${detail ? `: ${detail.slice(0, 240)}` : ''}`,
    );
  }
  return res.json();
}

// Ask the server to remove a template's catalog entry. The .docx file
// is intentionally left in place (orphan) — the catalog is the single
// source of truth for the UI.
export async function deleteTemplate(id) {
  const res = await fetch(`${STUDIO_API}/template/${encodeURIComponent(id)}`, {
    method: 'DELETE',
  });
  if (!res.ok && res.status !== 404) {
    let detail = '';
    try { detail = await res.text(); } catch { /* ignore */ }
    throw new Error(
      `Delete failed (${res.status})${detail ? `: ${detail.slice(0, 200)}` : ''}`,
    );
  }
  return res.json().catch(() => ({ ok: true }));
}

// ---------------------------------------------------------------------------
// Hidden built-in templates.
//
// Deleting a built-in template is a client-side hide: the .docx stays
// on the server and the entry stays in the catalog file. We record the
// hidden ids in window.localStorage so the UI hides them on next load.
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
    return [];
  }
}

function writeHiddenBuiltIns(ids) {
  try {
    window.localStorage.setItem(HIDDEN_BUILTINS_KEY, JSON.stringify(ids));
  } catch {
    // Best-effort.
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

// POST a custom merge field to the server's StudioController. The
// server writes to common-merge-fields.json (common scope) or updates
// the matching catalog entry's fieldKeys (template scope).
export async function addCustomMergeField({
  scope,
  templateId,
  templateName,
  templateType,
  templateDescription,
  key,
}) {
  const fd = new FormData();
  fd.append('scope', scope);
  if (templateId) fd.append('templateId', templateId);
  if (templateName) fd.append('templateName', templateName);
  if (templateType) fd.append('templateType', templateType);
  if (templateDescription) fd.append('templateDescription', templateDescription);
  fd.append('key', key);
  const res = await fetch(`${STUDIO_API}/mergefield`, {
    method: 'POST',
    body: fd,
  });
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err.error || `Add field failed (${res.status})`);
  }
  const json = await res.json();
  if (!json.ok) throw new Error(json.error || 'Add field failed');
  return json;
}

// GET the common (global) custom merge-field catalog from the server.
export async function fetchCommonMergeFields() {
  try {
    const res = await fetch(`${STUDIO_API}/common-fields`);
    if (!res.ok) return {};
    const json = await res.json();
    return json.fields || {};
  } catch {
    return {};
  }
}
