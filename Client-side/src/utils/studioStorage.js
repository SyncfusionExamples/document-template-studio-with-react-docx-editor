import { DOCUMENT_EDITOR_BASE_URL } from '../data/sampleTemplates.js';

const API = `${DOCUMENT_EDITOR_BASE_URL}`;

const DOC_EDITOR_IMPORT_URL = `${DOCUMENT_EDITOR_BASE_URL}/api/DocumentEditor/Import`;
const DOC_EDITOR_IMPORT_FILE_URL = `${DOCUMENT_EDITOR_BASE_URL}/api/DocumentEditor/ImportFileURL`;

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

export function readBlobAsDataUrl(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(reader.error || new Error('FileReader failed'));
    reader.readAsDataURL(blob);
  });
}

export async function addCommonMergeField(key) {
  const response = await fetch(
    `${API}/api/TemplateStudio/merge-fields/common`,
    {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({ key }),
    },
  );

  if (!response.ok) {
    const error = await response.json().catch(() => ({}));
    throw new Error(
      error.message || `Failed to add common merge field (${response.status})`,
    );
  }

  return response.json();
}

export async function addTemplateMergeField(templateId, key) {
  const response = await fetch(
    `${API}/api/TemplateStudio/templates/${encodeURIComponent(templateId)}/merge-fields`,
    {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({ key }),
    },
  );

  if (!response.ok) {
    const error = await response.json().catch(() => ({}));
    throw new Error(
      error.message || `Failed to add template merge field (${response.status})`,
    );
  }

  return response.json();
}

export async function fetchCommonMergeFields() {
  const response = await fetch(
    `${API}/api/TemplateStudio/merge-fields/common`,
  );

  if (!response.ok) {
    throw new Error(
      `Failed to load common merge fields (${response.status})`,
    );
  }

  const keys = await response.json();

  return Object.fromEntries(
    (Array.isArray(keys) ? keys : []).map((key) => [key, true]),
  );
}

export async function createTemplate(template) {
  const response = await fetch(
    `${API}/api/TemplateStudio/templates`,
    {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(template),
    },
  );

  if (!response.ok) {
    const error = await response.json().catch(() => ({}));

    throw new Error(
      error.message ||
        `Failed to create template (${response.status})`,
    );
  }

  return response.json();
}

export async function fetchTemplates() {
  const response = await fetch(
    `${API}/api/TemplateStudio/templates`,
  );

  if (!response.ok) {
    throw new Error(
      `Failed to load templates (${response.status})`,
    );
  }

  return response.json();
}

export async function fetchTemplate(templateId) {
  const response = await fetch(
    `${API}/api/TemplateStudio/templates/${encodeURIComponent(templateId)}`,
  );

  if (!response.ok) {
    throw new Error(
      `Failed to load template (${response.status})`,
    );
  }

  return response.json();
}

export async function updateTemplate(templateId, data) {
  const response = await fetch(
    `${API}/api/TemplateStudio/templates/${encodeURIComponent(templateId)}`,
    {
      method: 'PUT',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(data),
    },
  );

  if (!response.ok) {
    const error = await response.json().catch(() => ({}));
    throw new Error(
      error.message || `Failed to update template (${response.status})`,
    );
  }

  return response.json();
}

export async function deleteTemplate(templateId) {
  const response = await fetch(
    `${API}/api/TemplateStudio/templates/${encodeURIComponent(templateId)}`,
    {
      method: 'DELETE',
    },
  );

  if (!response.ok) {
    const error = await response.json().catch(() => ({}));
    throw new Error(
      error.message || `Failed to delete template (${response.status})`,
    );
  }
}

