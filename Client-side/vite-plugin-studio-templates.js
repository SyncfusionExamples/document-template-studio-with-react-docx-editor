// Vite dev-only middleware plugin.
//
// Endpoints (all dev-only, registered with apply:'serve'):
//   PUT    /studio-api/catalog        - persist the client's templates.json collection
//   GET    /studio-api/common-fields  - read the persisted common merge-field catalog
//   POST   /studio-api/mergefield     - persist a custom merge field (template / common)

import { promises as fs } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// The single client-side catalog file. The client reads this on startup
// (one place to look up every template's metadata) and writes back to it
// via PUT /studio-api/catalog whenever the collection changes.
const CATALOG_FILE = path.resolve(__dirname, 'src', 'data', 'templates.json');

// Common (global) custom merge-field catalog, also living next to the
// templates catalog. Created on first use.
const COMMON_FIELDS_FILE = path.resolve(__dirname, 'src', 'data', 'common-merge-fields.json');

const API_PREFIX = '/studio-api';

// --- helpers -------------------------------------------------------------

// Read a multipart/form-data body from a Node IncomingMessage and pull out
// the named file fields. Vite's dev server gives us the raw stream; we parse
// just enough to recover file buffers + filenames.
async function readMultipart(req) {
  const boundaryMatch = /boundary=(?:"([^"]+)"|([^;]+))/i.exec(
    req.headers['content-type'] || '',
  );
  if (!boundaryMatch) return { fields: {}, files: {} };
  const boundaryBuf = Buffer.from('--' + (boundaryMatch[1] || boundaryMatch[2]));
  // Some clients (e.g. .NET HttpClient) put the boundary in quotes inside
  // the Content-Type header; strip the surrounding quotes if present.
  const boundaryText = (boundaryMatch[1] || boundaryMatch[2] || '').replace(/^"|"$/g, '');

  const chunks = [];
  for await (const chunk of req) chunks.push(chunk);
  const buf = Buffer.concat(chunks);

  const fields = {};
  const files = {};

  // Walk the buffer once, splitting on each occurrence of `--<boundary>`.
  // Each part sits between two boundaries: a header block (terminated by
  // the first \r\n\r\n) followed by the body. We find the header/body
  // boundary by searching for \r\n\r\n WITHIN just the header section —
  // not the whole segment — so a binary body that happens to contain
  // \r\n\r\n (e.g. the inside of a .docx ZIP) can't trick the parser
  // into thinking the headers run on. The header section ends at the
  // first \r\n\r\n that is followed by either a CRLF or directly by
  // non-header bytes. In practice, headers never contain \r\n\r\n, so
  // searching within the first 8 KiB is plenty.
  let start = 0;
  let nextPart;
  // eslint-disable-next-line no-cond-assign
  while ((nextPart = buf.indexOf(boundaryBuf, start)) !== -1) {
    // Skip the boundary itself + the CRLF that separates it from the
    // part's headers.
    let segStart = nextPart + boundaryBuf.length;
    if (
      segStart + 1 < buf.length &&
      buf[segStart] === 0x0d && buf[segStart + 1] === 0x0a
    ) {
      segStart += 2;
    }

    // Find the NEXT boundary — this bounds the part. We allow the closing
    // "--<boundary>--" form to be matched too (its prefix is "--<boundary>"
    // so indexOf will find it).
    const segEnd = buf.indexOf(boundaryBuf, segStart);
    if (segEnd === -1) break; // malformed trailing data

    // The raw segment between the two boundaries (may include a trailing
    // \r\n right before the next boundary; strip it).
    let s = buf.subarray(segStart, segEnd);
    if (s.length >= 2 && s[s.length - 2] === 0x0d && s[s.length - 1] === 0x0a) {
      s = s.subarray(0, s.length - 2);
    }

    // If this is the closing marker segment (s starts with "--" — the
    // "--<boundary>--" tail), skip it.
    if (s.length >= 2 && s[0] === 0x2d && s[1] === 0x2d) {
      start = segEnd;
      continue;
    }
    // Empty segment => skip.
    if (s.length === 0) {
      start = segEnd;
      continue;
    }

    // Find the header/body separator. Headers are always short ASCII
    // (Content-Disposition, Content-Type, ...) so we can safely scan
    // for \r\n\r\n within the first 8 KiB without risk of finding a
    // \r\n\r\n inside a binary body.
    const headerScanLimit = Math.min(s.length, 8 * 1024);
    let headerEnd = -1;
    for (let i = 0; i < headerScanLimit - 3; i++) {
      if (
        s[i] === 0x0d && s[i + 1] === 0x0a &&
        s[i + 2] === 0x0d && s[i + 3] === 0x0a
      ) {
        headerEnd = i;
        break;
      }
    }
    if (headerEnd === -1) {
      start = segEnd;
      continue;
    }
    const headerStr = s.subarray(0, headerEnd).toString('utf8');
    const valueBuf = s.subarray(headerEnd + 4);

    const nameMatch = /name=(?:"([^"]+)"|([^\s;]+))/.exec(headerStr);
    if (!nameMatch) {
      start = segEnd;
      continue;
    }
    const fieldName = nameMatch[1] || nameMatch[2];
    const fileMatch = /filename=(?:"([^"]*)"|([^\s;]+))/.exec(headerStr);

    if (fileMatch) {
      files[fieldName] = {
        filename: fileMatch[1] || fileMatch[2] || '',
        data: valueBuf,
      };
    } else {
      fields[fieldName] = valueBuf.toString('utf8');
    }

    start = segEnd;
  }
  return { fields, files };
}

function send(res, status, body) {
  res.statusCode = status;
  res.setHeader('Content-Type', 'application/json');
  res.end(JSON.stringify(body));
}

// --- route handlers ------------------------------------------------------

// PUT /studio-api/catalog
// Body: JSON array — the full template collection the client currently has.
//
// Persists the client's single templates.json file at
// Client-side/src/data/templates.json so the latest collection survives
// Vite dev server restarts. The client always posts the *entire* array
// (no per-entry diffing needed).
async function handleCatalog(req, res) {
  const chunks = [];
  for await (const chunk of req) chunks.push(chunk);
  const body = Buffer.concat(chunks).toString('utf8');
  let parsed;
  try {
    parsed = JSON.parse(body);
  } catch (err) {
    return send(res, 400, { error: `Invalid JSON: ${err.message}` });
  }
  if (!Array.isArray(parsed)) {
    return send(res, 400, { error: 'Catalog body must be a JSON array' });
  }
  await fs.mkdir(path.dirname(CATALOG_FILE), { recursive: true });
  await fs.writeFile(CATALOG_FILE, JSON.stringify(parsed, null, 2));
  return send(res, 200, { ok: true, count: parsed.length });
}

async function handleAddMergeField(req, res) {
  const { fields } = await readMultipart(req);
  const scope = fields.scope || '';
  const key = (fields.key || '').trim();

  if (!key) {
    return send(res, 400, { error: 'key is required' });
  }
  if (scope !== 'template' && scope !== 'common') {
    return send(res, 400, { error: 'scope must be "template" or "common"' });
  }
  if (scope === 'template' && !fields.templateId) {
    return send(res, 400, { error: 'templateId required for template scope' });
  }

  // A field is just its name. Store `true` so downstream lookups
  // (MERGE_FIELDS[key] || customMap[key]) treat the key as recognised.
  const field = true;

  // Make sure the parent directories exist for the catalog files.
  await fs.mkdir(path.dirname(CATALOG_FILE), { recursive: true });
  await fs.mkdir(path.dirname(COMMON_FIELDS_FILE), { recursive: true });

  if (scope === 'common') {
    // Persist into a shared catalog file. Schema: { fields: { Key: true } }
    // Always normalize to `true` so the on-disk shape stays consistent
    // (an earlier version of the app wrote { label, group, sample }
    // objects here; if those entries survived into the current file we
    // simply overwrite them with `true` on the next add, so the catalog
    // converges on the flat shape that the rest of the app expects).
    let catalog = { fields: {} };
    try {
      catalog = JSON.parse(await fs.readFile(COMMON_FIELDS_FILE, 'utf8'));
      if (!catalog.fields || typeof catalog.fields !== 'object') catalog.fields = {};
    } catch { /* missing -> start fresh */ }
    // Drop any non-truthy garbage entries (e.g. a stray null that a
    // previous write might have left behind) so the lookup in
    // MergeFieldsPanel.listFields stays clean.
    for (const k of Object.keys(catalog.fields)) {
      if (!catalog.fields[k] || typeof catalog.fields[k] !== 'object') continue;
      // Old shape { label, group, sample } — drop the metadata and keep
      // just the recognition flag.
      catalog.fields[k] = true;
    }
    catalog.fields[key] = field;
    catalog.updatedAt = new Date().toISOString();
    await fs.writeFile(COMMON_FIELDS_FILE, JSON.stringify(catalog, null, 2));
    return send(res, 200, { ok: true, scope, key, field });
  }

  // scope === "template": update the matching entry inside the single
  // client-side templates.json collection. The .docx stays untouched on
  // the server — we only mutate the catalog entry's fieldKeys.
  const templateId = fields.templateId;
  let catalog = [];
  try {
    catalog = JSON.parse(await fs.readFile(CATALOG_FILE, 'utf8'));
    if (!Array.isArray(catalog)) catalog = [];
  } catch { /* missing -> empty */ }

  let entry = catalog.find((t) => t && t.id === templateId);
  let createdNew = false;
  if (!entry) {
    // Bootstrap a fresh entry from the info the client sent so the field
    // can land even for templates not yet present in the catalog.
    createdNew = true;
    entry = {
      id: templateId,
      name: (fields.templateName || '').trim() || templateId,
      type: (fields.templateType || '').trim() || 'General',
      description: (fields.templateDescription || '').trim() || 'Blank letter template.',
      fieldKeys: [],
      createdAt: new Date().toISOString(),
    };
    catalog.push(entry);
  }

  // Append the new FieldName to the entry's fieldKeys (de-duplicating).
  const fieldKeys = Array.isArray(entry.fieldKeys) ? entry.fieldKeys : [];
  if (!fieldKeys.includes(key)) fieldKeys.push(key);
  entry.fieldKeys = fieldKeys;
  entry.updatedAt = new Date().toISOString();

  await fs.writeFile(CATALOG_FILE, JSON.stringify(catalog, null, 2));
  return send(res, 200, {
    ok: true,
    scope,
    key,
    field,
    fieldKeys,
    entry,
    catalog,
    created: createdNew,
  });
}

// GET /studio-api/common-fields
// Returns the persisted common (global) merge-field catalog. The file may
// not exist yet — in that case an empty { fields: {} } payload is returned.
async function handleCommonFields(_req, res) {
  let catalog = { fields: {} };
  try {
    const raw = await fs.readFile(COMMON_FIELDS_FILE, 'utf8');
    const parsed = JSON.parse(raw);
    if (parsed && typeof parsed === 'object') catalog = parsed;
    if (!catalog.fields || typeof catalog.fields !== 'object') catalog.fields = {};
  } catch { /* missing -> default */ }
  return send(res, 200, catalog);
}

function tryParse(s, fallback) {
  try { return JSON.parse(s); } catch { return fallback; }
}

// --- plugin factory ------------------------------------------------------

export function studioTemplateFiles() {
  return {
    name: 'studio-template-files',
    apply: 'serve', // dev-only
    configureServer(server) {
      server.middlewares.use(async (req, res, next) => {
        if (!req.url || !req.url.startsWith(API_PREFIX)) return next();
        const url = new URL(req.url, 'http://localhost');
        const route = url.pathname.slice(API_PREFIX.length);

        try {
          if (req.method === 'PUT' && route === '/catalog') return await handleCatalog(req, res);
          if (req.method === 'GET' && route === '/common-fields') return await handleCommonFields(req, res);
          if (req.method === 'POST' && route === '/mergefield') return await handleAddMergeField(req, res);
        } catch (err) {
          // eslint-disable-next-line no-console
          console.error('[studio-api] error:', err);
          return send(res, 500, { error: String(err && err.message || err) });
        }
        return next();
      });
    },
  };
}
