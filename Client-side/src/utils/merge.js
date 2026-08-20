// Merge helpers: render template body HTML with sample data substituted.
import { MERGE_FIELDS } from '../data/sampleTemplates.js';

// Build the {FieldName: value} map used for preview substitution.
export function buildSampleContext(fieldKeys) {
  const ctx = {};
  for (const k of fieldKeys) {
    const f = MERGE_FIELDS[k];
    if (!f) continue;
    ctx[k] = f.sample;
  }
  return ctx;
}

// Render a repeating table block from its merge-field definition.
// Returns HTML string for tbody rows, or the table for stamping on insert.
function renderRepeatRows(def) {
  return def.sampleRows
    .map((row) => `<tr>${row.map((c) => `<td class="num">${c}</td>`).join('')}</tr>`)
    .join('');
}

// Replace {{FieldName}} tokens. Repeat fields expand to sample rows.
export function renderTemplateBody(body, context) {
  let html = body;
  // Repeating tables first (so a token like {{LineItems}} becomes rows).
  Object.keys(MERGE_FIELDS)
    .filter((k) => MERGE_FIELDS[k].repeat)
    .forEach((k) => {
      const token = new RegExp(`{{${k}}}`, 'g');
      html = html.replace(token, renderRepeatRows(MERGE_FIELDS[k]));
    });
  // Scalar fields.
  Object.keys(context).forEach((k) => {
    const token = new RegExp(`{{${k}}}`, 'g');
    html = html.replace(token, context[k] ?? '');
  });
  return html;
}

// Wrap a scalar merge field in a styled span so it is visually editable.
export function fieldBadgeHtml(key) {
  const f = MERGE_FIELDS[key];
  const label = f ? f.label : key;
  return `<span class="merge-badge" data-field="${key}" contenteditable="false">${label}</span>`;
}
