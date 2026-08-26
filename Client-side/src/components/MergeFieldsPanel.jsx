import { useEffect, useMemo, useRef, useState } from 'react';
import { ButtonComponent } from '@syncfusion/ej2-react-buttons';
import { MERGE_FIELDS } from '../data/sampleTemplates.js';
import { addCustomMergeField } from '../utils/studioStorage.js';

// Single source of truth for the MIME type we put on dataTransfer so the
// editor's drop handler can unambiguously tell a merge-field chip apart
// from a plain text drag (or a tab/header reorder). Both this and the
// callback below are read by TemplateViewer when parsing the dropped
// payload.
export const MERGE_FIELD_MIME = 'application/x-ts-mergefield';
export const MERGE_FIELD_PAYLOAD_MIME = 'application/json';

// Build a custom drag image that mirrors the chip layout (so the user
// sees a small floating « FieldName » ghost following the cursor).
// Returns an opaque wrapper element composed in JS — kept off-DOM
// (zero reflow cost) until dragend clears it. We use the same
// .ts-field-chip-label styling that the real chip uses so the ghost
// looks identical at a glance.
function buildDragGhost(key) {
  const ghost = document.createElement('div');
  ghost.className = 'ts-drag-ghost';
  // Same chrome as .ts-field-chip + the »« display field text the
  // editor will render after insertion. Positioned off-screen so the
  // ghost never flashes into the layout during setDragImage.
  ghost.style.position = 'fixed';
  ghost.style.top = '-9999px';
  ghost.style.left = '-9999px';
  ghost.style.zIndex = '2147483647';
  const label = document.createElement('span');
  label.className = 'ts-drag-ghost-label';
  // Show the display form (matching the editor's MERGEFIELD result
  // text) instead of the raw field name to make it obvious what the
  // drop will insert.
  label.textContent = `\u00ab ${key} \u00bb`;
  ghost.appendChild(label);
  document.body.appendChild(ghost);
  return ghost;
}

// Build a flat list of field descriptors for the currently selected template.
// Each field is just its `FieldName`. `customMap` (optional) augments the
// base MERGE_FIELDS catalog with template-scoped + common custom fields
// added at runtime. Since fields are now just their names, custom fields
// are stored as `true` (recognized) just like the built-ins, so no extra
// metadata (label/group/sample/repeat) is needed downstream.
//
// IMPORTANT: the panel must show every recognized field, not just the
// ones that happen to be in `fieldKeys`. `fieldKeys` is the set of
// fields the current template "owns" (used by Preview / Mail Merge),
// but the user can also pick a common (global) field and insert it
// into ANY template. Without the union below, a common field would
// only appear in a template's panel after the user explicitly added it
// at "template" scope for that template. We union fieldKeys with the
// commonFieldsProp so every global field is available everywhere.
//
// The third input source is `documentMergeFields`: the live set of
// MERGEFIELDs the Server-side DocumentEditorController.ImportFileURL
// Web API reported for the currently-loaded .docx. These are
// insertion-only fields — the .docx references them but the template's
// own `fieldKeys` (or common-merge-fields catalog) haven't necessarily
// claimed them yet. The panel lists them so the user can click → insert
// without first having to add the field to the template's catalog. We
// deliberately do NOT write them back into `fieldKeys` — the catalog
// stays the single source of truth for what fields "belong" to a
// template, and the doc-only fields are surfaced purely for chip-list
// convenience. De-duplication across all three sources is handled by
// the same `seen` Set regardless of where the key comes from.
function listFields(fieldKeys, customMap, commonFieldsProp, documentMergeFields) {
  const out = [];
  const seen = new Set();
  const push = (k) => {
    if (seen.has(k)) return;
    seen.add(k);
    out.push({ key: k });
  };
  // `template.fieldKeys` is the source of truth for what fields belong to
  // this template — anything in there MUST show in the chip list, even if
  // the field is template-scoped custom and not yet in the in-memory
  // `customMap` (e.g. fields the server already persisted on a previous
  // session, or fields added via the API in this session before
  // setCustomFieldMap commits). The previous filter
  // (`MERGE_FIELDS[k] || customMap[k]`) silently dropped exactly those
  // fields, which is the user-visible symptom of the bug we're fixing.
  for (const k of (fieldKeys || [])) {
    push(k);
  }
  // Common (global) keys — appended after the template's own fields so
  // they show up in every template's panel and persist across reloads
  // (loaded from common-merge-fields.json at app startup).
  for (const k of Object.keys(commonFieldsProp || {})) {
    // Avoid re-adding a key already in fieldKeys (template-scoped takes
    // precedence in the UI), and skip unknown common-field keys.
    push(k);
  }
  // Doc-only merge fields reported by ImportFileURL — these are
  // MERGEFIELDs that exist in the loaded .docx but aren't yet in
  // `fieldKeys` or commonFields. They're shown for convenience only:
  // the user can click to insert (the editor already knows the field
  // name), but they aren't promoted into the template catalog here.
  // We append them last so the template/common keys always sort first.
  for (const k of (documentMergeFields || [])) {
    if (!k) continue;
    push(k);
  }
  return out;
}

// Right panel: lists merge fields belonging to the currently selected template.
// Clicking a field inserts a MERGEFIELD into the editor at the caret.
// Includes a footer "Add Field" button that opens a dialog for creating
// custom FieldNames scoped to either the current template or globally.
// The dialog only collects two things:
//   - Save to (Template vs Common)
//   - Field Name (the new identifier; no separate label/sample/group/repeat)
function MergeFieldsPanel({
  template,
  onInsertField,
  customFieldMap,
  onCustomFieldAdded,
  commonFieldsProp = {},
  // Doc-only MERGEFIELDs returned by
  // DocumentEditorController.ImportFileURL. Array of field-name strings;
  // unioned into the chip list alongside `template.fieldKeys` and
  // `commonFieldsProp`. Defaults to [] so older callers (and the blank
  // "+ New Template" flow) still work.
  documentMergeFields = [],
}) {
  const fields = useMemo(
    () => (template
      ? listFields(template.fieldKeys, customFieldMap, commonFieldsProp, documentMergeFields)
      : []),
    [template, customFieldMap, commonFieldsProp, documentMergeFields],
  );

  // ----- Add-Field dialog state -----
  const [showAdd, setShowAdd] = useState(false);
  const [scope, setScope] = useState('template');   // 'template' | 'common'
  const [key, setKey] = useState('');
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState('');

  // Ref so the Field Name input focuses on open.
  const keyInputRef = useRef(null);

  // Reset form state when the dialog opens.
  useEffect(() => {
    if (!showAdd) return;
    setScope(template ? 'template' : 'common');
    setKey('');
    setError('');
    // Defer focus so the dialog is mounted.
    const t = setTimeout(() => keyInputRef.current?.focus(), 30);
    return () => clearTimeout(t);
  }, [showAdd, template]);

  // Validate the Field Name before allowing submit. Same camelCase rule
  // as before: letters/digits, leading letter.
  const validate = () => {
    if (!key.trim()) return 'Field Name is required.';
    if (!/^[A-Za-z][A-Za-z0-9]*$/.test(key.trim())) {
      return 'Field Name must start with a letter and contain only letters/digits.';
    }
    return '';
  };

  // ----- Drag-and-drop to the editor -----
  // The chip is both click-insertable (default behaviour on click) AND
  // HTML5-draggable into the DocumentEditor canvas on the left.
  // onDragStart: stash the field key under two MIME types so editors
  // (and our own drop target) can both recognise it:
  //   - MERGE_FIELD_MIME for the editor's specific drop handler
  //   - text/plain as a fallback for browsers that always expose that
  //     type but not custom MIME types
  //   - MERGE_FIELD_PAYLOAD_MIME for a JSON envelope that includes
  //     the source so the editor can refuse drops it did not initiate
  // We also create a custom ghost via setDragImage so the user sees a
  // "« FieldName »" pill following the cursor instead of the default
  // browser ghost. The element is removed on dragend to keep the DOM
  // clean.
  const handleChipDragStart = (e, k) => {
    if (!e || !e.dataTransfer) return;
    try {
      e.dataTransfer.setData(MERGE_FIELD_MIME, k);
      e.dataTransfer.setData('text/plain', k);
      e.dataTransfer.setData(
        MERGE_FIELD_PAYLOAD_MIME,
        JSON.stringify({ source: 'merge-fields-panel', key: k }),
      );
      e.dataTransfer.effectAllowed = 'copy';
      const ghost = buildDragGhost(k);
      // Offset so the cursor sits over the centre of the ghost pill.
      try {
        e.dataTransfer.setDragImage(ghost, ghost.offsetWidth / 2, ghost.offsetHeight / 2);
      } catch {
        // Some browsers throw if setDragImage is called outside an
        // active drag — safe to ignore; the default ghost takes over.
      }
      // Stash the ghost on the event so dragend (or unmount) can
      // remove it deterministically.
      e.currentTarget.__tsDragGhost = ghost;
    } catch {
      // dataTransfer not writeable for some reason — still allow the
      // native click behaviour.
    }
  };

  const handleChipDragEnd = (e) => {
    const ghost = e?.currentTarget?.__tsDragGhost;
    if (ghost && ghost.parentNode) {
      ghost.parentNode.removeChild(ghost);
    }
    if (e?.currentTarget) {
      e.currentTarget.__tsDragGhost = null;
    }
  };

  const handleSubmit = async (e) => {
    e?.preventDefault?.();
    const v = validate();
    if (v) { setError(v); return; }
    setError('');
    setSaving(true);
    try {
      
      const result = await addCustomMergeField({
        scope,
        templateId: scope === 'template' ? template?.id : undefined,
        // Pass through the template's identifying info so the server can
        // bootstrap a fresh .json metadata file when the template hasn't
        // been saved to disk yet (e.g. the in-memory blank "General
        // Correspondence" template).
        templateName: template?.name,
        templateType: template?.type,
        templateDescription: template?.description,
        key: key.trim()
      });
      // Bubble the new field up so the parent (TemplateViewer) can refresh
      // its in-memory catalogs and the editor insert path.
      if (onCustomFieldAdded) {
        onCustomFieldAdded({
          scope,
          templateId: scope === 'template' ? template?.id : null,
          key: result.key,
          field: result.field,
          fieldKeys: result.fieldKeys || null, // present for template scope
        });
      }
      setShowAdd(false);
    } catch (err) {
      setError(err.message || String(err));
    } finally {
      setSaving(false);
    }
  };

  return (
    <aside className="ts-fields-panel">
      <header className="ts-fields-head">
        <h3>Merge Fields</h3>
        {template ? (
          <p className="ts-fields-subtitle">
            For <strong>{template.name}</strong>
          </p>
        ) : (
          <p className="ts-fields-subtitle">Select a template to view fields</p>
        )}
      </header>

      {!template && (
        <p className="ts-empty">
          Merge fields will appear here when a template is opened.
        </p>
      )}

      {template && (
        <p className="ts-fields-hint">
          Click a field to insert at the caret, or drag it into the document.
        </p>
      )}

      {template && (
        <div className="ts-fields-groups">
                {fields.map((f) => (
                  // The <li> is the draggable handle — placing
                  // draggable on the <li> (not on the inner Syncfusion
                  // ButtonComponent) avoids the React-19 re-render
                  // wiping out the drag image in the middle of a drag.
                  <li
                    key={f.key}
                    className="ts-field-chip-li"
                    draggable
                    onDragStart={(e) => handleChipDragStart(e, f.key)}
                    onDragEnd={handleChipDragEnd}
                  >
                    <ButtonComponent
                      cssClass="e-block ts-field-chip"
                      title={`Insert ${f.key} — drag to drop into document`}
                      onClick={() => onInsertField(f.key)}
                    >
                      <span className="ts-field-chip-label">{f.key}</span>
                    </ButtonComponent>
                  </li>
                ))}
                {fields.length === 0 && (
                  <p className="ts-empty">
                    No merge fields yet. Use the <strong>Add Field</strong> button below to add one.
                  </p>
                )}
        </div>
      )}

      {/* Footer: Add Field button + hint text. Sticks to the bottom of the
          scrollable panel so the action is always reachable. */}
      <footer className="ts-fields-foot">
        <p className="ts-fields-hint-foot">
          Need a new field? Add a custom one — saved into this template or globally for all templates.
        </p>
        <ButtonComponent
          cssClass="e-primary e-block ts-btn-add-field"
          iconCss="e-icons e-plus"
          onClick={() => setShowAdd(true)}
          disabled={!template}
        >Add Field</ButtonComponent>
      </footer>

      {/* Add-Field dialog (built with native HTML, not the Syncfusion
          DialogComponent, to avoid the Syncfusion+React 19 lifecycle issues
          documented elsewhere in the app. The modal is always rendered; CSS
          controls visibility so the form fields are mounted only when open. */}
      {showAdd && (
        <div
          className="ts-add-field-overlay"
          role="dialog"
          aria-modal="true"
          aria-labelledby="ts-add-field-title"
          onClick={(e) => {
            // Close when clicking the backdrop (but not the dialog body).
            if (e.target === e.currentTarget && !saving) setShowAdd(false);
          }}
        >
          <div className="ts-add-field-dialog" onClick={(e) => e.stopPropagation()}>
            <header className="ts-add-field-head">
              <h3 id="ts-add-field-title">Add Merge Field</h3>
              <button
                type="button"
                className="ts-add-field-close"
                aria-label="Close"
                onClick={() => !saving && setShowAdd(false)}
              >×</button>
            </header>

            <form className="ts-add-field-form" onSubmit={handleSubmit}>
              {/* Scope radio: This Template vs Common (all templates). */}
              <fieldset className="ts-add-field-scope">
                <legend>Save to</legend>
                <label className="ts-radio">
                  <input
                    type="radio"
                    name="ts-add-scope"
                    value="template"
                    checked={scope === 'template'}
                    onChange={() => setScope('template')}
                    disabled={!template || saving}
                  />
                  <span>
                    <strong>This template</strong>
                    <em>{template ? ` — ${template.name}` : ' (no template open)'}</em>
                  </span>
                </label>
                <label className="ts-radio">
                  <input
                    type="radio"
                    name="ts-add-scope"
                    value="common"
                    checked={scope === 'common'}
                    onChange={() => setScope('common')}
                    disabled={saving}
                  />
                  <span>
                    <strong>Common (all templates)</strong>
                    <em> — added to the global field catalog</em>
                  </span>
                </label>
              </fieldset>

              <label className="ts-add-field-label" htmlFor="ts-af-key">
                Field Name
                <input
                  ref={keyInputRef}
                  id="ts-af-key"
                  type="text"
                  className="ts-input"
                  placeholder="e.g. DonorEmail"
                  value={key}
                  onChange={(e) => setKey(e.target.value)}
                  disabled={saving}
                  pattern="^[A-Za-z][A-Za-z0-9]*$"
                  required
                />
                <small className="ts-hint">CamelCase, letters/digits only, leading letter.</small>
              </label>

              {error && <p className="ts-add-field-error">{error}</p>}

              <footer className="ts-add-field-actions">
                <ButtonComponent
                  cssClass="e-flat"
                  onClick={() => setShowAdd(false)}
                  disabled={saving}
                >Cancel</ButtonComponent>
                <ButtonComponent
                  cssClass="e-primary"
                  type="submit"
                  disabled={saving}
                  iconCss={saving ? 'e-icons e-refresh' : 'e-icons e-check'}
                >{saving ? 'Saving…' : 'Add Field'}</ButtonComponent>
              </footer>
            </form>
          </div>
        </div>
      )}
    </aside>
  );
}

export default MergeFieldsPanel;
