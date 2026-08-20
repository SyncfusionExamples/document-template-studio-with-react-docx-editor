import { useEffect, useMemo, useRef, useState } from 'react';
import { ButtonComponent } from '@syncfusion/ej2-react-buttons';
import { MERGE_FIELDS } from '../data/sampleTemplates.js';
import { addCustomMergeField } from '../utils/studioStorage.js';

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
function listFields(fieldKeys, customMap, commonFieldsProp) {
  const out = [];
  const seen = new Set();
  const push = (k) => {
    if (seen.has(k)) return;
    seen.add(k);
    out.push({ key: k });
  };
  // 1. Template-scoped keys (built-in or template-custom) — preserve
  //    the original order from `fieldKeys`.
  for (const k of (fieldKeys || [])) {
    if (MERGE_FIELDS[k] || (customMap && customMap[k])) push(k);
  }
  // 2. Common (global) keys — appended after the template's own fields
  //    so they show up in every template's panel and persist across
  //    reloads (loaded from common-merge-fields.json at app startup).
  for (const k of Object.keys(commonFieldsProp || {})) {
    if (MERGE_FIELDS[k] || (customMap && customMap[k])) push(k);
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
}) {
  const fields = useMemo(
    () => (template
      ? listFields(template.fieldKeys, customFieldMap, commonFieldsProp)
      : []),
    [template, customFieldMap, commonFieldsProp],
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
          Click a field to insert it at the caret.
        </p>
      )}

      {template && (
        <div className="ts-fields-groups">
                {fields.map((f) => (
                  <li key={f.key}>
                    <ButtonComponent
                      cssClass="e-block ts-field-chip"
                      title={`Insert ${f.key}`}
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
          Need a new field? Add a custom one — saved into this template’s JSON, or globally for all templates.
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
              <h3 id="ts-add-field-title">Add Custom Merge Field</h3>
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
