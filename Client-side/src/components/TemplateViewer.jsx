import { useEffect, useMemo, useRef, useState } from 'react';
import {
  DocumentEditorContainerComponent,
  Inject,
  Toolbar,
  Ribbon
} from '@syncfusion/ej2-react-documenteditor';
import { ButtonComponent } from '@syncfusion/ej2-react-buttons';
import { DialogComponent } from '@syncfusion/ej2-react-popups';
import { MERGE_FIELDS, DOCUMENT_EDITOR_SERVICE_URL } from '../data/sampleTemplates.js';
import { fetchSfdtFromDocx, saveTemplateToServer, mailMergePreview, readBlobAsDataUrl } from '../utils/studioStorage.js';
import MergeFieldsPanel from './MergeFieldsPanel.jsx';

// TemplateViewer-local collection of merge fields that don't live in the
// static MERGE_FIELDS catalog. It holds two kinds of entries:
//   - template-scoped custom fields (key -> true), derived at runtime from
//     `template.fieldKeys` (any key not in MERGE_FIELDS is treated as a
//     custom field) plus any new fields added via the "Add Field" UI.
//   - common (global) custom fields (key -> true). The Common catalog is
//     owned by App (which loads it from the dev server on startup and
//     forwards it down as a prop) so that every TemplateViewer sees the
//     same source of truth and new common fields are immediately visible
//     across all templates (and persist across reloads — the server
//     writes them to src/data/common-merge-fields.json).
// Both are exposed to the MergeFieldsPanel as `customFieldMap` so newly
// added fields are immediately insertable in the editor.

// TemplateViewer: a Word-like editor built on Syncfusion's DocumentEditor.
// Opens every template directly in Edit mode (no separate View/Edit toggle).
// - "Save and Publish"  = serializes the live document to SFDT JSON and
//                          POSTs it to the backend's
//                          DocumentEditorController.Save endpoint
//                          (http://localhost:5212/api/documenteditor/Save)
//                          with a SaveParameter { Content, FileName, Format:
//                          'Docx' }. The backend overwrites <name>.docx in
//                          Server-sde/wwwroot/Templates/ — saving with the
//                          same FileName replaces the existing file. The
//                          Vite dev plugin is intentionally NOT used.
// - "Download"           = calls Syncfusion's
//                          documentEditor.save(<document name>, 'Docx') to
//                          export a .docx copy straight to the user's local
//                          Downloads folder (browser-side save-as).
// The DocumentEditorContainer ships with its own built-in toolbar
// (enableToolbar) so we render NO manual formatting menu buttons.
function TemplateViewer({
  template,
  onThumbnailUpdated,
  onBack,
  commonFields: commonFieldsProp = {},
  onCommonFieldAdded = () => {},
  onTemplateFieldKeyAdded = () => {},
  onRequestPublish = () => {},
  registerPublishExecutor = () => {},
  onPublished = () => {},
  existingCategories: _existingCategories = [],
}) {
  const editorRef = useRef(null);
  const saveDialogRef = useRef(null);
  // dirty = the document has unsaved edits. The Save button is disabled
  // until the user actually changes something in the editor, and re-enabled
  // the moment they do (via the DocumentEditor's contentChange event).
  const [dirty, setDirty] = useState(false);
  // While the template is being loaded into the editor (open/openBlank +
  // seed-line insertions), contentChange WILL fire for the programmatic
  // inserts — we must NOT treat those as user edits. This ref gates the
  // contentChange handler so the document only goes dirty on real edits.
  const isLoadingRef = useRef(false);

  // A ref that always holds the latest template so the `created` callback
  // (which fires once on mount, before effects) can access it without
  // stale-closure issues.
  const templateRef = useRef(template);
  templateRef.current = template;

  function loadTemplateIntoEditor(de, tpl) {
    // Programmatic inserts below will fire contentChange — gate the handler
    // so this initial load doesn't mark the document dirty.
    isLoadingRef.current = true;
    if (tpl.docxUrl) {
      // Fetch the .docx, convert it to SFDT via the same-origin dev proxy
      // (calling the Syncfusion Import service directly from the browser
      // is blocked by CORS), then open() the resulting SFDT string. This
      // bypasses the DocumentEditor's internal AJAX handling and gives us
      // full control over the conversion step. The document name (without
      // .docx) is forwarded as `FileName` so the backend's Import call has
      // it for the upcoming Save round-trip.
      const baseName = (tpl.name || '').replace(/\.docx$/i, '').trim();
      fetchSfdtFromDocx({ url: tpl.docxUrl, name: baseName })
        .then((sfdt) => {
          de.open(sfdt);
          // open() is synchronous from the editor's perspective — content
          // is loaded by the time the promise resolves, so we can safely
          // re-enable the contentChange handler.
          isLoadingRef.current = false;
        })
        .catch((err) => {
          // eslint-disable-next-line no-console
          console.error('DOCX import failed:', err);
          de.openBlank();
          isLoadingRef.current = false;
        });
    } else {
      de.openBlank();
      const lines = tpl.seedLines ?? [];
      lines.forEach((line, i) => {
        de.editor.insertText(line);
        if (i < lines.length - 1) de.editor.insertText('\n');
      });
      // All seed-line insertions are done — re-enable the handler.
      isLoadingRef.current = false;
    }
  }

  // Called by the DocumentEditorContainer's `created` event — fires once
  // when the inner DocumentEditor is fully initialised. Because the
  // container is keyed by `template.id`, this callback runs fresh for each
  // template and there is no risk of stale document content.
  const handleCreated = () => {
    const de = editorRef.current?.documentEditor;
    if (!de) return;
    const tpl = templateRef.current;
    if (!tpl) return;

    // A fresh document load should NOT mark the template dirty — only
    // subsequent user edits should. We reset dirty here (and again after
    // the async load completes) so the Save button starts disabled.
    setDirty(false);
    loadTemplateIntoEditor(de, tpl);
  };

  // Fired by the DocumentEditor whenever the document content changes
  // (typing, formatting, insertions, deletions). Marks the template as
  // having unsaved edits so the Save button becomes enabled. Ignored while
  // a template is being loaded programmatically (open/openBlank + seed
  // lines) so the initial load doesn't falsely mark it dirty.
  const handleContentChange = () => {
    if (isLoadingRef.current) return;
    setDirty(true);
  };

  // The editor is always in Edit mode. We still force isReadOnly /
  // restrictEditing off once per template load — Syncfusion's
  // DocumentEditorContainer occasionally leaves the inner document
  // read-only when a new SFDT is open()'d; this guarantees the user can
  // type immediately. (The View/Edit toggle was removed per request.)
  useEffect(() => {
    const inst = editorRef.current;
    if (!inst) return;
    const de = inst.documentEditor;
    de.isReadOnly = false;
    de.restrictEditing = false;
  }, [template]);

  // Reset the dirty flag whenever a different template is opened — a fresh
  // load (handleCreated) shouldn't inherit "unsaved edits" from the
  // previously-opened template. The container is keyed by template.id so
  // it remounts fresh; this effect keeps the parent's dirty flag in sync.
  useEffect(() => {
    setDirty(false);
  }, [template?.id]);

  // -------- Custom (user-added) merge fields --------
  // A custom field is now just its `FieldName` — the template no longer
  // carries a `customFields` map<key, def>; it carries the field name in
  // its `fieldKeys` array. We seed the customFieldMap here as
  // { <key>: true } for every `fieldKeys` entry that's NOT already in the
  // built-in MERGE_FIELDS catalog (so user-added template-scoped fields
  // are recognized by the panel + insertField lookups). Plus the
  // `commonFields` map (globally-scoped custom fields), which is owned by
  // App and forwarded down as a prop so the same source of truth is
  // shared across every TemplateViewer.
  const [customFieldMap, setCustomFieldMap] = useState({});
  // Reset session-added custom fields whenever the template changes.
  useEffect(() => {
    // Seed the customFieldMap from the template's fieldKeys: any key not
    // in the built-in MERGE_FIELDS catalog is a template-scoped custom
    // field, marked `true` here so the panel + insertField recognise it.
    const seed = {};
    const keys = (template && Array.isArray(template.fieldKeys)) ? template.fieldKeys : [];
    for (const k of keys) {
      if (!MERGE_FIELDS[k]) seed[k] = true;
    }
    setCustomFieldMap({ ...seed });
  }, [template?.id, template?.fieldKeys]);

  // Callback fired by MergeFieldsPanel when a new field was successfully
  // persisted. Updates the in-memory map so the chip appears immediately,
  // appends the key to the current template's fieldKeys (so the field
  // counts toward "Fields in template" in the status bar and is remembered
  // for the rest of the session), and refreshes App's common-fields
  // catalog when the scope is "common" so every template sees it.
  const handleCustomFieldAdded = (info) => {
    setCustomFieldMap((prev) => ({ ...prev, [info.key]: info.field }));
    if (info.scope === 'common') {
      // Forward to App so the global commonFields state (and every other
      // template that opens afterwards) picks up the new key. The
      // already-rendered panel sees it via the commonFieldsProp below.
      onCommonFieldAdded(info.key, info.field);
    }
    if (info.scope === 'template' && Array.isArray(info.fieldKeys) && template) {
      // Mirror the new fieldKeys onto the in-memory template object so
      // MergeFieldsPanel (which reads template.fieldKeys directly) sees
      // the freshly-added chip for the rest of the session, and tell App
      // so the templates.json catalog reflects the new fieldKeys too.
      template.fieldKeys = info.fieldKeys;
      onTemplateFieldKeyAdded(template.id, info.fieldKeys);
    }
  };

  // Combined custom-field lookup passed to the panel: per-template first,
  // then common, so per-template entries can override a common one.
  // The "common" half comes from App (so it's shared across templates and
  // survives reloads via the server-persisted common-merge-fields.json).
  const combinedCustomFields = useMemo(
    () => ({ ...commonFieldsProp, ...customFieldMap }),
    [commonFieldsProp, customFieldMap],
  );

  // Insert a merge field at the current caret using the editor's API.
  // Falls back to the custom-field map for user-added fields
  // (template + common). A field is now just its `FieldName` (key) —
  // there's no separate label to use, so the MERGEFIELD uses the key
  // directly as the field name. The editor is always in Edit mode now,
  // so we just focusIn + insert.
  const insertField = (key) => {
    const inst = editorRef.current;
    if (!inst) return;
    const f = MERGE_FIELDS[key] || customFieldMap[key] || commonFieldsProp[key];
    if (!f) return;
    let fieldName = key
        .replace(/\n/g, "")
        .replace(/\r/g, "")
        .replace(/\r\n/g, "");
    let fieldCode = "MERGEFIELD  " + fieldName + "  \\* MERGEFORMAT ";
    const de = inst.documentEditor;
    de.focusIn();
    de.editor.insertField(fieldCode, "«" + fieldName + "»" );
  };

  // Save the document via the backend's DocumentEditorController.Save endpoint.
  //   1. Serialize the editor's content to a SFDT JSON string
  //      (de.serialize() returns the Syncfusion in-memory representation).
  //   2. POST it to http://localhost:5212/api/documenteditor/Save with a
  //      SaveParameter body of { Content, FileName, Format: 'Docx' }.
  //   3. The backend writes/overwrites <name>.docx in
  //      Server-side/wwwroot/Templates/ via FileMode.OpenOrCreate, so a
  //      second save with the same FileName REPLACES the existing file
  //      (no "Save as new file" behaviour).
  //   4. Re-render the thumbnail from the live editor and propagate it up
  //      via onThumbnailUpdated so the dashboard reflects the saved state.
  // The Vite dev plugin's /studio-api/save is intentionally NOT used for
  // this flow — we talk to the authoritative backend directly.
  //
  // First-time publish gate: when the template has not been published
  // yet (no `docxUrl`, has `seedLines` — i.e. a freshly "+ New Template"
  // doc still in `openBlank()` state), the user must confirm the file
  // name and Category before the .docx lands on disk. We surface the
  // publish dialog in App.jsx via `onRequestPublish`, which collects
  // the name + category and then calls back into `publishExecuteRef`
  // to do the actual Save. The dialog's confirmPublish handler in
  // App.jsx updates `templates.json` afterwards (sets `docxUrl`, drops
  // `seedLines`) so a reload loads from wwwroot/Templates/.
  const [isSaving, setIsSaving] = useState(false);
  // Hold a "pending publish" target so the dialog's confirm step can
  // pick it up. While this is non-null, the dirty flag is cleared as
  // soon as the dialog confirms the publish (the user has explicitly
  // chosen to publish via the dialog, so the document is in sync with
  // what will land on disk).
  const publishExecuteRef = useRef(null);

  // Register the publish-execute function with App.jsx so the publish
  // dialog's confirm handler can call into the editor to do the Save.
  // The function returns { docxBaseName, thumbnailDataUri } which the
  // dialog uses to update templates.json.
  useEffect(() => {
    registerPublishExecutor(async ({ name }) => {
      const inst = editorRef.current;
      if (!inst) throw new Error('Editor is not ready.');
      const de = inst.documentEditor;
      // 1. Serialize.
      let sfdtContent = '';
      try {
        sfdtContent = de.serialize();
      } catch (err) {
        // eslint-disable-next-line no-console
        console.error('serialize SFDT failed:', err);
        throw new Error('Could not serialize the document. Please try again.');
      }
      if (!sfdtContent) {
        throw new Error('Document serialized to an empty payload.');
      }
      // 2. Build the slug. We use a timestamp suffix to guarantee
      //    uniqueness on first-publish (so a second "+ New Template"
      //    saved with the same name never overwrites the wrong file).
      //    Subsequent saves use the same slug (held in template.docxUrl)
      //    so the editor reopens the right file.
      const base = (name || template.name || 'template').replace(/\.[^.]+$/, '').trim();
      const slug = `${base.replace(/[^A-Za-z0-9-_]+/g, '_').replace(/^_+|_+$/g, '') || 'template'}-${Date.now().toString(36)}`;
      const saveResult = await saveTemplateToServer({
        sfdtContent,
        documentName: slug,
        format: 'Docx',
      });
      // eslint-disable-next-line no-console
      console.log(`[studio] "${template.name}" published via DocumentEditorController.Save -> ${saveResult.fileName}.docx`);
      // 3. Refresh thumbnail so the dashboard reflects the published
      //    content without waiting for App.jsx's effect to re-render.
      let thumbnailDataUri = '';
      try {
        const { generateThumbnailFromEditor } = await import('../utils/thumbnailGenerator.js');
        const out = await generateThumbnailFromEditor(de);
        thumbnailDataUri = out.thumbnailDataUri || '';
      } catch (err) {
        // eslint-disable-next-line no-console
        console.warn('thumbnail refresh failed:', err);
      }
      // 4. Clear the editor's dirty flag so the Save button disables
      //    again — the document is now in sync with disk.
      setDirty(false);
      // Notify App.jsx so it can show a confirmation toast/dialog
      // (and the templates state is updated outside this callback).
      onPublished({ docxBaseName: saveResult.fileName || slug, thumbnailDataUri });
      return { docxBaseName: saveResult.fileName || slug, thumbnailDataUri };
    });
    return () => {
      // Clear on unmount so a previous viewer's executor doesn't linger
      // in App.jsx's ref.
      registerPublishExecutor(null);
    };
  }, [template?.id, registerPublishExecutor, onPublished]);

  // The Save button entry point. When the template has already been
  // published (docxUrl present), we save directly to the same slug.
  // When the template has NOT been published yet (fresh "+ New
  // Template" with seedLines), we open the publish dialog so the user
  // can confirm/correct the name + category. App.jsx's publish dialog
  // calls back into `publishExecuteRef` to do the actual Save.
  async function handleSave() {
    const inst = editorRef.current;
    if (!inst) return;
    // First-time publish: open the dialog and let App.jsx drive the
    // actual save + catalog update.
    if (!template.docxUrl) {
      onRequestPublish(template);
      return;
    }
    setIsSaving(true);
    try {
      // 1. Serialize the document to SFDT JSON (the format the backend's
      //    Save endpoint expects in `Content`).
      let sfdtContent = '';
      try {
        sfdtContent = inst.documentEditor.serialize();
      } catch (err) {
        // eslint-disable-next-line no-console
        console.error('serialize SFDT failed:', err);
        throw new Error('Could not serialize the document. Please try again.');
      }
      if (!sfdtContent) {
        throw new Error('Document serialized to an empty payload.');
      }

      // 2. POST the SFDT + name + format to the backend's Save endpoint.
      //    FileName MUST be sent WITHOUT the .docx extension (Format pins
      //    the type as Docx so the backend never sees ambiguous exts).
      //    The FileName MUST match the slug that `docxUrl` already points
      //    at — otherwise the saved file lands at a different path from
      //    the one the editor loads on reopen, and edits appear lost.
      //    For upload-style docs `docxUrl` is `/Templates/<slug>.docx`
      //    where <slug> is the safeFileName(name) + timestamp suffix; we
      //    reuse that exact slug as the Save FileName so Save overwrites
      //    the same .docx that `loadTemplateIntoEditor` reads on reopen.
      const docxBaseName = (() => {
        if (template.docxUrl) {
          // `/Templates/Donation_Thank-You_Letter-mt13cjft.docx` -> `Donation_Thank-You_Letter-mt13cjft`
          const tail = template.docxUrl.split('/').pop() || '';
          return tail.replace(/\.docx$/i, '').trim();
        }
        // Fallback for templates without an on-disk slug (e.g. a freshly
        // "+ New Template" doc still on `openBlank()`): slug the template
        // name the same way the upload flow does so the first Save lands
        // at a deterministic, reopen-safe path.
        const base = (template.name || 'template').replace(/\.[^.]+$/, '').trim();
        return base.replace(/[^A-Za-z0-9-_]+/g, '_').replace(/^_+|_+$/g, '') || 'template';
      })();
      const saveResult = await saveTemplateToServer({
        sfdtContent,
        documentName: docxBaseName,
        format: 'Docx',
      });
      // eslint-disable-next-line no-console
      console.log(`[studio] "${template.name}" saved via DocumentEditorController.Save -> ${saveResult.fileName}.docx`);

      // 3. Re-render the thumbnail from the live editor so the dashboard
      //    card reflects the saved content. This is best-effort — if it
      //    fails we still consider the save successful (the .docx on the
      //    server is already up to date).
      let thumbnailDataUri = '';
      try {
        const { generateThumbnailFromEditor } = await import('../utils/thumbnailGenerator.js');
        const out = await generateThumbnailFromEditor(inst.documentEditor);
        thumbnailDataUri = out.thumbnailDataUri || '';
      } catch (err) {
        // eslint-disable-next-line no-console
        console.warn('thumbnail refresh failed:', err);
      }

      // 4. Propagate the refreshed thumbnail into the in-memory template
      //    so the dashboard re-renders with the new image without a reload.
      if (thumbnailDataUri && onThumbnailUpdated) {
        onThumbnailUpdated(template.id, thumbnailDataUri);
      }

      saveDialogRef.current?.show();
      // The document is now in sync with what's on disk — clear the dirty
      // flag so the Save button disables again until the next edit.
      setDirty(false);
    } catch (err) {
      // eslint-disable-next-line no-console
      console.error('Save failed:', err);
      alert(`Save failed: ${err.message || err}`);
    } finally {
      setIsSaving(false);
    }
  }

  // Export the live document as a .docx saved into the user's local
  // Downloads folder. Uses Syncfusion's built-in `documentEditor.save()`
  // API, which serialises the editor's content and triggers a browser
  // download of `<fileName>.docx`. This is purely client-side — it does
  // NOT touch the backend or `wwwroot/Templates/`, so it's safe to call
  // even when the user just wants a copy without publishing changes.
  const [isDownloading, setIsDownloading] = useState(false);

  // -------- Preview with Data (mail merge) --------
  // Reveals a native HTML modal dialog where the user BROWSES for a .json
  // file of merge data (no paste). On OK we:
  //   1. Read the selected File as text, JSON.parse + validate it.
  //   2. Export the live editor's doc to a .docx BLOB via
  //      `documentEditor.saveAsBlob('Docx')` (exactly like the Syncfusion
  //      online sample).
  //   3. Read the BLOB as a base64 Data URL via `FileReader.readAsDataURL`
  //      (so documentData is a "data:application/...;base64,..." string).
  //   4. POST `{ fileName, documentData, mailMergeData }` to the backend's
  //      `/api/DocumentEditor/MailMerge` endpoint (ExportData on the .NET
  //      side). fileName = the editor's current documentName + ".docx".
  //   5. On 200, the backend returns the merged document as SFDT JSON.
  //      We feed it back to `de.open(sfdt)` so the editor re-renders with
  //      the merge fields replaced by the user's real data (a true preview).
  // The dialog is always rendered to the DOM (CSS toggles visibility) to
  // match the Add-Field dialog pattern and avoid Syncfusion+React-19
  // Dialog lifecycle issues.
  const [previewOpen, setPreviewOpen] = useState(false);
  // previewFile   = the File the user picked via the file input.
  // previewParsed = the JSON.parse'd object from that file (validated on
  //                 pick so OK can run immediately without re-reading).
  // previewFileName = display string of the selected file name (shown
  //                   next to the Browse button).
  const [previewFile, setPreviewFile] = useState(null);
  const [previewParsed, setPreviewParsed] = useState(null);
  const [previewFileName, setPreviewFileName] = useState('');
  const [previewError, setPreviewError] = useState('');
  const [isMerging, setIsMerging] = useState(false);
  const previewFileInputRef = useRef(null);

  // Default example JSON shown in the dialog as a reference for the
  // format the .json file must follow (a "Organization" root group whose
  // array of objects contains one entry per merged record). The keys
  // (OrgName, OrgAddress, DonorName, DonorAddress, DonationAmount,
  // DonationDate) deliberately match the built-in MERGE_FIELDS so users
  // see a working example for the donation-letter templates shipped
  // with the studio.
  const PREVIEW_EXAMPLE_JSON = `{
  "Organization": [
    {
      "OrgName": "ABC Foundation",
      "OrgAddress": "123 Main Street, New York, NY 10001",
      "DonorName": "John Smith",
      "DonorAddress": "45 Oak Street, New York, NY 10002",
      "DonationAmount": "$1,500.00",
      "DonationDate": "August 15, 2026"
    }
  ]
}`;

  function openPreviewDialog() {
    // Start with no file selected — the user must Browse.
    setPreviewFile(null);
    setPreviewParsed(null);
    setPreviewFileName('');
    setPreviewError('');
    setPreviewOpen(true);
  }

  function closePreviewDialog() {
    if (isMerging) return; // don't allow closing mid-merge
    setPreviewOpen(false);
    setPreviewError('');
  }

  // Read the user-selected .json File, parse it, stash the parsed object
  // for OK, and surface validation errors inline. We only accept .json
  // and require the top-level object to have an array property (matching
  // the backend's GetJsonData helper).
  async function handlePreviewFileChange(e) {
    const file = e.target.files && e.target.files[0];
    if (!file) {
      setPreviewFile(null);
      setPreviewParsed(null);
      setPreviewFileName('');
      return;
    }
    const lowerName = (file.name || '').toLowerCase();
    if (!lowerName.endsWith('.json')) {
      setPreviewFile(null);
      setPreviewParsed(null);
      setPreviewFileName('');
      setPreviewError('Please choose a .json file.');
      // Allow re-picking the same file later by clearing the input value.
      if (previewFileInputRef.current) previewFileInputRef.current.value = '';
      return;
    }
    setPreviewError('');
    try {
      const text = await file.text();
      const parsed = text.trim() ? JSON.parse(text) : null;
      if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) {
        setPreviewFile(null);
        setPreviewParsed(null);
        setPreviewFileName('');
        setPreviewError('The .json file must contain a JSON object (not an array or scalar). See the example format.');
        if (previewFileInputRef.current) previewFileInputRef.current.value = '';
        return;
      }
      // Top-level object must have at least one array property — the
      // backend's GetJsonData expects the .First() value to be a List of
      // row objects (a "group"). Surface a clear error here instead of
      // letting the .NET side throw.
      const firstKey = Object.keys(parsed)[0];
      if (!firstKey || !Array.isArray(parsed[firstKey])) {
        setPreviewFile(null);
        setPreviewParsed(null);
        setPreviewFileName('');
        setPreviewError(
          'The JSON object must have at least one array property, e.g. { "Organization": [ { ... } ] }. See the example format.',
        );
        if (previewFileInputRef.current) previewFileInputRef.current.value = '';
        return;
      }
      setPreviewFile(file);
      setPreviewParsed(parsed);
      setPreviewFileName(file.name);
    } catch (err) {
      setPreviewFile(null);
      setPreviewParsed(null);
      setPreviewFileName('');
      setPreviewError(`Could not read JSON file: ${err.message}`);
      if (previewFileInputRef.current) previewFileInputRef.current.value = '';
    }
  }

  function handleBrowseClick() {
    // Trigger the hidden <input type="file"> open dialog.
    previewFileInputRef.current?.click();
  }

  async function handlePreviewWithData() {
    const inst = editorRef.current;
    if (!inst) return;
    const de = inst.documentEditor;
    // The file + parsed payload are validated at pick time, but guard
    // against the user clicking OK without any selection.
    if (!previewFile || !previewParsed) {
      setPreviewError('Please browse for a .json data file first.');
      return;
    }
    const parsed = previewParsed;

    setIsMerging(true);
    setPreviewError('');
    try {
      // 1. Export the live document to a .docx BLOB (Syncfusion API).
      const blob = await de.saveAsBlob('Docx');
      // 2. Read the BLOB as a base64 Data URL (FileReader.readAsDataURL).
      const base64DataUrl = await readBlobAsDataUrl(blob);
      // 3. fileName = the editor's current documentName + ".docx".
      //    documentName is the slug set when the template was loaded
      //    (e.g. "Donation_Thank-You_Letter-mt13cjft"); fall back to the
      //    template name if documentName is empty.
      const docName = (de.documentName || '').trim()
        || (template.docxUrl ? template.docxUrl.split('/').pop().replace(/\.docx$/i, '') : template.name || 'Document');
      // 4. POST to /api/DocumentEditor/MailMerge.
      const mergedSfdt = await mailMergePreview({
        fileName: `${docName}.docx`,
        documentData: base64DataUrl,
        mailMergeData: JSON.stringify(parsed),
      });
      // 5. Open the merged SFDT in the editor — the live preview.
      de.open(mergedSfdt);
      // The merge has rewritten the document, so it diverges from the
      // on-disk file — mark dirty so the user can Save & Publish if they
      // want to keep the merged version (otherwise they can close).
      setDirty(true);
      setPreviewOpen(false);
      setPreviewFile(null);
      setPreviewParsed(null);
      setPreviewFileName('');
      if (previewFileInputRef.current) previewFileInputRef.current.value = '';
    } catch (err) {
      // eslint-disable-next-line no-console
      console.error('Mail merge preview failed:', err);
      setPreviewError(err.message || String(err) || 'Mail merge failed.');
    } finally {
      setIsMerging(false);
    }
  }
  async function handleDownload() {
    const inst = editorRef.current;
    if (!inst) return;
    const de = inst.documentEditor;
    // Derive a clean download file name from the template's docxUrl slug
    // (or template.name as a fallback). `.docx` is appended by Syncfusion
    // because the FormatType passed below is 'Docx'.
    let baseName = 'Document';
    if (template.name) {
      baseName = template.name.trim() || baseName;
    } else {
      const nm = (template.name || 'template').replace(/\.[^.]+$/, '').trim();
      baseName = nm.replace(/[^A-Za-z0-9-_]+/g, '_').replace(/^_+|_+$/g, '') || 'Document';
    }
    setIsDownloading(true);
    try {
      // de.save(fileName, formatType) triggers a browser-side download
      // of <fileName>.<ext>. 'Docx' is the Syncfusion FormatType enum
      // value for Word's .docx, so the file lands as <baseName>.docx.
      de.save(baseName, 'Docx');
    } catch (err) {
      // eslint-disable-next-line no-console
      console.error('Download failed:', err);
      alert(`Download failed: ${err.message || err}`);
    } finally {
      // Syncfusion's save() is fire-and-forget; release the spinner
      // shortly afterwards so the button label returns to normal.
      setTimeout(() => setIsDownloading(false), 800);
    }
  }

  if (!template) return null;

  return (
    <section className="ts-viewer">
      <header className="ts-viewer-head">
        <div className="ts-viewer-title">
          <ButtonComponent
            cssClass="e-flat ts-back"
            onClick={onBack}
            aria-label="Back to dashboard"
          >‹ Dashboard</ButtonComponent>
          <h2>{template.name}</h2>
          <p>{template.description}</p>
        </div>
        <div className="ts-viewer-actions">
          <ButtonComponent
            iconCss="e-icons e-eye"
            cssClass="e-flat e-info ts-btn-preview"
            disabled={isMerging}
            title="Preview the template with Mail Merge JSON data."
            onClick={openPreviewDialog}
          >{isMerging ? 'Merging…' : 'Preview with Data'}</ButtonComponent>
          <ButtonComponent
            iconCss="e-icons e-save"
            cssClass="e-outline e-primary ts-btn-save"
            disabled={isSaving || !dirty}
            title="Save and publish the template."
            onClick={handleSave}
          >{isSaving ? 'Saving…' : 'Save and Publish'}</ButtonComponent>
          <ButtonComponent
            iconCss="e-icons e-download"
            cssClass="e-flat e-primary ts-btn-download"
            disabled={isDownloading}
            title="Download the document content."
            onClick={handleDownload}
          >{isDownloading ? 'Preparing…' : 'Download'}</ButtonComponent>
        </div>
      </header>

      {/* Body: a two-column grid that holds the document editor (left) and
          the Merge Fields side panel (right). The header above spans both
          columns (see .ts-viewer grid layout in App.css), giving the
          ts-header the visual width of the editor + panel combined. */}
      <div className="ts-viewer-body">
        <div className="ts-viewer-canvas">
          {/* Syncfusion DocumentEditorContainer — built-in Word-like toolbar.
              The toolbar is enabled once at mount and stays mounted for the
              component lifetime; flipping enableToolbar on prop changes
              causes Syncfusion to stack duplicate toolbars. The editor is
              always in Edit mode (View/Edit toggle was removed per request);
              isReadOnly/restrictEditing are forced off in the effect above.
              serviceUrl points at the public Syncfusion web service that
              converts .docx → SFDT on import. */}
          <DocumentEditorContainerComponent
            key={template.id}
            ref={editorRef}
            height="100%"
            width="100%"
            enableToolbar={true}
            toolbarMode="Ribbon"
            showPropertiesPane={false}
            serviceUrl={DOCUMENT_EDITOR_SERVICE_URL}
            created={handleCreated}
            contentChange={handleContentChange}
          >
            <Inject services={[Ribbon]} />
          </DocumentEditorContainerComponent>
        </div>

        {/* Merge Fields side panel — now embedded inside the viewer so the
            ts-header above spans the editor + this panel. Field clicks go
            straight to the local insertField() (no global ref wiring).
            `customFieldMap` merges template-scoped + common custom fields so
            user-added fields appear in the chip list immediately. */}
        <MergeFieldsPanel
          template={template}
          onInsertField={insertField}
          customFieldMap={combinedCustomFields}
          onCustomFieldAdded={handleCustomFieldAdded}
          commonFieldsProp={commonFieldsProp}
        />
      </div>

      {/* Syncfusion dialog for "Save and Publish" confirmation */}
      <DialogComponent
        ref={saveDialogRef}
        id={`save-dialog-${template.id}`}
        header="Template published"
        content={`"${template.name}" has been saved and published to wwwroot/Templates.`}
        showCloseIcon
        target=".ts-app"
        width="360px"
        animationSettings={{ effect: 'Zoom', duration: 200 }}
        visible={false}
      />

      {/* Preview-with-Data (Mail Merge) dialog — native HTML modal (matches
          the Add-Field dialog pattern in MergeFieldsPanel; avoids
          Syncfusion+React-19 DialogComponent lifecycle issues). The user
          pastes a JSON object whose first property is an array of row
          objects; on OK we POST { fileName, documentData (base64 Data URL),
          mailMergeData (JSON.stringify(userInput)) } to the backend's
          /api/DocumentEditor/MailMerge endpoint and open() the returned
          merged SFDT back in the editor. */}
      {previewOpen && (
        <div
          className="ts-preview-overlay"
          role="dialog"
          aria-modal="true"
          aria-labelledby="ts-preview-title"
          onClick={(e) => {
            if (e.target === e.currentTarget && !isMerging) closePreviewDialog();
          }}
        >
          <div className="ts-preview-dialog" onClick={(e) => e.stopPropagation()}>
            <header className="ts-preview-head">
              <h3 id="ts-preview-title">Preview with Data</h3>
              <button
                type="button"
                className="ts-preview-close"
                aria-label="Close"
                onClick={closePreviewDialog}
                disabled={isMerging}
              >×</button>
            </header>

            <div className="ts-preview-body">

              <div className="ts-preview-file-row">
                <input
                  ref={previewFileInputRef}
                  type="file"
                  accept=".json,application/json"
                  onChange={handlePreviewFileChange}
                  disabled={isMerging}
                  className="ts-preview-file-input"
                />
                <ButtonComponent
                  cssClass="e-outline e-primary ts-preview-browse"
                  iconCss="e-icons e-folder"
                  onClick={handleBrowseClick}
                  disabled={isMerging}
                >Browse…</ButtonComponent>
                <span className="ts-preview-file-name">
                  {previewFileName || 'No file chosen'}
                </span>
              </div>

              <details className="ts-preview-example">
                <summary>Show example JSON format</summary>
                <pre>{PREVIEW_EXAMPLE_JSON}</pre>
              </details>

              {previewError && <p className="ts-preview-error">{previewError}</p>}
            </div>

            <footer className="ts-preview-actions">
              <ButtonComponent
                cssClass="e-flat"
                onClick={closePreviewDialog}
                disabled={isMerging}
              >Cancel</ButtonComponent>
              <ButtonComponent
                cssClass="e-primary"
                onClick={handlePreviewWithData}
                disabled={isMerging || !previewFile || !previewParsed}
                iconCss={isMerging ? 'e-icons e-refresh' : 'e-icons e-check'}
              >{isMerging ? 'Merging…' : 'OK'}</ButtonComponent>
            </footer>
          </div>
        </div>
      )}
    </section>
  );
}

export default TemplateViewer;
