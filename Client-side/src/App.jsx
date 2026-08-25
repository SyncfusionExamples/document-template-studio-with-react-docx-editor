import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { DialogComponent } from '@syncfusion/ej2-react-popups';
import { TextBoxComponent } from '@syncfusion/ej2-react-inputs';
import { ComboBoxComponent } from '@syncfusion/ej2-react-dropdowns';
import { ButtonComponent } from '@syncfusion/ej2-react-buttons';
import Sidebar from './components/Sidebar.jsx';
import Dashboard from './components/Dashboard.jsx';
import TemplateViewer from './components/TemplateViewer.jsx';
import {
  NEW_TEMPLATE_FIELD_KEYS,
  DOCUMENT_EDITOR_BASE_URL
} from './data/sampleTemplates.js';
// The template catalog is now maintained entirely on the server
// (Server-side/wwwroot/Data/templates.json) by the ASP.NET Core
// StudioController. The client fetches it on startup, sends the full
// collection back on each save, and delegates uploads to a single
// server-side endpoint that writes the .docx + the catalog entry.
import {
  fetchTemplatesCatalog,
  saveTemplatesCatalog,
  getHiddenBuiltInIds,
  hideBuiltInTemplate,
  fetchCommonMergeFields,
  uploadTemplate,
} from './utils/studioStorage.js';
import './App.css';

// Build an absolute `${DOCUMENT_EDITOR_BASE_URL}/Templates/<slug>.docx`
// URL from a stored `docxUrl`. Stored values come in two shapes:
//   1. absolute `http(s)://host[:port]/Templates/<slug>.docx` (current
//      multi-machine shape), or
//   2. relative `/Templates/<slug>.docx` (legacy shape, kept readable
//      by `templates.json` shipped from older builds).
// Both are folded into an absolute URL so `fetchSfdtFromDocx({ url })`
// hits the ASP.NET Core service directly — no Vite proxy is involved.
function absoluteDocxUrl(docxUrl) {
  if (!docxUrl) return '';
  if (/^[a-z][a-z0-9+.-]*:\/\//i.test(docxUrl)) return docxUrl; // already absolute
  if (docxUrl.startsWith('/')) {
    return `${DOCUMENT_EDITOR_BASE_URL}${docxUrl}`;
  }
  // Bare slug fallback (defensive — no caller should write this form).
  return `${DOCUMENT_EDITOR_BASE_URL}/Templates/${docxUrl}`;
}

// Helper to generate a unique id for new (blank) templates. The server
// assigns its own slug on upload, so for client-only blank templates
// (no .docx yet) this id is purely a React-state identifier.
function makeId(prefix = 'tpl') {
  return `${prefix}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 7)}`;
}

// The ids of the entries the server shipped with on first fetch. We use
// this to detect "is this id from the seeded catalog?" so deleting a
// built-in is treated as a client-side hide (no .docx unlink) while
// deleting a user-uploaded template removes its entry from the
// server-side catalog file.
let BUILTIN_IDS = new Set();

function App() {
  // Catalog is now fetched from the server on first mount. We start with
  // an empty list and let `loadCatalogFromServer` populate it. Built-in
  // templates the user has previously removed are filtered out via
  // localStorage so the deletion persists across reloads / service
  // restarts. The fetched .docx URLs are server-relative
  // (`/Templates/<slug>.docx`) and are resolved to absolute URLs by
  // `absoluteDocxUrl` for the Syncfusion DocumentEditor import path.
  const [templates, setTemplates] = useState([]);
  const [catalogLoaded, setCatalogLoaded] = useState(false);
  // null  = dashboard view; an id = that template opened in the editor.
  const [selectedId, setSelectedId] = useState(null);
  const deleteDialogRef = useRef(null);
  const [deleteTarget, setDeleteTarget] = useState(null);

  // Initial load: fetch the server-side catalog and the common merge
  // fields in parallel. The same common-fields source of truth is
  // forwarded to every TemplateViewer so adding a common field on one
  // template is immediately visible in every other template's panel and
  // persists across reloads (the server writes to
  // wwwroot/Data/common-merge-fields.json).
  const [commonFields, setCommonFields] = useState({});
  useEffect(() => {
    let cancelled = false;
    Promise.all([fetchTemplatesCatalog(), fetchCommonMergeFields()])
      .then(([catalog, fields]) => {
        if (cancelled) return;
        const hidden = new Set(getHiddenBuiltInIds());
        // Capture the built-in ids the first time we load the catalog
        // so subsequent deletes can be classified as hide-vs-remove.
        BUILTIN_IDS = new Set(catalog.map((t) => t.id));
        // Normalize legacy absolute URLs to the absolute form so the
        // Syncfusion DocumentEditor import path always talks to the .NET
        // service directly.
        const normalized = catalog
          .filter((t) => !hidden.has(t.id))
          .map((t) => (
            t.docxUrl && !/^[a-z][a-z0-9+.-]*:\/\//i.test(t.docxUrl)
              ? { ...t, docxUrl: absoluteDocxUrl(t.docxUrl) }
              : t
          ));
        setTemplates(normalized);
        setCommonFields(fields || {});
        setCatalogLoaded(true);
      })
      .catch((err) => {
        // eslint-disable-next-line no-console
        console.warn('[catalog] initial load failed:', err);
        setCatalogLoaded(true);
      });
    return () => { cancelled = true; };
  }, []);

  const selected = templates.find((t) => t.id === selectedId) ?? null;

  // Open a template from the dashboard (or sidebar) in the editor.
  const handleOpen = useCallback((id) => setSelectedId(id), []);
  const handleBack = useCallback(() => setSelectedId(null), []);

  const publishDialogRef = useRef(null);
  const [isPublishDialogOpen, setIsPublishDialogOpen] = useState(false);
  const [publishTarget, setPublishTarget] = useState(null); // template being published
  const [publishName, setPublishName] = useState('');
  const [publishCategory, setPublishCategory] = useState('General');
  const [publishPurpose, setPublishPurpose] = useState('');
  const [isPublishing, setIsPublishing] = useState(false);
  const [publishError, setPublishError] = useState('');
  // TemplateViewer fills this in when the user opens the editor so the
  // publish dialog's "OK" handler can call back into it to run the real
  // Save. Returns a promise that resolves to { docxBaseName, thumbnailDataUri }.
  const publishExecuteRef = useRef(null);
  // Confirmation dialog shown after a successful first-time publish.
  // (The editor's own saveDialogRef is local to the editor instance, so
  // we drive a sibling dialog here. We keep this small surface so the
  // user gets the same "Template published" feedback as for subsequent
  // saves.)
  const [showPublishedDialog, setShowPublishedDialog] = useState(false);
  const [publishedDialogName, setPublishedDialogName] = useState('');
  // When the user closes the confirmation, scroll back to the dashboard.
  const [lastPublishedId, setLastPublishedId] = useState(null);
  // Whether the confirmation was triggered from first-time publish or a
  // subsequent Save (overwrite path). Both flows reuse the same
  // Stay / Back-to-dashboard dialog so the UI is identical; this flag
  // only tweaks the dialog's header + body copy.
  const [publishedDialogMode, setPublishedDialogMode] = useState('publish'); // 'publish' | 'save'

  // Open the Save-success confirmation (overwrite path) with the same
  // Stay / Back-to-dashboard dialog used by first-time publish. Called
  // by TemplateViewer's handleSave once the .docx has been overwritten
  // on disk. Naming here is "publishedDialog*" for historical reasons
  // (it was added for the publish flow); the dialog is generic and
  // serves both.
  const handleSaved = useCallback(({ templateId, name } = {}) => {
    setPublishedDialogName(name || (templateId && (templates.find((t) => t.id === templateId) || {}).name) || '');
    setLastPublishedId(templateId || null);
    setPublishedDialogMode('save');
    setShowPublishedDialog(true);
  }, [templates]);

  // Open the publish dialog for a given template. Called by
  // TemplateViewer when the user clicks "Save and Publish" on a template
  // that has not been published yet.
  const handleRequestPublish = useCallback((tpl) => {
    if (!tpl) return;
    setPublishTarget(tpl);
    setPublishName(tpl.name || '');
    setPublishCategory(tpl.type || 'General');
    setPublishPurpose(tpl.description || '');
    setPublishError('');
    setIsPublishDialogOpen(true);
    publishDialogRef.current?.show();
  }, []);

  const cancelPublish = useCallback(() => {
    publishDialogRef.current?.hide();
    setIsPublishDialogOpen(false);
    setPublishTarget(null);
    setPublishError('');
  }, []);

  const confirmPublish = useCallback(async () => {
    if (!publishTarget) return;
    if (!publishName.trim()) {
      setPublishError('Template name is required.');
      return;
    }
    setPublishError('');
    setIsPublishing(true);
    try {
      // Ask TemplateViewer to do the actual Save (it owns the editor
      // instance + the saveTemplateToServer call). The execute function
      // is expected to return the slug it saved to, plus an optional
      // refreshed thumbnail.
      if (typeof publishExecuteRef.current !== 'function') {
        throw new Error('Editor is not ready — please try again.');
      }
      const { docxBaseName, thumbnailDataUri } = await publishExecuteRef.current({
        name: publishName.trim(),
        // The category is informational; the saved .docx is named after
        // `name`, not after the category. The backend Save endpoint
        // derives the on-disk filename from FileName only.
        category: publishCategory.trim() || 'General',
      });
      const slug = docxBaseName;
      // Rework the in-memory template: stamp docxUrl so the dashboard
      // and the next "Save and Publish" both read it from disk, drop
      // seedLines (the file is now the source of truth), refresh
      // name/type/description, and bump updatedAt.
      setTemplates((prev) => prev.map((t) => {
        if (t.id !== publishTarget.id) return t;
        const next = {
          ...t,
          name: publishName.trim(),
          type: (publishCategory || '').trim() || 'General',
          description: (publishPurpose || '').trim() || t.description,
          docxUrl: absoluteDocxUrl(`/Templates/${slug}.docx`),
          updatedAt: new Date().toISOString(),
        };
        // The blank template is now backed by an on-disk file — drop
        // the in-memory seed so the editor reloads from the file.
        delete next.seedLines;
        if (thumbnailDataUri) next.thumbnailUrl = thumbnailDataUri;
        return next;
      }));
      // The existing useEffect([templates]) hook writes the catalog to
      // disk via PUT /studio-api/catalog, so reload picks up docxUrl.
      cancelPublish();
    } catch (err) {
      // eslint-disable-next-line no-console
      console.error('Publish failed:', err);
      setPublishError(err?.message || String(err));
    } finally {
      setIsPublishing(false);
    }
  }, [publishTarget, publishName, publishCategory, publishPurpose, cancelPublish]);

  // ----- Persist the template collection back to the server -----
  // The server is authoritative: wwwroot/Data/templates.json is the
  // single source of truth and we PUT the full collection whenever it
  // changes. We skip the very first render so we don't immediately
  // rewrite the file we just loaded, and we wait for the initial
  // catalog fetch to complete (catalogLoaded) so transient empty-state
  // renders don't wipe the server file.
  //
  // The full collection (including the `thumbnailUrl` data URI on each
  // entry) is persisted as-is so the on-disk catalog is the single home
  // of thumbnail images too — no separate .png files, no re-generation
  // on every load. Once a thumbnail has been rendered for a template
  // it is cached inside the catalog entry for subsequent reloads.
  const firstRenderRef = useRef(true);
  useEffect(() => {
    if (firstRenderRef.current) {
      firstRenderRef.current = false;
      return;
    }
    if (!catalogLoaded) return;

    saveTemplatesCatalog(templates).catch((err) => {
      // eslint-disable-next-line no-console
      console.warn('[catalog] saveTemplatesCatalog failed:', err);
    });
  }, [templates, catalogLoaded]);

  // "+ New Template" creates a blank template doc + its merge-field set,
  // places it in the sidebar list, and opens it immediately so the user
  // can design the letter and Save (persist to the data folder). The
  // template body starts empty — the editor opens with `openBlank()`
  // (see TemplateViewer.loadTemplateIntoEditor) and only the merge-field
  // catalog is pre-populated.
  const handleAdd = useCallback(() => {
    const id = makeId();
    const tpl = {
      id,
      name: 'New Letter Template',
      type: 'General',
      description: 'Blank letter template — edit to customize.',
      fieldKeys: [...NEW_TEMPLATE_FIELD_KEYS],
    };
    setTemplates((prev) => [...prev, tpl]);
    setSelectedId(id);
  }, []);

  // ---- Upload flow ----
  // The dialog collects the .docx + metadata. On confirm we POST it
  // to the server's /api/studio/upload endpoint, which writes the
  // .docx to wwwroot/Templates/<slug>.docx AND adds a new entry to
  // wwwroot/Data/templates.json in a single transaction. The returned
  // entry is added to React state and the editor is opened on it. The
  // browser File is held in `pendingUpload` until the editor has
  // rendered and captured a thumbnail.
  const [isUploading, setIsUploading] = useState(false);
  const uploadDialogRef = useRef(null);
  const dialogFileInputRef = useRef(null);
  // { file, templateId } — the browser File remains here until the
  // initial Import + DocumentEditor Save has completed successfully.
  const [pendingUpload, setPendingUpload] = useState(null);
  const [uploadName, setUploadName] = useState('');
  const [uploadCategory, setUploadCategory] = useState('General');
  const [uploadPurpose, setUploadPurpose] = useState('');
  // Whether the Upload Template dialog is currently open. The Syncfusion
  // dialog stays mounted in the DOM even when hidden, which means its child
  // ComboBox/TextBox components are always rendered — and the ComboBox in
  // particular crashes on prop re-comparison while hidden (a Syncfusion +
  // React 19 interaction). We therefore only mount the form's children
  // while the dialog is open.
  const [isUploadDialogOpen, setIsUploadDialogOpen] = useState(false);

  // Existing categories are collected from the current template collection
  // (built-in + uploaded) so the user can reuse them — plus the 4 built-in
  // defaults. De-duplicated and sorted. The Category field is a free-text
  // input backed by a <datalist>, so the user can either pick an existing
  // category OR type a brand-new one.
  const existingCategories = useMemo(() => {
    const set = new Set(['General', 'Invoice', 'ThankYou', 'TaxReceipt']);
    for (const t of templates) set.add(t.type);
    return Array.from(set).sort();
  }, [templates]);

  // Open the Upload Template dialog (the file is picked from INSIDE the
  // dialog via its Browse button).
  const handleUploadClick = useCallback(() => {
    setPendingUpload(null);
    setUploadName('');
    setUploadCategory('General');
    setUploadPurpose('');
    setIsUploadDialogOpen(true);
    uploadDialogRef.current?.show();
  }, []);

  // Triggered by the in-dialog Browse button — opens the OS file picker.
  const handleBrowseClick = useCallback(() => {
    dialogFileInputRef.current?.click();
  }, []);

  // File picked from the in-dialog Browse button — stash it and pre-fill
  // the Name with the file's base name (the dialog is already open).
  const handleUploadFile = useCallback((e) => {
    const file = e.target.files?.[0];
    e.target.value = '';
    if (!file) return;
    if (!/\.docx$/i.test(file.name)) {
      alert('Please select a .docx file.');
      return;
    }
    setUploadName(file.name.replace(/\.docx$/i, ''));
    // templateId is filled in by `confirmUpload` once the dialog
    // confirms — setting it here to `null` so reviewers reading the
    // state shape can tell the two apart.
    setPendingUpload({ file, templateId: null });
  }, []);

  // The user confirmed the dialog. We POST the file + metadata to the
  // server's /api/studio/upload endpoint in a single transaction. The
  // server writes the .docx to wwwroot/Templates/<slug>.docx AND appends
  // a new entry to wwwroot/Data/templates.json. We then open the
  // returned entry in the editor and capture a thumbnail from the live
  // editor for the dashboard card.
  const confirmUpload = useCallback(async () => {
    const file = pendingUpload?.file;
    if (!file) {
      alert('Please choose a .docx file to upload.');
      return;
    }

    const name = (uploadName || '').trim() || file.name.replace(/\.docx$/i, '');
    const category = (uploadCategory || '').trim() || 'General';
    const purposeWasProvided = Boolean((uploadPurpose || '').trim());
    const description = purposeWasProvided
      ? uploadPurpose.trim()
      : 'Uploaded .docx template';

    uploadDialogRef.current?.hide();
    setIsUploadDialogOpen(false);
    setIsUploading(true);

    try {
      const result = await uploadTemplate({
        file,
        name,
        type: category,
        description,
      });
      const serverEntry = result?.entry;
      if (!serverEntry || !serverEntry.id) {
        throw new Error('Server returned an invalid upload response.');
      }
      // Normalize the docxUrl to the absolute form so the editor's
      // ImportFileURL flow always targets the .NET service directly.
      const entry = {
        ...serverEntry,
        docxUrl: absoluteDocxUrl(serverEntry.docxUrl || `/Templates/${result.docxFileName}`),
      };
      setTemplates((prev) => {
        const existingIds = new Set(prev.map((t) => t.id));
        if (existingIds.has(entry.id)) {
          // Server already knows about this id (extremely unlikely);
          // overwrite the local copy to keep it in sync.
          return prev.map((t) => (t.id === entry.id ? entry : t));
        }
        return [...prev, entry];
      });
      setPendingUpload({ file, templateId: entry.id });
      setSelectedId(entry.id);
    } catch (err) {
      // eslint-disable-next-line no-console
      console.error('[upload] failed:', err);
      alert(`Upload failed: ${err.message || err}`);
      setIsUploading(false);
      setPendingUpload(null);
    }
  }, [pendingUpload, uploadName, uploadCategory, uploadPurpose]);

  // Successful callback fired by TemplateViewer once the editor has
  // finished opening the uploaded .docx and captured its thumbnail. At
  // this point the server already has the .docx and a catalog entry
  // (the upload endpoint handled that), so all we do here is attach
  // the captured thumbnail to the existing entry and clear the
  // in-flight browser File from React state.
  const handleInitialUploadSaved = useCallback(({
    templateId,
    thumbnailDataUri = '',
    pageCount = 0,
  } = {}) => {
    if (!templateId) {
      setIsUploading(false);
      setPendingUpload(null);
      return;
    }

    setTemplates((prev) => prev.map((t) => {
      if (t.id !== templateId) return t;
      const next = { ...t };
      if (pageCount > 0 && (!next.description || next.description.startsWith('Uploaded .docx'))) {
        next.description = `Uploaded .docx (${pageCount} page${pageCount > 1 ? 's' : ''})`;
      }
      if (thumbnailDataUri) next.thumbnailUrl = thumbnailDataUri;
      return next;
    }));

    setPendingUpload(null);
    setIsUploading(false);
  }, []);

  // Failure callback fired by TemplateViewer when the editor cannot
  // import or open the uploaded .docx. The server already has the file
  // and the catalog entry, so we keep the entry on disk but alert the
  // user; the in-memory editor view is reset to the dashboard.
  const handleInitialUploadFailed = useCallback((error) => {
    setSelectedId(null);
    setPendingUpload(null);
    setIsUploading(false);

    const message = error?.message || String(error || 'Unknown error');
    // eslint-disable-next-line no-console
    console.error('[upload] editor open failed:', error);
    alert(`Editor could not open the uploaded template: ${message}`);
  }, []);

  const cancelUpload = useCallback(() => {
    uploadDialogRef.current?.hide();
    setIsUploadDialogOpen(false);
    // While the upload is in flight the transient template is already in
    // React state, so leave `pendingUpload` alone — it will be cleared
    // by the success/failure callback. Cancel without a freeze means
    // the user simply dismissed the dialog before confirming.
    if (!isUploading) setPendingUpload(null);
  }, [isUploading]);

  // ---- Save flow ----
  // The editor's "Save Template" button talks directly to the backend's
  // DocumentEditorController.Save endpoint
  // (${DOCUMENT_EDITOR_BASE_URL}/api/DocumentEditor/Save) with a
  // SaveParameter body — see TemplateViewer.handleSave and
  // saveTemplateToServer() in studioStorage.js. The Vite dev plugin's
  // /studio-api/save is NOT used for this flow. App's only role after
  // save is to update the in-memory thumbnail (see handleThumbnailUpdated
  // below).

  // After a save, the viewer hands us a refreshed thumbnail data URI. We
  // convert it to a Blob URL so the dashboard <img> can render it without
  // needing a server round-trip / page reload.
  const handleThumbnailUpdated = useCallback((id, dataUri) => {
    setTemplates((prev) => prev.map((t) => {
      if (t.id !== id) return t;
      return { ...t, thumbnailUrl: dataUri };
    }));
  }, []);

  // ---- Delete flow (requirement 2) ----
  // Open the Syncfusion dialog asking the user to confirm deletion.
  const handleDelete = useCallback((id) => {
    const tpl = templates.find((t) => t.id === id);
    setDeleteTarget(tpl ?? null);
    deleteDialogRef.current?.show();
  }, [templates]);

  const confirmDelete = useCallback(async () => {
    if (!deleteTarget) return;
    const id = deleteTarget.id;
    const wasSelected = id === selectedId;
    // Delete semantics: built-in templates are hidden client-side
    // (the .docx stays on the server, the catalog entry stays on the
    // server — the next reload uses localStorage to filter them out).
    // User-uploaded templates are removed from the in-memory list and
    // the persistence effect PUTs the new (smaller) collection to the
    // server, so wwwroot/Data/templates.json no longer references
    // them. The .docx itself is left orphaned (the catalog is the
    // single source of truth for the UI).
    const isBuiltin = BUILTIN_IDS.has(id);
    if (isBuiltin) {
      hideBuiltInTemplate(id);
    }
    setTemplates((prev) => prev.filter((t) => t.id !== id));
    if (wasSelected) setSelectedId(null);
    setDeleteTarget(null);
    deleteDialogRef.current?.hide();
  }, [deleteTarget, selectedId]);

  return (
    <div className={selected ? 'ts-app' : 'ts-app is-dashboard'}>
      <Sidebar
        templates={templates}
        selectedId={selectedId}
        onSelect={handleOpen}
        onAdd={handleAdd}
        onUpload={handleUploadClick}
        isUploading={isUploading}
        onDelete={handleDelete}
      />

      <main className="ts-main">
        {selected ? (
          <TemplateViewer
            template={selected}
            initialUploadFile={
              pendingUpload?.templateId === selected.id ? pendingUpload.file : null
            }
            onInitialUploadSaved={handleInitialUploadSaved}
            onInitialUploadFailed={handleInitialUploadFailed}
            onThumbnailUpdated={handleThumbnailUpdated}
            onBack={handleBack}
            commonFields={commonFields}
            onCommonFieldAdded={(key, field) => {
              setCommonFields((prev) => ({ ...prev, [key]: field }));
            }}
            onTemplateFieldKeyAdded={(templateId, newFieldKeys) => {
              setTemplates((prev) => prev.map((t) =>
                t.id === templateId ? { ...t, fieldKeys: newFieldKeys } : t,
              ));
            }}
            onRequestPublish={handleRequestPublish}
            registerPublishExecutor={(fn) => { publishExecuteRef.current = fn; }}
            onPublished={() => {
              // The editor's publish executor has already cleared the
              // dirty flag and refreshed the thumbnail. Show the same
              // "Template published" dialog used by subsequent saves.
              // The actual templates-state update is handled by
              // confirmPublish so a reload picks up docxUrl.
              setPublishedDialogName(publishName.trim() || (selected && selected.name) || '');
              setLastPublishedId(selected ? selected.id : null);
              setPublishedDialogMode('publish');
              setShowPublishedDialog(true);
            }}
            onSaved={handleSaved}
            existingCategories={existingCategories}
          />
        ) : (
          <Dashboard
            templates={templates}
            onOpen={handleOpen}
            onAdd={handleAdd}
            onUpload={handleUploadClick}
            onDelete={handleDelete}
            isUploading={isUploading}
          />
        )}
      </main>

      {/* Hidden file input used by the in-dialog Browse button. Kept outside
          the dialog DOM so the file picker is always reachable. */}
      <input
        ref={dialogFileInputRef}
        type="file"
        accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        onChange={handleUploadFile}
        style={{ display: 'none' }}
      />

      {/* Upload-template details dialog: collects the .docx file (via an
          in-dialog Browse button), a Name, a Category (existing OR new), and
          a very short Purpose (shown at the bottom of the card) before the
          thumbnail is generated and the .docx is persisted. */}
      <DialogComponent
        ref={uploadDialogRef}
        id="ts-upload-dialog"
        header="Upload Template"
        showCloseIcon
        visible={false}
        target=".ts-app"
        width="480px"
        animationSettings={{ effect: 'Zoom', duration: 200 }}
        close={cancelUpload}
      >
        {isUploadDialogOpen && (
        <div className="ts-upload-form">
          {/* .docx file picker. Shows the chosen filename or a Browse button. */}
          <label className="ts-upload-label">Template file (.docx)</label>
          <div className="ts-upload-file-row">
            <span className="ts-upload-file-name" title={pendingUpload?.file?.name || ''}>
              {pendingUpload?.file ? pendingUpload.file.name : 'No file chosen'}
            </span>
            <ButtonComponent
              iconCss="e-icons e-folder-open"
              cssClass="e-outline ts-upload-browse"
              onClick={handleBrowseClick}
            >Browse</ButtonComponent>
          </div>
          <label className="ts-upload-label" htmlFor="ts-upload-name">Template name</label>
          <TextBoxComponent
            id="ts-upload-name"
            placeholder="e.g. Annual Appeal Letter"
            value={uploadName}
            input={(e) => setUploadName(e.value ?? '')}
            floatLabelType="Never"
          />
          <label className="ts-upload-label" htmlFor="ts-upload-category">Category</label>
          <ComboBoxComponent
            id="ts-upload-category"
            cssClass="ts-upload-combobox"
            dataSource={existingCategories}
            value={uploadCategory || 'General'}
            allowCustom={true}
            allowFiltering={true}
            placeholder="Select or type a new category"
            floatLabelType="Never"
            change={(e) => setUploadCategory(e.value ?? 'General')}
            input={(e) => setUploadCategory(e.value ?? '')}
          />
          <label className="ts-upload-label" htmlFor="ts-upload-purpose">Purpose</label>
          <TextBoxComponent
            id="ts-upload-purpose"
            placeholder="Very short description shown on the card"
            value={uploadPurpose}
            input={(e) => setUploadPurpose(e.value ?? '')}
            floatLabelType="Never"
          />
          {/* Action buttons rendered as custom children (not via the
              Syncfusion `buttons` prop) so React onClick handlers fire
              reliably and we control the alignment via CSS flexbox. */}
          <div className="ts-dialog-actions">
            <ButtonComponent
              cssClass="e-flat"
              onClick={cancelUpload}
            >Cancel</ButtonComponent>
            <ButtonComponent
              cssClass="e-primary"
              onClick={confirmUpload}
            >Upload</ButtonComponent>
          </div>
        </div>
        )}
      </DialogComponent>

      <DialogComponent
        ref={deleteDialogRef}
        id="ts-delete-dialog"
        header="Remove template?"
        showCloseIcon
        visible={false}
        target=".ts-app"
        width="420px"
        animationSettings={{ effect: 'Zoom', duration: 200 }}
      >
        <p className="ts-dialog-text">
          {deleteTarget ? `"${deleteTarget.name}" will be removed from the studio.` : ''}
        </p>
        <div className="ts-dialog-actions">
          <ButtonComponent
            cssClass="e-flat"
            onClick={() => { setDeleteTarget(null); deleteDialogRef.current?.hide(); }}
          >Cancel</ButtonComponent>
          <ButtonComponent
            cssClass="e-danger e-primary"
            onClick={confirmDelete}
          >Delete</ButtonComponent>
        </div>
      </DialogComponent>

      {/* Publish dialog: opens when the user clicks "Save and Publish" on
          a template that has never been published yet (no `docxUrl`, has
          `seedLines`). It is modeled after the Upload dialog so the user
          can confirm the template name and pick a Category/Type. The
          Browse row is disabled and labeled "From current document"
          because the .docx is the editor's content — there is no
          separate user file to pick. On confirm, the editor serializes
          its current content and POSTs the SFDT to the backend's
          DocumentEditorController.Save endpoint; App.jsx then rewrites
          the template's catalog entry to set docxUrl (and drop
          seedLines), so a reload loads the published .docx from
          wwwroot/Templates/ via the existing /studio-api catalog write. */}
      <DialogComponent
        ref={publishDialogRef}
        id="ts-publish-dialog"
        header="Save and Publish"
        showCloseIcon
        visible={false}
        target=".ts-app"
        width="480px"
        animationSettings={{ effect: 'Zoom', duration: 200 }}
        close={cancelPublish}
      >
        {isPublishDialogOpen && (
        <div className="ts-upload-form">
          {/* .docx file picker — disabled here. The .docx comes from the
              live editor (it has just been edited by the user), so there
              is no separate user file to pick. The Browse button is shown
              but disabled, matching the Upload dialog's shape. */}
          <label className="ts-upload-label">Template file (.docx)</label>
          <div className="ts-upload-file-row">
            <span
              className="ts-upload-file-name"
              title="The .docx is the current editor content"
            >
              {publishTarget
                ? `From current document (${publishTarget.name || 'untitled'})`
                : 'From current document'}
            </span>
          </div>
          <label className="ts-upload-label" htmlFor="ts-publish-name">Template name</label>
          <TextBoxComponent
            id="ts-publish-name"
            placeholder="e.g. Annual Appeal Letter"
            value={publishName}
            input={(e) => setPublishName(e.value ?? '')}
            floatLabelType="Never"
            disabled={isPublishing}
          />
          <label className="ts-upload-label" htmlFor="ts-publish-category">Category</label>
          <ComboBoxComponent
            id="ts-publish-category"
            cssClass="ts-upload-combobox"
            dataSource={existingCategories}
            value={publishCategory || 'General'}
            allowCustom={true}
            allowFiltering={true}
            placeholder="Select or type a new category"
            floatLabelType="Never"
            change={(e) => setPublishCategory(e.value ?? 'General')}
            input={(e) => setPublishCategory(e.value ?? '')}
            enabled={!isPublishing}
          />
          <label className="ts-upload-label" htmlFor="ts-publish-purpose">Purpose</label>
          <TextBoxComponent
            id="ts-publish-purpose"
            placeholder="Very short description shown on the card"
            value={publishPurpose}
            input={(e) => setPublishPurpose(e.value ?? '')}
            floatLabelType="Never"
            disabled={isPublishing}
          />
          {publishError && <p className="ts-dialog-text ts-publish-error">{publishError}</p>}
          <div className="ts-dialog-actions">
            <ButtonComponent
              cssClass="e-flat"
              onClick={cancelPublish}
              disabled={isPublishing}
            >Cancel</ButtonComponent>
            <ButtonComponent
              cssClass="e-primary"
              onClick={confirmPublish}
              disabled={isPublishing}
              iconCss={isPublishing ? 'e-icons e-refresh' : 'e-icons e-save'}
            >{isPublishing ? 'Publishing…' : 'Publish'}</ButtonComponent>
          </div>
        </div>
        )}
      </DialogComponent>

      {/* Confirmation dialog shown after a successful first-time publish.
          Mirrors the inline saveDialogRef used by the editor so the user
          gets a consistent "Template published" feedback, and gives them
          a one-click "Back to dashboard" affordance. */}
      <DialogComponent
        id="ts-published-dialog"
        header={publishedDialogMode === 'save' ? 'Template saved' : 'Template published'}
        showCloseIcon
        visible={showPublishedDialog}
        target=".ts-app"
        width="360px"
        animationSettings={{ effect: 'Zoom', duration: 200 }}
        close={() => { setShowPublishedDialog(false); }}
      >
        <p className="ts-dialog-text">
          {publishedDialogName
            ? `"${publishedDialogName}" has been ${publishedDialogMode === 'save' ? 'saved' : 'published'}.`
            : `The template has been ${publishedDialogMode === 'save' ? 'saved' : 'published'}.`}
        </p>
        <div className="ts-dialog-actions">
          <ButtonComponent
            cssClass="e-flat"
            onClick={() => { setShowPublishedDialog(false); }}
          >Stay</ButtonComponent>
          <ButtonComponent
            cssClass="e-primary"
            onClick={() => {
              setShowPublishedDialog(false);
              // Return to dashboard so the user sees the new card in
              // the list (with its refreshed thumbnail). The template's
              // docxUrl is now set, so on next reopen it loads from
              // wwwroot/Templates/ via fetchSfdtFromDocx.
              if (lastPublishedId) {
                setSelectedId(null);
              }
            }}
          >Back to dashboard</ButtonComponent>
        </div>
      </DialogComponent>
    </div>
  );
}

export default App;
