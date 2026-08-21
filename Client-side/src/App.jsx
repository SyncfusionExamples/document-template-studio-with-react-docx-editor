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
} from './data/sampleTemplates.js';
// Single source of truth for every .docx-backed template the studio
// ships with. Each entry's `docxUrl` points at the Server-sde static
// file in `wwwroot/Templates/` (served via `app.UseStaticFiles()`).
import TEMPLATE_CATALOG from './data/templates.json';
import {
  uploadTemplate,
  deleteTemplateFiles,
  saveTemplatesCatalog,
  getHiddenBuiltInIds,
  hideBuiltInTemplate,
  fetchCommonMergeFields,
} from './utils/studioStorage.js';
import './App.css';

// Helper to generate a unique id for new (blank) templates.
function makeId(prefix = 'tpl') {
  return `${prefix}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 7)}`;
}

// The ids of the original built-in entries that ship with templates.json.
// We use this to detect "is this id from the static catalog?" so deleting
// a built-in is treated as a client-side hide (no .docx unlink) while
// deleting a user-uploaded template removes its .docx from the server.
const BUILTIN_IDS = new Set(TEMPLATE_CATALOG.map((t) => t.id));

function App() {
  // Seed the catalog from the static templates.json shipped with the app.
  // These .docx files live on the server (Server-sde/wwwroot/Templates/)
  // and are served via `app.UseStaticFiles()`. The mapping below tells
  // the client where each .docx can be fetched from. Built-in templates
  // the user has previously removed are filtered out via localStorage so
  // the deletion persists across reloads / service restarts.
  const [templates, setTemplates] = useState(() => {
    const hidden = new Set(getHiddenBuiltInIds());
    return TEMPLATE_CATALOG.filter((t) => !hidden.has(t.id));
  });
  // null  = dashboard view; an id = that template opened in the editor.
  const [selectedId, setSelectedId] = useState(null);
  const deleteDialogRef = useRef(null);
  const [deleteTarget, setDeleteTarget] = useState(null);

  // Global (common) merge-field catalog, loaded once at app startup from
  // the dev server. The same source of truth is forwarded to every
  // TemplateViewer so that adding a common field on one template is
  // immediately visible in every other template's panel and persists
  // across reloads (the server writes to src/data/common-merge-fields.json).
  // `commonFields` is a flat { key: true } map of recognized keys.
  const [commonFields, setCommonFields] = useState({});
  useEffect(() => {
    let cancelled = false;
    fetchCommonMergeFields().then((m) => {
      if (!cancelled) setCommonFields(m || {});
    });
    return () => { cancelled = true; };
  }, []);

  const selected = templates.find((t) => t.id === selectedId) ?? null;

  // Generate PNG thumbnails for every template that has a .docx but no
  // thumbnail yet — this covers the built-in samples too, so no hardcoded
  // SVG thumbnails are used anywhere (requirement 4 applied uniformly).
  // Runs whenever the template list changes; a ref guards against
  // re-processing the same id (only marked after success, so a cancelled
  // attempt can be retried), and we mutate state one-at-a-time as each
  // off-screen DocumentEditor finishes.
  const processedThumbnailIds = useRef(new Set());
  useEffect(() => {
    let cancelled = false;
    const pending = templates.filter(
      (t) => t.docxUrl && !t.thumbnailUrl && !processedThumbnailIds.current.has(t.id),
    );
    if (pending.length === 0) return;
    (async () => {
      const { generateThumbnailFromUrl } = await import('./utils/thumbnailGenerator.js');
      for (const tpl of pending) {
        if (cancelled) break;
        try {
          const { thumbnailDataUri } = await generateThumbnailFromUrl(tpl.docxUrl);
          if (cancelled || !thumbnailDataUri) continue;
          // Mark processed only after we have a result, so a cancelled/
          // failed attempt can be retried on the next effect run.
          processedThumbnailIds.current.add(tpl.id);
          setTemplates((prev) => prev.map((t) => (t.id === tpl.id ? { ...t, thumbnailUrl: thumbnailDataUri } : t)));
        } catch (err) {
          // eslint-disable-next-line no-console
          console.warn(`[thumbnail] failed for ${tpl.id}:`, err);
          // Don't mark processed — allow a retry on a future run.
        }
      }
    })();
    return () => { cancelled = true; };
  }, [templates]);

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
          docxUrl: `/Templates/${slug}.docx`,
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

  // ----- Persist the template collection back to disk -----
  // templates.json is the single source of truth on the client side and
  // its on-disk copy lives at src/data/templates.json. The dev plugin
  // exposes PUT /studio-api/catalog to write the file. Skip the very first
  // render so we don't immediately rewrite the file we just read.
  //
  // The full collection (including the `thumbnailUrl` data URI on each
  // entry) is persisted as-is so the on-disk catalog is the single home of
  // thumbnail images too — no separate .png files, no re-generation on
  // every load. Once a thumbnail has been rendered for a template it is
  // cached inside the catalog entry for subsequent reloads.
  const firstRenderRef = useRef(true);
  useEffect(() => {
    if (firstRenderRef.current) {
      firstRenderRef.current = false;
      return;
    }
    // NOTE: the thumbnail data URIs are intentionally kept here so the
    // catalog IS the persistence layer for thumbnails.
    saveTemplatesCatalog(templates).catch((err) => {
      // eslint-disable-next-line no-console
      console.warn('[catalog] saveTemplatesCatalog failed:', err);
    });
  }, [templates]);

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

  // ---- Upload flow (requirement 1) ----
  // The "Upload Template" buttons open a dialog that collects the .docx
  // file (via an in-dialog Browse button), a Name, a Category (the user can
  // pick an existing one OR type a new one), and a very short Purpose.
  // On confirm, a thumbnail is generated client-side (via
  // DocumentEditor.exportAsImage + ej2-pdf-export) and both are POSTed to
  // the dev middleware which writes them into src/data/user-templates/.
  const [isUploading, setIsUploading] = useState(false);
  const uploadDialogRef = useRef(null);
  const dialogFileInputRef = useRef(null);
  // Pending upload: the picked file is held here while the details dialog
  // collects the Name / Category / Purpose from the user.
  const [pendingUpload, setPendingUpload] = useState(null); // { file }
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

  // Open the Upload Template dialog directly (the file is picked from
  // INSIDE the dialog via its Browse button).
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
    setPendingUpload({ file });
  }, []);

  // User confirmed the dialog → validate, generate thumbnail + persist.
  const confirmUpload = useCallback(async () => {
    const file = pendingUpload?.file;
    if (!file) {
      alert('Please choose a .docx file to upload.');
      return;
    }
    uploadDialogRef.current?.hide();
    setIsUploadDialogOpen(false);
    setIsUploading(true);
    try {
      const { generateThumbnailFromFile } = await import('./utils/thumbnailGenerator.js');
      let thumbnailDataUri = '';
      let pageCount = 0;
      try {
        const out = await generateThumbnailFromFile(file);
        thumbnailDataUri = out.thumbnailDataUri || '';
        pageCount = out.pageCount || 0;
      } catch (err) {
        // Thumbnail generation is best-effort — upload should still succeed.
        // eslint-disable-next-line no-console
        console.warn('Thumbnail generation failed:', err);
      }

      // Purpose is a very short description shown at the bottom of the card.
      // If the user left it blank, fall back to a page-count summary so the
      // card always has meaningful footer text (matching the built-ins).
      const purpose = (uploadPurpose || '').trim()
        || (pageCount
          ? `Uploaded .docx (${pageCount} page${pageCount > 1 ? 's' : ''})`
          : 'Uploaded .docx template');

      // Category: use the entered text as-is (existing OR brand new). Fall
      // back to "General" if the user left it blank.
      const category = (uploadCategory || '').trim() || 'General';

      const saved = await uploadTemplate({
        file,
        name: (uploadName || '').trim() || file.name.replace(/\.docx$/i, ''),
        type: category,
        description: purpose,
        fieldKeys: [...NEW_TEMPLATE_FIELD_KEYS],
        thumbnailDataUri,
      });

      setTemplates((prev) => {
        const existingIds = new Set(prev.map((t) => t.id));
        return existingIds.has(saved.id) ? prev : [...prev, saved];
      });
      setSelectedId(saved.id);
    } catch (err) {
      // eslint-disable-next-line no-console
      console.error('Upload failed:', err);
      alert(`Upload failed: ${err.message || err}`);
    } finally {
      setIsUploading(false);
      setPendingUpload(null);
    }
  }, [pendingUpload, uploadName, uploadCategory, uploadPurpose]);

  const cancelUpload = useCallback(() => {
    uploadDialogRef.current?.hide();
    setIsUploadDialogOpen(false);
    setPendingUpload(null);
  }, []);

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
    // If the template is one of the built-ins (i.e. it came from the
    // static templates.json catalog), remember that the user removed it
    // so it stays hidden on subsequent reloads / service restarts. The
    // .docx file on the server is left intact — hiding is purely a
    // client-side preference so the user can restore the built-ins
    // (e.g. by clearing localStorage) without us having to repopulate
    // wwwroot/Templates/.
    const isBuiltin = BUILTIN_IDS.has(id);
    if (isBuiltin) {
      hideBuiltInTemplate(id);
    } else {
      // User-uploaded templates: remove the underlying .docx from the
      // server so it doesn't take up disk space.
      try { await deleteTemplateFiles(id); } catch { /* noop */ }
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
            <ButtonComponent
              iconCss="e-icons e-folder-open"
              cssClass="e-outline ts-upload-browse"
              disabled
            >Browse</ButtonComponent>
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
