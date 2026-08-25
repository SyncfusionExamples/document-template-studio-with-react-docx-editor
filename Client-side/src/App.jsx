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
import {
  createTemplate,
  fetchTemplate,
  fetchTemplates,
  updateTemplate,
  deleteTemplate,
  addCommonMergeField,
  addTemplateMergeField,
  fetchCommonMergeFields,
} from './utils/studioStorage.js';
import './App.css';

// Build an absolute `${DOCUMENT_EDITOR_BASE_URL}/Templates/<slug>.docx`
// URL from a stored `docxUrl`. Stored values come in two shapes:
//   1. absolute `http(s)://host[:port]/Templates/<slug>.docx` (current
//      multi-machine shape), or
//   2. relative `/Templates/<slug>.docx` (legacy shape, kept readable
//      by `Server-side/wwwroot/Templates/templates.json` shipped from
//      older builds).
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

// Helper to generate a unique id for new (blank) templates.
function makeId(prefix = 'tpl') {
  return `${prefix}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 7)}`;
}

// Generate the stable server-side slug used by the automatic first save
// for an uploaded template. The same slug is stored in the template id
// (`tpl-<slug>`) so the subsequent Save & Publish flow overwrites the
// same DOCX file.
function makeUploadSlug(name) {
  const base = (name || 'template').replace(/\.[^.]+$/, '').trim();
  const cleaned = base.replace(/[^A-Za-z0-9-_]+/g, '_').replace(/^_+|_+$/g, '');
  return `${cleaned || 'template'}-${Date.now().toString(36)}`;
}

function App() {
  const [templates, setTemplates] = useState([]);
  // Common (global) merge fields loaded from the .NET service's
  // /api/TemplateStudio/merge-fields/common endpoint on startup.
  // Forwarded as a prop into TemplateViewer so the Merge Fields
  // panel can render every common field on every template. Stays
  // an empty object when the file is empty/missing — the panel
  // handles an empty map without crashing.
  const [commonFields, setCommonFields] = useState({});
  const [isLoadingTemplates, setIsLoadingTemplates] = useState(true);
  const [isLoadingCatalog, setIsCatalogLoading] = useState(true);
  const [selectedId, setSelectedId] = useState(null);
  const deleteDialogRef = useRef(null);
  const [deleteTarget, setDeleteTarget] = useState(null);

  useEffect(() => {
    let cancelled = false;

    const loadStudioData = async () => {
      setIsCatalogLoading(true);

      try {
        const [templatesResult, commonFieldsResult] =
          await Promise.allSettled([
            fetchTemplates(),
            fetchCommonMergeFields(),
          ]);

        if (cancelled) return;

        if (templatesResult.status === 'fulfilled') {
          setTemplates(templatesResult.value);
        } else {
          throw templatesResult.reason;
        }

        if (commonFieldsResult.status === 'fulfilled') {
          setCommonFields(commonFieldsResult.value);
        } else {
          console.warn(
            'Common merge fields could not be loaded:',
            commonFieldsResult.reason,
          );
        }
      } catch (error) {
        if (cancelled) return;

        console.error(
          'Template Studio initialization failed:',
          error,
        );

        alert(
          `Unable to load Template Studio data: ${
            error.message || error
          }`,
        );
      } finally {
        if (!cancelled) {
          setIsCatalogLoading(false);
        }
      }
    };

    loadStudioData();

    return () => {
      cancelled = true;
    };
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
  // "+ New Template" also needs an executor — pressing the +
  // button creates a transient template, mounts the viewer, and the
  // viewer's registered `addExecutor` does the actual
  // editor.Serialize + DocumentEditorController.Save round-trip.
  // App.jsx then POSTs the resulting `docxUrl` + `thumbnailUrl` into
  // CreateTemplate so the new catalog row has the on-disk file path
  // (no separate "Save and Publish" press required).
  const addExecuteRef = useRef(null);
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
      // Mirror the React-state update onto the server's catalog so a
      // page reload picks up docxUrl + name + type without a fresh
      // POST. Best-effort: a server-side failure leaves the React
      // row updated, and the user simply has to save again.
      try {
        await updateTemplate(publishTarget.id, {
          name: publishName.trim(),
          type: (publishCategory || '').trim() || 'General',
          description:
            (publishPurpose || '').trim() || publishTarget.description,
        });
      } catch (syncErr) {
        // eslint-disable-next-line no-console
        console.warn('[publish] server catalog sync failed:', syncErr);
      }
      // The catalog is persisted server-side via PUT to the .NET
      // TemplateStudioController.UpdateTemplate endpoint, which
      // rewrites Server-side/wwwroot/Templates/templates.json so a
      // reload picks up docxUrl.
      cancelPublish();
    } catch (err) {
      // eslint-disable-next-line no-console
      console.error('Publish failed:', err);
      setPublishError(err?.message || String(err));
    } finally {
      setIsPublishing(false);
    }
  }, [publishTarget, publishName, publishCategory, publishPurpose, cancelPublish]); 
  

  // "+ New Template" creates a blank template doc + its merge-field set,
  // places it in the sidebar list, and opens it immediately so the user
  // can design the letter and Save (persist to the data folder). The
  // template body starts empty — the editor opens with `openBlank()`
  // (see TemplateViewer.loadTemplateIntoEditor) and only the merge-field
  // catalog is pre-populated.
  //
  // The flow now (per requirement) ALSO calls the
  // DocumentEditorController.Save API automatically and registers the
  // resulting `docxUrl` into the .NET catalog so the new template is
  // immediately backed by an on-disk .docx (no separate Save & Publish
  // press required). This guarantees the dashboard card has a real
  // docxUrl from the moment the template is added.
  const handleAdd = useCallback(async () => {
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
    // The TemplateViewer mounts on the next render and registers the
    // addExecutor via a useEffect. Wait briefly so the ref is set
    // before we try to invoke it; if it never shows up, fall back to
    // the legacy behavior (template still in React state, no docxUrl
    // until the user presses Save & Publish manually).
    let docxBaseName = '';
    let thumbnailDataUri = '';
    let addAttempted = false;
    for (let i = 0; i < 30; i++) {
      if (typeof addExecuteRef.current === 'function') {
        addAttempted = true;
        try {
          const result = await addExecuteRef.current({
            name: tpl.name,
          });
          docxBaseName = result?.docxBaseName || '';
          thumbnailDataUri = result?.thumbnailDataUri || '';
        } catch (err) {
          // eslint-disable-next-line no-console
          console.error('[add] auto-save round-trip failed:', err);
        }
        break;
      }
      // 100 ms between polls, max ~3 s — covers the typical
      // DocumentEditorContainer `created` -> `useEffect` race.
      // eslint-disable-next-line no-await-in-loop
      await new Promise((r) => setTimeout(r, 100));
    }
    if (!addAttempted) {
      // eslint-disable-next-line no-console
      console.warn('[add] addExecutor never registered; falling back to manual save.');
      return;
    }
    if (!docxBaseName) {
      // Round-trip failed or returned an empty slug — leave the
      // template as-is in React state (the user can still press
      // "Save and Publish" manually to retry).
      return;
    }
    const now = new Date().toISOString();
    const docxUrl = absoluteDocxUrl(`/Templates/${docxBaseName}.docx`);
    try {
      const saved = await createTemplate({
        id: tpl.id,
        name: tpl.name,
        type: tpl.type,
        description: tpl.description,
        fileName: `${docxBaseName}.docx`,
        docxUrl,
        thumbnailUrl: thumbnailDataUri || '',
        fieldKeys: [...tpl.fieldKeys],
        createdAt: now,
        updatedAt: now,
      });
      // Replace the transient React-state entry with whatever the
      // server returned so docxUrl/thumbnailUrl/fileName are sourced
      // from one row of truth.
      setTemplates((prev) =>
        prev.map((existing) =>
          existing.id === tpl.id ? { ...existing, ...saved } : existing,
        ),
      );
    } catch (err) {
      // eslint-disable-next-line no-console
      console.error('[add] CreateTemplate failed:', err);
      // Best-effort: the .docx is already on disk from the editor's
      // Save round-trip, so we can still patch the React row to
      // carry docxUrl even if the catalog POST fails. Subsequent
      // page reloads will surface the on-disk file via
      // wwwroot/Templates/<filename>.docx.
      setTemplates((prev) =>
        prev.map((existing) =>
          existing.id === tpl.id
            ? {
                ...existing,
                docxUrl,
                thumbnailUrl: thumbnailDataUri || existing.thumbnailUrl || '',
                fileName: `${docxBaseName}.docx`,
              }
            : existing,
        ),
      );
    }
  }, []);

  // ---- Upload flow (requirement 1) ----
  // The uploaded File is kept only in App state while the TemplateViewer
  // performs the document-level initialization. No Vite filesystem
  // upload is performed. The transient template gets a stable id of
  // `tpl-<slug>` so the automatic Save and all later saves target the
  // same DOCX.
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

  // The user confirmed the dialog. We create only a transient React
  // template and route the original File to TemplateViewer.
  // TemplateViewer then:
  //   1. imports DOCX -> SFDT,
  //   2. opens the SFDT in the live DocumentEditor,
  //   3. captures the thumbnail from that same editor, and
  //   4. performs the initial DocumentEditor Save -> DOCX.
  // Only after those steps succeed do we remove the transient marker
  // and allow the normal [templates] persistence effect to write
  // templates.json.
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

    // Stable slug becomes the template id (`tpl-<slug>`) AND the
    // DocumentEditor Save FileName, so a later Save & Publish overwrites
    // the same DOCX that the initial upload wrote.
    const slug = makeUploadSlug(name);
    const templateId = `tpl-${slug}`;

    const transientTemplate = {
      id: templateId,
      name,
      type: category,
      description,
      fieldKeys: [...NEW_TEMPLATE_FIELD_KEYS],
      docxUrl: '',
      thumbnailUrl: '',
      _pendingUpload: true,           // skip persistence until Save succeeds
      _autoDescription: !purposeWasProvided, // refresh description with page count
    };

    uploadDialogRef.current?.hide();
    setIsUploadDialogOpen(false);
    setIsUploading(true);

    setTemplates((prev) => {
      const existingIds = new Set(prev.map((t) => t.id));
      if (existingIds.has(templateId)) return prev;
      return [...prev, transientTemplate];
    });
    setPendingUpload({ file, templateId });
    setSelectedId(templateId);
  }, [pendingUpload, uploadName, uploadCategory, uploadPurpose]);

  const handleInitialUploadSaved = useCallback(async ({
    templateId,
    name,
    type,
    description,
    fieldKeys = [],
    docxBaseName,
    thumbnailDataUri = '',
    pageCount = 0,
  } = {}) => {
    if (!templateId || !docxBaseName) {
      setIsUploading(false);
      setPendingUpload(null);
      return;
    }

    try {
      let finalDescription = description;

      if (!finalDescription && pageCount > 0) {
        finalDescription =
          `Uploaded .docx (${pageCount} page${pageCount > 1 ? 's' : ''})`;
      }

      const now = new Date().toISOString();
      // The actual .docx was already written to wwwroot/Templates/ by
      // the DocumentEditorController.Save call from
      // TemplateViewer.initializeUploadedTemplate. Build the canonical
      // URL the editor will use on reopen so the catalog row carries a
      // real docxUrl (per requirement: "After adding a template or
      // upload a template, it should call DocumentEditorController's
      // Save API internally and append the docxUrl").
      const docxUrl = absoluteDocxUrl(`/Templates/${docxBaseName}.docx`);

      const templateMetadata = {
        id: templateId,
        name,
        type: type || 'General',
        description: finalDescription || null,
        fileName: `${docxBaseName}.docx`,
        docxUrl,
        // Single canonical thumbnail field — `thumbnailUrl`. The
        // server previously also accepted `thumbnail`; that alias has
        // been removed in favour of just `thumbnailUrl`.
        thumbnailUrl: thumbnailDataUri || '',
        fieldKeys,
        createdAt: now,
        updatedAt: now,
      };

      const savedTemplate =
        await createTemplate(templateMetadata);

      setTemplates((prev) =>
        prev.map((t) =>
          t.id === templateId
            ? {
                ...t,
                ...savedTemplate,
              }
            : t,
        ),
      );

      setPendingUpload(null);
      setIsUploading(false);
    } catch (error) {
      console.error(
        '[upload] failed to persist template metadata:',
        error,
      );

      // Even when the CreateTemplate POST fails, the .docx is
      // already on disk from the editor's Save round-trip. Patch
      // the React row with the resolved docxUrl so subsequent
      // reloads still pick up the on-disk file via
      // wwwroot/Templates/.
      const fallbackDocxUrl = docxBaseName
        ? absoluteDocxUrl(`/Templates/${docxBaseName}.docx`)
        : '';
      setTemplates((prev) =>
        prev.map((t) =>
          t.id === templateId
            ? {
                ...t,
                docxUrl: t.docxUrl || fallbackDocxUrl,
                thumbnailUrl:
                  t.thumbnailUrl || thumbnailDataUri || '',
                fileName: t.fileName || `${docxBaseName}.docx`,
              }
            : t,
        ),
      );

      setPendingUpload(null);
      setIsUploading(false);

      alert(
        `Template upload completed, but metadata could not be saved: ${
          error?.message || error
        }`,
      );
    }
  }, []);

  // Failure callback fired by TemplateViewer when any step of the
  // Import -> open -> initial Save sequence errors out. We drop the
  // transient template entirely (no docxUrl = not a valid catalog
  // entry) and return the user to the dashboard with an alert.
  const handleInitialUploadFailed = useCallback((error) => {
    setTemplates((prev) => prev.filter((t) => !t._pendingUpload));
    setSelectedId(null);
    setPendingUpload(null);
    setIsUploading(false);

    const message = error?.message || String(error || 'Unknown error');
    // eslint-disable-next-line no-console
    console.error('[upload] initial template initialization failed:', error);
    alert(`Upload failed: ${message}`);
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
  const handleDelete = useCallback(async (id) => {
  try {
    await deleteTemplate(id);

    setTemplates((prev) =>
      prev.filter((template) => template.id !== id),
    );

    if (selectedId === id) {
      setSelectedId(null);
    }
  } catch (error) {
    console.error('Delete template failed:', error);
    alert(error.message || 'Failed to delete template.');
  }
}, [selectedId]);

  const confirmDelete = useCallback(async () => {
    if (!deleteTarget) return;

    const id = deleteTarget.id;
    const wasSelected = id === selectedId;

    try {
      await deleteTemplate(id);

      setTemplates((prev) =>
        prev.filter((template) => template.id !== id)
      );

      if (wasSelected) {
        setSelectedId(null);
      }

      setDeleteTarget(null);
      deleteDialogRef.current?.hide();
    } catch (error) {
      console.error('Failed to delete template:', error);

      alert(
        error?.message ||
        'Failed to delete the template.'
      );
    }
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
            registerAddExecutor={(fn) => { addExecuteRef.current = fn; }}
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
