import { ButtonComponent } from '@syncfusion/ej2-react-buttons';

const TYPE_COLORS = {
  Invoice:    '#d6336c',
  ThankYou:   '#2b8a3e',
  TaxReceipt: '#1971c2',
  General:    '#5c3eb1',
};

// Dashboard: thumbnail grid of templates. Clicking a card opens that
// template in the editor (main area). Provides a "+ New Template" card and
// an "Upload Template" card. Every template — built-in or uploaded —
// displays a PNG thumbnail generated from the .docx file itself
// (requirement: no hardcoded SVG thumbnails). Each card also has a trash
// button (top-right) to remove the template after confirmation.
function Dashboard({ templates, onOpen, onAdd, onUpload, onDelete, isUploading }) {
  return (
    <section className="ts-dashboard">
      <header className="ts-dashboard-head">
        <h2>Template Studio</h2>
        <p>Create, manage, and edit your document templates in one place.</p>
      </header>

      <div className="ts-thumb-grid">
        {/* "+ New Template" card */}
        <button
          type="button"
          className="ts-thumb ts-thumb-new"
          onClick={onAdd}
          aria-label="Create a new template"
        >
          <span className="ts-thumb-new-icon" aria-hidden="true">+</span>
          <span className="ts-thumb-name">New Template</span>
          <span className="ts-thumb-type">Blank document</span>
        </button>

        {/* "+ Upload Template" card (requirement 1) */}
        <button
          type="button"
          className="ts-thumb ts-thumb-upload"
          onClick={onUpload}
          disabled={isUploading}
          aria-label="Upload a .docx template"
        >
          <span className="ts-thumb-new-icon" aria-hidden="true">
            {isUploading ? '⟳' : '⬆'}
          </span>
          <span className="ts-thumb-name">
            {isUploading ? 'Uploading…' : 'Upload Template'}
          </span>
          <span className="ts-thumb-type">Import existing template</span>
        </button>

        {templates.map((t) => {
          const color = TYPE_COLORS[t.type] ?? '#5c3eb1';
          const hasImage = Boolean(t.thumbnailUrl);
          // The card is a div[role=button] (not a <button>) so the trash
          // ButtonComponent inside it (a real <button>) doesn't violate the
          // HTML rule that buttons can't contain nested buttons.
          // The card's onClick checks whether the click came from the trash
          // button (or its wrapper) and, if so, skips opening the editor —
          // the trash button's own handler takes care of the delete flow.
          return (
            <div
              key={t.id}
              className="ts-thumb"
              onClick={(e) => {
                if (e.target.closest('.ts-thumb-delete-wrap')) return;
                onOpen(t.id);
              }}
              onKeyDown={(e) => {
                if (e.key === 'Enter' || e.key === ' ') {
                  e.preventDefault();
                  onOpen(t.id);
                }
              }}
              role="button"
              tabIndex={0}
              style={{ '--ts-thumb-color': color }}
              aria-label={`Open ${t.name}`}
              title={t.name}
            >
              {/* Trash button: removes the template (with confirmation)
                  and deletes its .docx + thumbnail + meta from disk.
                  The wrapper span's stopPropagation keeps the card's
                  onClick from firing when the trash is clicked. */}
              <span
                className="ts-thumb-delete-wrap"
                onClick={(e) => { e.stopPropagation(); }}
              >
                <ButtonComponent
                  iconCss="e-icons e-trash"
                  cssClass="e-flat ts-thumb-delete ts-icon-btn ts-delete"
                  title={`Remove ${t.name}`}
                  aria-label={`Remove ${t.name}`}
                  onClick={(e) => {
                    e.originalEvent?.stopPropagation();
                    e.originalEvent?.preventDefault();
                    onDelete?.(t.id);
                  }}
                />
              </span>
              {hasImage ? (
                <span className="ts-thumb-art ts-thumb-art-image">
                  <img src={t.thumbnailUrl} alt="" />
                </span>
              ) : (
                // Placeholder while the .docx thumbnail is being generated
                // client-side (or for blank templates that have no .docx).
                <span className="ts-thumb-art ts-thumb-art-placeholder">
                  <span className="ts-thumb-placeholder-doc" aria-hidden="true">
                    <span className="ts-thumb-placeholder-line" />
                    <span className="ts-thumb-placeholder-line" />
                    <span className="ts-thumb-placeholder-line" />
                    <span className="ts-thumb-placeholder-line" />
                  </span>
                </span>
              )}
              <span className="ts-thumb-name" title={t.name}>{t.name}</span>
              <span className="ts-thumb-type">
                <span className="ts-type-dot" style={{ background: color }} />
                {t.type}
              </span>
              <span className="ts-thumb-desc" title={t.description}>{t.description}</span>
            </div>
          );
        })}
      </div>
    </section>
  );
}

export default Dashboard;
