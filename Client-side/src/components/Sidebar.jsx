import { useMemo, useState } from 'react';
import { ButtonComponent } from '@syncfusion/ej2-react-buttons';
import { TextBoxComponent } from '@syncfusion/ej2-react-inputs';

const TYPE_COLORS = {
  Invoice:     '#d6336c',
  ThankYou:    '#2b8a3e',
  TaxReceipt:  '#1971c2',
  General:     '#5c3eb1',
};

// Pretty labels for the type chips. Falls back to the raw type when a
// template has a custom category that isn't in the built-in list.
const TYPE_LABELS = {
  Invoice: 'Invoice',
  ThankYou: 'Thank You',
  TaxReceipt: 'Tax Receipt',
  General: 'General',
};

// Build ordered groups from the current template list. Built-in categories
// (Invoice, Thank You, Tax Receipt, General) appear in a fixed order; any
// new / custom categories are appended alphabetically below. Templates
// without a recognized `type` are routed to the nearest custom bucket and
// never surface an "Uncategorized" group.
function groupTemplates(list) {
  const order = ['General', 'ThankYou', 'TaxReceipt', 'Invoice'];
  const known = new Set(order);
  const groups = [];
  const seen = new Set();
  const add = (key, label) => {
    if (seen.has(key)) return;
    seen.add(key);
    groups.push({ key, label, items: [] });
  };
  // Pre-seed the known groups in the desired order.
  for (const k of order) add(k, TYPE_LABELS[k] || k);

  const customKeys = new Set();
  for (const t of list) {
    const k = t.type || '';
    if (k && !known.has(k)) customKeys.add(k);
  }
  Array.from(customKeys).sort().forEach((k) => add(k, k));

  for (const t of list) {
    // Templates without a type (or with an empty type) are dropped — the
    // sidebar only ever shows recognized categories, so there is no
    // "Uncategorized" fallback.
    const key = t.type || '';
    if (!key) continue;
    const g = groups.find((gr) => gr.key === key);
    if (g) g.items.push(t);
  }
  // Drop empty groups so the sidebar doesn't show empty headers.
  return groups.filter((g) => g.items.length > 0);
}

// Sidebar: list of templates (dashboard) using Syncfusion EJ2 controls.
// - Search: TextBoxComponent.
// - "+ New Template": ButtonComponent — creates a blank template.
// - "+ Upload .docx": ButtonComponent — opens the file picker so the user
//   can import an existing DOCX file into the studio (requirement 1).
// - Templates are grouped by `type` (General, Thank You, Tax Receipt,
//   Invoice, ...). The first three letters of each template's display
//   name are shown in the colored bar, so groups are easy to scan.
// - Per-row delete: small ButtonComponent (icon-only). Actual confirmation
//   happens in App via a Syncfusion Dialog.
function Sidebar({ templates, selectedId, onSelect, onAdd, onUpload, isUploading, onDelete }) {
  const [query, setQuery] = useState('');

  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase();
    if (!q) return templates;
    return templates.filter(
      (t) => t.name.toLowerCase().includes(q) || t.type.toLowerCase().includes(q),
    );
  }, [templates, query]);

  const groups = useMemo(() => groupTemplates(filtered), [filtered]);

  return (
    <aside className="ts-sidebar">
      <header className="ts-sidebar-head">
        <h2>Template Studio</h2>
        <TextBoxComponent
          placeholder="Search templates"
          value={query}
          input={(e) => setQuery(e.value ?? '')}
          cssClass="ts-search"
          floatLabelType="Never"
        />
      </header>

      <nav className="ts-tree" aria-label="Templates list">
        {groups.length === 0 && (
          <p className="ts-empty">No templates match your search.</p>
        )}
        {groups.map((g) => (
          <section key={g.key} className="ts-tree-group">
            <header className="ts-tree-group-head">
              <span
                className="ts-type-dot"
                style={{ background: TYPE_COLORS[g.key] ?? '#5c3eb1' }}
                aria-hidden="true"
              />
              <span className="ts-tree-group-label">{g.label}</span>
              <span className="ts-tree-group-count">{g.items.length}</span>
            </header>
            <ul className="ts-tree-group-list">
              {g.items.map((t) => {
                const isSel = t.id === selectedId;
                return (
                  <li key={t.id}>
                    <div
                      className={`ts-tree-row${isSel ? ' is-selected' : ''}`}
                      onClick={() => onSelect(t.id)}
                      role="button"
                      tabIndex={0}
                      onKeyDown={(e) => {
                        if (e.key === 'Enter' || e.key === ' ') {
                          e.preventDefault();
                          onSelect(t.id);
                        }
                      }}
                    >
                      <span className="ts-file-icon" aria-hidden="true">📄</span>
                      <span className="ts-tree-text">
                        <span className="ts-tree-name">{t.name}</span>
                      </span>
                      <ButtonComponent
                        iconCss="e-icons e-close-icon"
                        cssClass="e-flat ts-icon-btn ts-delete"
                        title={`Remove ${t.name}`}
                        onClick={(e) => {
                          e.originalEvent?.stopPropagation();
                          onDelete(t.id);
                        }}
                      />
                    </div>
                  </li>
                );
              })}
            </ul>
          </section>
        ))}
      </nav>
    </aside>
  );
}

export default Sidebar;
