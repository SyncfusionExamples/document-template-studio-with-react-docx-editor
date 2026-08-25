// Merge field catalog, blank-template seed, and the .NET backend URL live
// here. The merge-field catalog is inlined below; everything else that
// needs to call the backend imports `DOCUMENT_EDITOR_BASE_URL` from this
// file so the host/port for the web API lives in one place.

// ---------------------------------------------------------------------------
// Backend Web API base URL.
//
// Local .NET server (Server-side / Program.cs) started with `dotnet run`,
// listening on http://localhost:5212/. Every cross-origin call from the
// React app — including the editor's own Import/Save/MailMerge path —
// targets `${DOCUMENT_EDITOR_BASE_URL}/api/DocumentEditor/`.
export const DOCUMENT_EDITOR_BASE_URL =
  import.meta.env.VITE_API_BASE_URL ||
  (import.meta.env.DEV ? 'http://localhost:5212' : '');

// Full Service URL the Syncfusion DocumentEditor container needs in its
// `serviceUrl` prop (note trailing slash — that's what the editor expects).
// Derived from the base above; do not hardcode elsewhere.
export const DOCUMENT_EDITOR_SERVICE_URL =
  `${DOCUMENT_EDITOR_BASE_URL}/api/DocumentEditor/`;

// ---------------------------------------------------------------------------
// Merge field catalog.
// Each entry maps a unique `FieldName` to itself ("true" means the field is
// recognized; the key IS the field name). Templates reference field names
// in their `fieldKeys` array. There's no separate display label, group,
// sample value, or repeating-block flag — a field is just its name. This
// keeps the catalog flat and removes the per-field metadata that the
// editor doesn't actually need (the .docx itself carries the content).
// ---------------------------------------------------------------------------
export const MERGE_FIELDS = {
  // Donor / recipient
  DonorName: true,
  DonorAddress: true,
  DonorEmail: true,

  // Donation
  DonationAmount: true,
  DonationDate: true,
  PaymentMethod: true,
  ReceiptNumber: true,

  // Invoice / Pledge specifics
  CustomerID: true,
  OrderID: true,
  InvoiceDate: true,

  // Tax / receipt
  TaxYear: true,
  TaxID: true,
  DeductibleAmount: true,

  // Organization
  OrgName: true,
  OrgAddress: true,
  OrgPhone: true,
  OrgEmail: true,
};

// All built-in .docx templates live exclusively on the server under
// Server-side/wwwroot/Templates/ and are served via
// `app.UseStaticFiles()`. The catalog (id, name, type, description,
// fieldKeys, docxUrl, thumbnailUrl) is owned by the .NET service and
// exposed at `${DOCUMENT_EDITOR_BASE_URL}/api/TemplateStudio/templates`
// — the React app does NOT import a local templates.json anymore
// (a previous version of this file did; the JSON was deleted when the
// server became the single source of truth). Thumbnails for every
// .docx-based template are generated client-side from the .docx —
// no per-type SVG art is used anywhere.

// Seeded merge fields for "+ New Template". The document body is always
// blank when a new template is created — only its merge-field catalog is
// pre-populated.
export const NEW_TEMPLATE_FIELD_KEYS = [
  'OrgName', 'OrgAddress', 'DonorName', 'DonorAddress', 'DonationAmount', 'DonationDate',
];
