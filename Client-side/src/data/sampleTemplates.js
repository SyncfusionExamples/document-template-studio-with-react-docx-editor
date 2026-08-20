// All .docx-based templates (built-in samples + user-uploaded) now live as
// JSON metadata in src/data/user-templates/ and are loaded uniformly via
// /studio-api/list — see App.jsx. The original .docx files for the built-in
// samples remain in this folder (src/data/) and are referenced by the
// `docxUrl` field in their JSON metadata. Only the merge-field catalog,
// the blank-template seed, and the Syncfusion service URL live here now.

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

// Local DocumentEditor web service (Server-sde / Program.cs) hosting the
// Syncfusion-compatible Import endpoint that converts a .docx into SFDT
// (so DocumentEditor.open() can load it). The server is started with
// `dotnet run` and listens on http://localhost:5212/.
export const DOCUMENT_EDITOR_SERVICE_URL =
  'http://localhost:5212/api/DocumentEditor/';

// All built-in .docx templates live on the server under
// Server-sde/wwwroot/Templates/ and are served via `app.UseStaticFiles()`.
// The single client-side catalog (src/data/templates.json) maps each
// template's id, name, type, description, and fieldKeys to the URL where
// the .docx can be fetched from. Thumbnails for every .docx-based
// template are generated client-side from the .docx — no per-type SVG
// art is used anywhere.

// A blank template's seeded merge fields and content used by "+ New Template".
export const NEW_TEMPLATE_FIELD_KEYS = [
  'OrgName', 'OrgAddress', 'DonorName', 'DonorAddress', 'DonationAmount', 'DonationDate',
];
export const NEW_TEMPLATE_SEED_LINES = [];
