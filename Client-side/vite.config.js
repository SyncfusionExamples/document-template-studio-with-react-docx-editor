import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import { studioTemplateFiles } from './vite-plugin-studio-templates.js'

// The Vite dev server no longer proxies DOCX fetches: the React app
// talks directly to the ASP.NET Core DocumentEditor controller (see
// src/utils/studioStorage.js) using the host/port exported from
// src/data/sampleTemplates.js, and .docx files in `templates.json` are
// stored as `${DOCUMENT_EDITOR_BASE_URL}/Templates/<slug>.docx`.
// What this plugin still owns in dev:
//   - PUT  /studio-api/catalog        — persist templates.json
//   - GET  /studio-api/common-fields  — read common merge-field catalog
//   - POST /studio-api/mergefield     — write a custom merge field
//
// https://vite.dev/config/
export default defineConfig({
  plugins: [react(), studioTemplateFiles()],
})
