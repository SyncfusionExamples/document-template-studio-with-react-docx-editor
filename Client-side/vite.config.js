import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// The React app talks directly to the ASP.NET Core service for
// everything: DocumentEditor import/save/mail-merge AND the template
// catalog (templates.json + common-merge-fields.json) under
// wwwroot/Data/. The previous Vite dev plugin that proxied
// /studio-api/catalog and friends has been removed — the ASP.NET Core
// service is now the single authoritative owner of the catalog and the
// .docx files (see Server-side/Controllers/StudioController.cs).
//
// https://vite.dev/config/
export default defineConfig({
  plugins: [react()],
})
