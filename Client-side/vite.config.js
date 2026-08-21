import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import { studioTemplateFiles } from './vite-plugin-studio-templates.js'
// Single source of truth for the .NET backend host/port — also imported
// by vite-plugin-studio-templates.js, studioStorage.js, and
// thumbnailGenerator.js.
import { DOCUMENT_EDITOR_BASE_URL } from './src/data/sampleTemplates.js'

// https://vite.dev/config/
export default defineConfig({
  plugins: [react(), studioTemplateFiles()],
  server: {
    proxy: {
      // The built-in .docx templates are served by the Server-side .NET
      // app at /Templates/<file>.docx. Proxy those requests through the
      // Vite dev server so the client can keep using relative URLs.
      '/Templates': {
        target: DOCUMENT_EDITOR_BASE_URL,
        changeOrigin: true,
      },
    },
  },
})
