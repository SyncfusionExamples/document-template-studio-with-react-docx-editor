import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import { studioTemplateFiles } from './vite-plugin-studio-templates.js'

// The built-in .docx templates are served by the Server-sde .NET app
// at /Templates/<file>.docx. Proxy those requests through the Vite dev
// server so the client can keep using relative URLs.
const DOTNET_SERVER = 'http://localhost:5212';

// https://vite.dev/config/
export default defineConfig({
  plugins: [react(), studioTemplateFiles()],
  server: {
    proxy: {
      '/Templates': {
        target: DOTNET_SERVER,
        changeOrigin: true,
      },
    },
  },
})
