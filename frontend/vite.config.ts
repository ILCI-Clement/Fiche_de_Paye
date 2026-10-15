import react from '@vitejs/plugin-react'
import { defineConfig, loadEnv } from 'vite'
import { cpSync, mkdirSync, readFileSync } from 'node:fs'
import { createRequire } from 'node:module'
import { dirname, resolve } from 'node:path'

// Serve PDF.js fonts and decoders locally in both development and production.
const pdfRoot = dirname(createRequire(import.meta.url).resolve('pdfjs-dist/package.json'))
const pdfVersion = JSON.parse(readFileSync(resolve(pdfRoot, 'package.json'), 'utf8')).version as string
const pdfAssets = resolve(import.meta.dirname, 'public/pdfjs', pdfVersion)
mkdirSync(pdfAssets, { recursive: true })
for (const directory of ['cmaps', 'standard_fonts', 'wasm', 'iccs']) cpSync(resolve(pdfRoot, directory), resolve(pdfAssets, directory), { recursive: true })

// https://vite.dev/config/
export default defineConfig(({ mode }) => {
  const env = loadEnv(mode, process.cwd(), '')
  return {
    plugins: [react()],
    server: {
      host: '127.0.0.1', port: 5173, strictPort: true,
      proxy: { '/api': { target: env.PRESENCE_API_PROXY || 'http://127.0.0.1:8013',
        changeOrigin: true, rewrite: path => path.replace(/^\/api/, '') } },
    },
  }
})
