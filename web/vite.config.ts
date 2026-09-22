import { defineConfig, loadEnv } from 'vite'
import tailwindcss from '@tailwindcss/vite'
import react from '@vitejs/plugin-react'

export default defineConfig(({ mode }) => {
  const env = loadEnv(mode, process.cwd(), '')

  const resolvedApiKey =
    env.VITE_OPENAI_API_KEY || env.VITE_OPENAI_SECRET_KEY || env.OPENAI_API_KEY || env.OPENAI_SECRET_KEY || ''
  const resolvedModel = env.VITE_OPENAI_MODEL || env.OPENAI_MODEL || ''

  return {
    plugins: [tailwindcss(), react()],
    server: {
      proxy: {
        '/api/v1': {
          target: env.VITE_API_PROXY_TARGET || 'http://127.0.0.1:43127',
          changeOrigin: false,
        },
        '/api': {
          changeOrigin: false,
          target: env.VITE_API_PROXY_TARGET || 'http://127.0.0.1:43127',
        },
      },
    },
    define: {
      'import.meta.env.VITE_OPENAI_API_KEY': JSON.stringify(resolvedApiKey),
      'import.meta.env.VITE_OPENAI_SECRET_KEY': JSON.stringify(resolvedApiKey),
      'import.meta.env.VITE_OPENAI_MODEL': JSON.stringify(resolvedModel),
    },
  }
})
