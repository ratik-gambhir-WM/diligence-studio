/// <reference types="vite/client" />

interface ImportMetaEnv {
  readonly VITE_API_BASE_URL?: string
  readonly VITE_APP_ID?: string
  readonly VITE_MICROSOFT_SUPPORT?: string
}

interface ImportMetaEnv {
  readonly VITE_OPENAI_API_KEY?: string
  readonly VITE_OPENAI_MODEL?: string
  readonly VITE_OPENAI_SECRET_KEY?: string
}

interface ImportMeta {
  readonly env: ImportMetaEnv
}

interface Window {
  __TTS_MERMAID_TEMPLATE_PREVIEW_INPUT__?: import('./lib/canvas-model/CanvasTypes').JsonValue
}
