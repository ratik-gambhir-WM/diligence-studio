import { isIP } from 'node:net'

export type ServerConfig = {
  host: string
  openaiApiKey: string | null
  openaiSlideClassificationModel: string | null
  openaiSlideEmbeddingDimensions: number | undefined
  openaiSlideEmbeddingModel: string | null
  maxExportJsonBytes: number
  maxPreviewBytes: number
  maxUploadBytes: number
  port: number
  quarryPreviewRenderUrl: string
  previewRenderSize: number
  previewTimeoutMs: number
  requestTimeoutMs: number
  slideClassificationMaxImageBytes: number
  slideClassificationMaxTextChars: number
  slideClassificationProvider: 'disabled' | 'openai'
  slideClassificationTimeoutMs: number
  slideEmbeddingMaxTextBytes: number
  slideEmbeddingTimeoutMs: number
}

const DEFAULT_MAX_UPLOAD_BYTES = 25 * 1024 * 1024
const DEFAULT_MAX_EXPORT_JSON_BYTES = 50 * 1024 * 1024
const DEFAULT_PORT = 43127
const DEFAULT_REQUEST_TIMEOUT_MS = 90_000
const DEFAULT_PREVIEW_BYTES = 10 * 1024 * 1024
const DEFAULT_PREVIEW_RENDER_SIZE = 1600
const DEFAULT_QUARRY_PREVIEW_RENDER_URL = 'http://localhost:1420/_internal/template-preview'
const DEFAULT_PREVIEW_TIMEOUT_MS = 15_000
const DEFAULT_CLASSIFICATION_TIMEOUT_MS = 45_000
const DEFAULT_EMBEDDING_TIMEOUT_MS = 20_000
export function loadServerConfig(environment: NodeJS.ProcessEnv = process.env): ServerConfig {
  const slideClassificationProvider = parseClassificationProvider(
    environment.SLIDE_CLASSIFICATION_PROVIDER,
  )
  const openaiApiKey = nonBlank(environment.OPENAI_API_KEY)
  const openaiSlideClassificationModel = nonBlank(environment.OPENAI_SLIDE_CLASSIFICATION_MODEL)
  const openaiSlideEmbeddingModel = nonBlank(environment.OPENAI_SLIDE_EMBEDDING_MODEL)
  if (slideClassificationProvider === 'openai' && !openaiApiKey) {
    throw new Error('OPENAI_API_KEY is required when SLIDE_CLASSIFICATION_PROVIDER is openai.')
  }
  if (slideClassificationProvider === 'openai' && !openaiSlideClassificationModel) {
    throw new Error('OPENAI_SLIDE_CLASSIFICATION_MODEL is required when SLIDE_CLASSIFICATION_PROVIDER is openai.')
  }
  if (slideClassificationProvider === 'openai' && !openaiSlideEmbeddingModel) {
    throw new Error('OPENAI_SLIDE_EMBEDDING_MODEL is required when SLIDE_CLASSIFICATION_PROVIDER is openai.')
  }

  return {
    host: parseHost(environment.HOST),
    openaiApiKey,
    openaiSlideClassificationModel,
    openaiSlideEmbeddingDimensions: parseOptionalPositiveInteger(
      environment.OPENAI_SLIDE_EMBEDDING_DIMENSIONS,
      'OPENAI_SLIDE_EMBEDDING_DIMENSIONS',
    ),
    openaiSlideEmbeddingModel,
    maxExportJsonBytes: parsePositiveInteger(
      environment.MAX_EXPORT_JSON_BYTES,
      DEFAULT_MAX_EXPORT_JSON_BYTES,
      'MAX_EXPORT_JSON_BYTES',
    ),
    maxPreviewBytes: parsePositiveInteger(
      environment.MAX_TEMPLATE_PREVIEW_BYTES,
      DEFAULT_PREVIEW_BYTES,
      'MAX_TEMPLATE_PREVIEW_BYTES',
    ),
    maxUploadBytes: parsePositiveInteger(
      environment.MAX_PPTX_UPLOAD_BYTES,
      DEFAULT_MAX_UPLOAD_BYTES,
      'MAX_PPTX_UPLOAD_BYTES',
    ),
    port: parsePort(environment.PORT),
    quarryPreviewRenderUrl: parsePreviewRenderUrl(
      environment.QUARRY_TEMPLATE_PREVIEW_RENDER_URL,
      DEFAULT_QUARRY_PREVIEW_RENDER_URL,
      'QUARRY_TEMPLATE_PREVIEW_RENDER_URL',
    ),
    previewRenderSize: parsePreviewRenderSize(environment.TEMPLATE_PREVIEW_RENDER_SIZE),
    previewTimeoutMs: parsePositiveInteger(
      environment.TEMPLATE_PREVIEW_TIMEOUT_MS,
      DEFAULT_PREVIEW_TIMEOUT_MS,
      'TEMPLATE_PREVIEW_TIMEOUT_MS',
    ),
    requestTimeoutMs: parsePositiveInteger(
      environment.REQUEST_TIMEOUT_MS,
      DEFAULT_REQUEST_TIMEOUT_MS,
      'REQUEST_TIMEOUT_MS',
    ),
    slideClassificationMaxImageBytes: parsePositiveInteger(
      environment.SLIDE_CLASSIFICATION_MAX_IMAGE_BYTES,
      5 * 1024 * 1024,
      'SLIDE_CLASSIFICATION_MAX_IMAGE_BYTES',
    ),
    slideClassificationMaxTextChars: parsePositiveInteger(
      environment.SLIDE_CLASSIFICATION_MAX_TEXT_CHARS,
      12_000,
      'SLIDE_CLASSIFICATION_MAX_TEXT_CHARS',
    ),
    slideClassificationProvider,
    slideClassificationTimeoutMs: parsePositiveInteger(
      environment.SLIDE_CLASSIFICATION_TIMEOUT_MS,
      DEFAULT_CLASSIFICATION_TIMEOUT_MS,
      'SLIDE_CLASSIFICATION_TIMEOUT_MS',
    ),
    slideEmbeddingMaxTextBytes: parsePositiveInteger(
      environment.SLIDE_EMBEDDING_MAX_TEXT_BYTES,
      32_000,
      'SLIDE_EMBEDDING_MAX_TEXT_BYTES',
    ),
    slideEmbeddingTimeoutMs: parsePositiveInteger(
      environment.SLIDE_EMBEDDING_TIMEOUT_MS,
      DEFAULT_EMBEDDING_TIMEOUT_MS,
      'SLIDE_EMBEDDING_TIMEOUT_MS',
    ),
  }
}

function parseClassificationProvider(value: string | undefined): 'disabled' | 'openai' {
  const provider = value ?? 'disabled'
  if (provider !== 'disabled' && provider !== 'openai') {
    throw new Error('SLIDE_CLASSIFICATION_PROVIDER must be disabled or openai.')
  }
  return provider
}

function nonBlank(value: string | undefined) {
  const normalized = value?.trim()
  return normalized ? normalized : null
}

function parsePreviewRenderUrl(
  value: string | undefined,
  fallback: string,
  name: string,
) {
  const rawUrl = value ?? fallback
  let url: URL
  try {
    url = new URL(rawUrl)
  } catch {
    throw new Error(`${name} must be a valid HTTP or HTTPS URL.`)
  }
  if (
    (url.protocol !== 'http:' && url.protocol !== 'https:')
    || url.username
    || url.password
    || url.hash
  ) {
    throw new Error(`${name} must be a valid HTTP or HTTPS URL.`)
  }
  return url.toString()
}

function parsePreviewRenderSize(value: string | undefined) {
  const size = parsePositiveInteger(
    value,
    DEFAULT_PREVIEW_RENDER_SIZE,
    'TEMPLATE_PREVIEW_RENDER_SIZE',
  )
  if (size < 320 || size > 4096) {
    throw new Error('TEMPLATE_PREVIEW_RENDER_SIZE must be between 320 and 4096 pixels.')
  }
  return size
}

function parsePort(value: string | undefined) {
  if (value === undefined) {
    return DEFAULT_PORT
  }
  const port = parseUnsignedInteger(value)
  if (!Number.isSafeInteger(port) || port < 1 || port > 65_535) {
    throw new Error('PORT must be between 1 and 65535.')
  }
  return port
}

function parseHost(value: string | undefined) {
  const host = value ?? '0.0.0.0'
  if (isIP(host) === 0) {
    throw new Error('HOST must be an IP address.')
  }
  return host
}

function parsePositiveInteger(value: string | undefined, fallback: number, name: string) {
  if (value === undefined) {
    return fallback
  }

  const parsed = parseUnsignedInteger(value)
  if (!Number.isSafeInteger(parsed) || parsed < 1) {
    throw new Error(`${name} must be a positive whole number.`)
  }
  return parsed
}

function parseOptionalPositiveInteger(value: string | undefined, name: string) {
  if (value === undefined) return undefined
  return parsePositiveInteger(value, 1, name)
}

function parseUnsignedInteger(value: string) {
  return /^\+?\d+$/.test(value) ? Number(value) : Number.NaN
}
