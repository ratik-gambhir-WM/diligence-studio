import { isIP } from 'node:net'
import { existsSync } from 'node:fs'
import path from 'node:path'
import { fileURLToPath } from 'node:url'

export type ServerConfig = {
  databasePath: string
  host: string
  maxExportJsonBytes: number
  maxPreviewBytes: number
  maxUploadBytes: number
  microsoftCookieSecure: boolean
  microsoftSupport: boolean
  port: number
  previewProvider: 'disabled' | 'headless' | 'quicklook'
  previewRenderSize: number
  previewRenderUrl: string
  previewTimeoutMs: number
  requestTimeoutMs: number
  sharePointAllowedHosts: string[]
}

const DEFAULT_MAX_UPLOAD_BYTES = 25 * 1024 * 1024
const DEFAULT_MAX_EXPORT_JSON_BYTES = 50 * 1024 * 1024
const DEFAULT_PORT = 43127
const DEFAULT_REQUEST_TIMEOUT_MS = 30_000
const DEFAULT_PREVIEW_BYTES = 10 * 1024 * 1024
const DEFAULT_PREVIEW_RENDER_SIZE = 1600
const DEFAULT_PREVIEW_RENDER_URL = 'http://localhost:5173/_internal/template-preview'
const DEFAULT_PREVIEW_TIMEOUT_MS = 15_000
const QUICK_LOOK_PATH = '/usr/bin/qlmanage'
const DEFAULT_DATABASE_PATH = fileURLToPath(new URL('../data/templates.sqlite', import.meta.url))

export function loadServerConfig(environment: NodeJS.ProcessEnv = process.env): ServerConfig {
  return {
    databasePath: environment.SQLITE_DB_PATH
      ? path.resolve(environment.SQLITE_DB_PATH)
      : DEFAULT_DATABASE_PATH,
    host: parseHost(environment.HOST),
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
    microsoftCookieSecure: parseBoolean(
      environment.MICROSOFT_COOKIE_SECURE,
      true,
      'MICROSOFT_COOKIE_SECURE',
    ),
    microsoftSupport: parseBoolean(
      environment.MICROSOFT_SUPPORT,
      false,
      'MICROSOFT_SUPPORT',
    ),
    port: parsePort(environment.PORT),
    previewProvider: parsePreviewProvider(environment.TEMPLATE_PREVIEW_PROVIDER),
    previewRenderSize: parsePreviewRenderSize(environment.TEMPLATE_PREVIEW_RENDER_SIZE),
    previewRenderUrl: parsePreviewRenderUrl(environment.TEMPLATE_PREVIEW_RENDER_URL),
    previewTimeoutMs: parsePositiveInteger(
      environment.TEMPLATE_PREVIEW_TIMEOUT_MS,
      DEFAULT_PREVIEW_TIMEOUT_MS,
      'TEMPLATE_PREVIEW_TIMEOUT_MS',
    ),
    requestTimeoutMs: DEFAULT_REQUEST_TIMEOUT_MS,
    sharePointAllowedHosts: parseAllowedHosts(environment.SHAREPOINT_ALLOWED_HOSTS),
  }
}

function parseAllowedHosts(value: string | undefined) {
  const hosts = (value ?? 'relentlessblue.sharepoint.com')
    .split(',')
    .map((host) => host.trim().toLowerCase())
    .filter(Boolean)

  if (hosts.length === 0 || hosts.some((host) => !/^[a-z0-9.-]+$/u.test(host))) {
    throw new Error('SHAREPOINT_ALLOWED_HOSTS must contain comma-separated hostnames.')
  }

  return [...new Set(hosts)]
}

function parseBoolean(value: string | undefined, fallback: boolean, name: string) {
  if (value === undefined) {
    return fallback
  }

  const normalized = value.trim().toLowerCase()
  if (normalized === 'true') {
    return true
  }
  if (normalized === 'false') {
    return false
  }

  throw new Error(`${name} must be true or false.`)
}

function parsePreviewProvider(value: string | undefined): 'disabled' | 'headless' | 'quicklook' {
  const provider = value ?? 'headless'
  if (provider !== 'disabled' && provider !== 'headless' && provider !== 'quicklook') {
    throw new Error('TEMPLATE_PREVIEW_PROVIDER must be headless, quicklook, or disabled.')
  }
  if (provider === 'quicklook' && !existsSync(QUICK_LOOK_PATH)) {
    return 'disabled'
  }
  return provider
}

function parsePreviewRenderUrl(value: string | undefined) {
  const rawUrl = value ?? DEFAULT_PREVIEW_RENDER_URL
  let url: URL
  try {
    url = new URL(rawUrl)
  } catch {
    throw new Error('TEMPLATE_PREVIEW_RENDER_URL must be a valid HTTP or HTTPS URL.')
  }
  if (
    (url.protocol !== 'http:' && url.protocol !== 'https:')
    || url.username
    || url.password
    || url.hash
  ) {
    throw new Error('TEMPLATE_PREVIEW_RENDER_URL must be a valid HTTP or HTTPS URL.')
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

function parseUnsignedInteger(value: string) {
  return /^\+?\d+$/.test(value) ? Number(value) : Number.NaN
}
