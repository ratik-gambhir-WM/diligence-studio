import type { JsonObject, JsonValue } from '../canvas-model/CanvasTypes'

export type PickerTemplateKind = 'commentary' | 'diagram'

export type PickerTemplateSummary = {
  description: string
  elementCount: number
  kind: PickerTemplateKind
  previewUrl: string | null
  slideCount: number
  templateId: string
  title: string
}

export type PickerTemplateCatalog = {
  templates: PickerTemplateSummary[]
}

export type CanvasDocument = JsonObject

export type ImportedTemplate = {
  document: CanvasDocument
  previewStatus: 'ready' | 'unavailable'
  templateId: string
}

export type BatchImportedTemplate = Omit<ImportedTemplate, 'document'>

export type BatchImportedTemplates = {
  templates: [BatchImportedTemplate, ...BatchImportedTemplate[]]
  warnings: string[]
}

export type PowerPointDownload = {
  bytes: Uint8Array
  fileName: string
}

export class TemplateApiError extends Error {
  constructor(readonly code: string, message: string) {
    super(message)
  }
}

const POWERPOINT_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.presentation'
const API_BASE = (import.meta.env.VITE_API_BASE_URL || '/api/v1').replace(/\/$/u, '')
const APP_ID = import.meta.env.VITE_APP_ID?.trim() || 'DiligenceStudio_WestMonroe'
const APP_ID_HEADERS = { 'X-App-Id': APP_ID } as const

export async function listTemplates(
  kind: PickerTemplateKind,
  signal?: AbortSignal,
): Promise<PickerTemplateCatalog> {
  const response = await fetch(apiUrl(`/templates?kind=${encodeURIComponent(kind)}`), {
    headers: APP_ID_HEADERS,
    signal,
  })
  await assertOk(response)
  assertJsonContentType(response)
  return parseCatalog(await response.json())
}

export async function getTemplate(templateId: string, signal?: AbortSignal) {
  const response = await fetch(apiUrl(`/templates/${encodeURIComponent(templateId)}`), {
    headers: APP_ID_HEADERS,
    signal,
  })
  await assertOk(response)
  assertJsonContentType(response)
  return parseCanvasDocument(await response.json())
}

export async function deleteTemplate(templateId: string, signal?: AbortSignal) {
  const response = await fetch(apiUrl(`/templates/${encodeURIComponent(templateId)}`), {
    headers: APP_ID_HEADERS,
    method: 'DELETE',
    signal,
  })
  await assertOk(response)
  if (response.status !== 204) {
    throw new TemplateApiError('invalid_api_response', 'The template service returned an invalid delete response.')
  }
}

export async function importTemplate(
  file: File,
  kind: PickerTemplateKind,
  signal?: AbortSignal,
): Promise<ImportedTemplate> {
  const response = await fetch(apiUrl(`/import?kind=${encodeURIComponent(kind)}`), {
    body: file,
    headers: { ...APP_ID_HEADERS, 'Content-Type': POWERPOINT_CONTENT_TYPE },
    method: 'POST',
    signal,
  })
  await assertOk(response)
  assertJsonContentType(response)
  const templateId = response.headers.get('X-Template-Id')
  const previewStatus = response.headers.get('X-Template-Preview-Status')
  if (!templateId || (previewStatus !== 'ready' && previewStatus !== 'unavailable')) {
    throw new TemplateApiError('invalid_api_response', 'The template import response was incomplete.')
  }
  return {
    document: parseCanvasDocument(await response.json()),
    previewStatus,
    templateId,
  }
}

export async function batchImportTemplates(
  file: File,
  kind: PickerTemplateKind,
  signal?: AbortSignal,
): Promise<BatchImportedTemplates> {
  const response = await fetch(apiUrl(`/batchImport?kind=${encodeURIComponent(kind)}`), {
    body: file,
    headers: { ...APP_ID_HEADERS, 'Content-Type': POWERPOINT_CONTENT_TYPE },
    method: 'POST',
    signal,
  })
  await assertOk(response)
  assertJsonContentType(response)
  return parseBatchImport(await response.json())
}

export async function exportPresentation(input: JsonValue, signal?: AbortSignal) {
  const response = await fetch(apiUrl('/export'), {
    body: JSON.stringify(input),
    headers: { ...APP_ID_HEADERS, 'Content-Type': 'application/json' },
    method: 'POST',
    signal,
  })
  return parsePowerPointDownload(response)
}

export async function insertPresentation(
  input: JsonValue,
  targetFile: File,
  insertAfterSlide: number,
  signal?: AbortSignal,
) {
  const form = new FormData()
  form.set('presentation', JSON.stringify(input))
  form.set('insertAfterSlide', String(insertAfterSlide))
  form.set('target', targetFile, targetFile.name)
  const response = await fetch(apiUrl('/export/insert'), {
    body: form,
    headers: APP_ID_HEADERS,
    method: 'POST',
    signal,
  })
  return parsePowerPointDownload(response)
}

export function resolvePreviewUrl(path: string | null) {
  return path ? apiUrl(path) : null
}

function apiUrl(path: string) {
  return `${API_BASE}${path.startsWith('/') ? path : `/${path}`}`
}

async function parsePowerPointDownload(response: Response): Promise<PowerPointDownload> {
  await assertOk(response)
  const contentType = response.headers.get('Content-Type')?.split(';', 1)[0]?.trim()
  if (contentType !== POWERPOINT_CONTENT_TYPE) {
    throw new TemplateApiError('invalid_api_response', 'The server did not return a PowerPoint file.')
  }
  const disposition = response.headers.get('Content-Disposition') ?? ''
  const fileName = disposition.match(/filename="([^"]+)"/u)?.[1] ?? 'presentation.pptx'
  return { bytes: new Uint8Array(await response.arrayBuffer()), fileName }
}

async function assertOk(response: Response) {
  if (response.ok) {
    return
  }
  let code = 'request_failed'
  let message = `The server returned ${response.status}. Try again.`
  if (response.headers.get('Content-Type')?.includes('application/json')) {
    try {
      const payload = await response.json() as unknown
      if (isRecord(payload) && isRecord(payload.error)) {
        if (typeof payload.error.code === 'string') code = payload.error.code
        if (typeof payload.error.message === 'string') message = payload.error.message
      }
    } catch {
      // Keep the sanitized fallback.
    }
  }
  throw new TemplateApiError(code, message)
}

function assertJsonContentType(response: Response) {
  if (!response.headers.get('Content-Type')?.includes('application/json')) {
    throw new TemplateApiError('invalid_api_response', 'The server returned an unexpected response format.')
  }
}

function parseCatalog(value: unknown): PickerTemplateCatalog {
  if (!isRecord(value) || !Array.isArray(value.templates)) {
    throw invalidResponse()
  }
  return { templates: value.templates.map(parseSummary) }
}

function parseBatchImport(value: unknown): BatchImportedTemplates {
  if (
    !isRecord(value)
    || !Array.isArray(value.templates)
    || value.templates.length === 0
    || !Array.isArray(value.warnings)
    || !value.warnings.every((warning) => typeof warning === 'string')
  ) {
    throw invalidResponse()
  }

  const templates = value.templates.map((template): BatchImportedTemplate => {
    if (
      !isRecord(template)
      || !isNonEmptyString(template.templateId)
      || typeof template.previewAvailable !== 'boolean'
    ) {
      throw invalidResponse()
    }
    parseCanvasDocument(template.templateJson)
    return {
      previewStatus: template.previewAvailable ? 'ready' : 'unavailable',
      templateId: template.templateId,
    }
  })
  const [firstTemplate, ...remainingTemplates] = templates
  if (!firstTemplate) {
    throw invalidResponse()
  }

  return {
    templates: [firstTemplate, ...remainingTemplates],
    warnings: [...value.warnings],
  }
}

function parseSummary(value: unknown): PickerTemplateSummary {
  if (
    !isRecord(value)
    || !isNonEmptyString(value.templateId)
    || !isTemplateKind(value.kind)
    || !isNonEmptyString(value.title)
    || typeof value.description !== 'string'
    || !isNonNegativeInteger(value.slideCount)
    || !isNonNegativeInteger(value.elementCount)
    || (value.previewUrl !== null && typeof value.previewUrl !== 'string')
  ) {
    throw invalidResponse()
  }
  return {
    description: value.description,
    elementCount: value.elementCount,
    kind: value.kind,
    previewUrl: resolvePreviewUrl(value.previewUrl),
    slideCount: value.slideCount,
    templateId: value.templateId,
    title: value.title,
  }
}

function parseCanvasDocument(value: unknown): CanvasDocument {
  if (!isJsonValue(value) || !isRecord(value) || !isRecord(value.presentation)) {
    throw invalidResponse()
  }
  const presentation = value.presentation
  if (!Array.isArray(presentation.slides) || presentation.slides.length === 0) {
    throw invalidResponse()
  }
  return value
}

function isJsonValue(value: unknown, depth = 0): value is JsonValue {
  if (depth > 100) return false
  if (value === null || typeof value === 'string' || typeof value === 'boolean') return true
  if (typeof value === 'number') return Number.isFinite(value)
  if (Array.isArray(value)) return value.every((entry) => isJsonValue(entry, depth + 1))
  if (!isRecord(value)) return false
  return Object.values(value).every((entry) => entry === undefined || isJsonValue(entry, depth + 1))
}

function invalidResponse() {
  return new TemplateApiError('invalid_api_response', 'The template service returned invalid data.')
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function isTemplateKind(value: unknown): value is PickerTemplateKind {
  return value === 'diagram' || value === 'commentary'
}

function isNonEmptyString(value: unknown): value is string {
  return typeof value === 'string' && value.trim().length > 0
}

function isNonNegativeInteger(value: unknown): value is number {
  return typeof value === 'number' && Number.isSafeInteger(value) && value >= 0
}
