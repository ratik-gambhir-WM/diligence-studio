export type SharePointFile = {
  driveId: string
  fileId: string
  lastModifiedDateTime: string
  mimeType: string
  name: string
  path: string
  size: number
  webUrl: string
}

export type SharePointResolvedResource = {
  files: SharePointFile[]
  kind: 'file' | 'folder'
  name: string
  resourceUrl: string
  webUrl: string
}

export class SharePointApiError extends Error {
  constructor(readonly code: string, message: string) {
    super(message)
    this.name = 'SharePointApiError'
  }
}

const API_BASE = (import.meta.env.VITE_API_BASE_URL || '/api/v1').replace(/\/$/u, '')

export async function resolveSharePointResource(
  url: string,
  signal?: AbortSignal,
): Promise<SharePointResolvedResource> {
  const response = await fetch(`${API_BASE}/sharepoint/resolve`, {
    body: JSON.stringify({ url }),
    credentials: 'include',
    headers: { 'Content-Type': 'application/json' },
    method: 'POST',
    signal,
  })
  await assertOk(response)
  assertJsonContentType(response)
  return parseResolvedResource(await response.json())
}

export async function downloadSharePointFile(
  file: SharePointFile,
  signal?: AbortSignal,
): Promise<File> {
  const response = await fetch(
    `${API_BASE}/sharepoint/files/${encodeURIComponent(file.driveId)}/${encodeURIComponent(file.fileId)}/content`,
    { credentials: 'include', signal },
  )
  await assertOk(response)

  const contentType = response.headers.get('Content-Type')?.split(';', 1)[0]?.trim() || file.mimeType
  const disposition = response.headers.get('Content-Disposition') ?? ''
  const fileName = disposition.match(/filename="([^"]+)"/u)?.[1] || file.name
  return new File([await response.arrayBuffer()], fileName, { type: contentType })
}

async function assertOk(response: Response) {
  if (response.ok) return

  let code = 'sharepoint_request_failed'
  let message = `SharePoint returned ${response.status}. Try again.`
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
  throw new SharePointApiError(code, message)
}

function assertJsonContentType(response: Response) {
  if (!response.headers.get('Content-Type')?.includes('application/json')) {
    throw new SharePointApiError(
      'invalid_sharepoint_response',
      'SharePoint returned an unexpected response format.',
    )
  }
}

function parseResolvedResource(value: unknown): SharePointResolvedResource {
  if (
    !isRecord(value)
    || (value.kind !== 'file' && value.kind !== 'folder')
    || !isNonEmptyString(value.name)
    || !isNonEmptyString(value.resourceUrl)
    || !isNonEmptyString(value.webUrl)
    || !Array.isArray(value.files)
  ) {
    throw new SharePointApiError(
      'invalid_sharepoint_response',
      'SharePoint returned an incomplete resource response.',
    )
  }

  return {
    files: value.files.map(parseSharePointFile),
    kind: value.kind,
    name: value.name,
    resourceUrl: value.resourceUrl,
    webUrl: value.webUrl,
  }
}

function parseSharePointFile(value: unknown): SharePointFile {
  if (
    !isRecord(value)
    || !isNonEmptyString(value.driveId)
    || !isNonEmptyString(value.fileId)
    || !isNonEmptyString(value.name)
    || typeof value.path !== 'string'
    || !isNonNegativeInteger(value.size)
    || !isNonEmptyString(value.mimeType)
    || typeof value.lastModifiedDateTime !== 'string'
    || !isNonEmptyString(value.webUrl)
  ) {
    throw new SharePointApiError(
      'invalid_sharepoint_response',
      'SharePoint returned an invalid file record.',
    )
  }

  return {
    driveId: value.driveId,
    fileId: value.fileId,
    lastModifiedDateTime: value.lastModifiedDateTime,
    mimeType: value.mimeType,
    name: value.name,
    path: value.path,
    size: value.size,
    webUrl: value.webUrl,
  }
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function isNonEmptyString(value: unknown): value is string {
  return typeof value === 'string' && value.trim().length > 0
}

function isNonNegativeInteger(value: unknown): value is number {
  return typeof value === 'number' && Number.isSafeInteger(value) && value >= 0
}
