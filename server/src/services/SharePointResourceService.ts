import { Buffer } from 'node:buffer'

const GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0'
const MAX_FILES = 500
const MAX_FOLDER_DEPTH = 8
const MAX_FOLDER_ITEMS = 5_000
const MAX_FOLDER_REQUESTS = 1_000
const PAGE_SIZE = 200
const DEFAULT_MAX_FILE_BYTES = 25 * 1024 * 1024
const MAX_ERROR_BODY_BYTES = 1_000
const SUPPORTED_EXTENSIONS = new Set([
  'docx',
  'pdf',
  'ppt',
  'pptx',
  'png',
  'jpg',
  'jpeg',
  'md',
  'markdown',
  'txt',
  'rtf',
])

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

export type SharePointDownloadedFile = {
  bytes: Buffer
  contentType: string
  fileName: string
}

export interface SharePointResourceServiceLike {
  resolveResource(
    resourceUrl: string,
    accessToken: string,
    signal?: AbortSignal,
  ): Promise<SharePointResolvedResource>
  downloadFile(
    driveId: string,
    fileId: string,
    accessToken: string,
    signal?: AbortSignal,
  ): Promise<SharePointDownloadedFile>
}

export type SharePointResourceServiceOptions = {
  allowedHosts?: readonly string[]
  maxFileBytes?: number
  maxFolderItems?: number
  maxFolderRequests?: number
}

export class SharePointResourceError extends Error {
  constructor(
    readonly statusCode: number,
    readonly code: string,
    message: string,
  ) {
    super(message)
    this.name = 'SharePointResourceError'
  }
}

export class SharePointResourceService implements SharePointResourceServiceLike {
  private readonly allowedHosts: ReadonlySet<string>
  private readonly maxFileBytes: number
  private readonly maxFolderItems: number
  private readonly maxFolderRequests: number

  constructor(options: SharePointResourceServiceOptions = {}) {
    this.allowedHosts = new Set(
      (options.allowedHosts ?? ['relentlessblue.sharepoint.com'])
        .map((host) => host.trim().toLowerCase())
        .filter(Boolean),
    )
    this.maxFileBytes = options.maxFileBytes ?? DEFAULT_MAX_FILE_BYTES
    this.maxFolderItems = options.maxFolderItems ?? MAX_FOLDER_ITEMS
    this.maxFolderRequests = options.maxFolderRequests ?? MAX_FOLDER_REQUESTS
  }

  async resolveResource(
    resourceUrl: string,
    accessToken: string,
    signal?: AbortSignal,
  ): Promise<SharePointResolvedResource> {
    const validatedUrl = this.validateResourceUrl(resourceUrl)
    const sharedItem = await this.fetchSharedDriveItem(validatedUrl, accessToken, signal)
    const item = parseDriveItem(sharedItem)
    const driveId = item.parentReference?.driveId

    if (!driveId) {
      throw new SharePointResourceError(
        422,
        'sharepoint_missing_drive',
        'The SharePoint resource did not include a document-library drive.',
      )
    }

    if (item.folder) {
      const files = await this.listFolderFiles(
        driveId,
        item.id,
        item.name,
        accessToken,
        signal,
        { itemCount: 0, requestCount: 0 },
      )
      return {
        files,
        kind: 'folder',
        name: item.name,
        resourceUrl: validatedUrl.toString(),
        webUrl: item.webUrl,
      }
    }

    if (!item.file) {
      throw new SharePointResourceError(
        422,
        'sharepoint_unsupported_resource',
        'The SharePoint link does not point to a file or folder.',
      )
    }

    if (!isSupportedFileName(item.name)) {
      throw new SharePointResourceError(
        415,
        'sharepoint_unsupported_file_type',
        'The SharePoint file type is not supported by this application.',
      )
    }

    return {
      files: [toSharePointFile(item, driveId, '')],
      kind: 'file',
      name: item.name,
      resourceUrl: validatedUrl.toString(),
      webUrl: item.webUrl,
    }
  }

  async downloadFile(
    driveId: string,
    fileId: string,
    accessToken: string,
    signal?: AbortSignal,
  ): Promise<SharePointDownloadedFile> {
    validateGraphIdentifier(driveId, 'driveId')
    validateGraphIdentifier(fileId, 'fileId')

    const metadata = parseDriveItem(await this.graphJson(
      `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(fileId)}?$select=id,name,size,file,parentReference,webUrl,lastModifiedDateTime`,
      accessToken,
      signal,
    ))
    if (!metadata.file) {
      throw new SharePointResourceError(
        422,
        'sharepoint_not_a_file',
        'The selected SharePoint item is not a file.',
      )
    }

    if (metadata.size > this.maxFileBytes) {
      throw new SharePointResourceError(
        413,
        'sharepoint_file_too_large',
        `The SharePoint file exceeds the ${this.maxFileBytes} byte download limit.`,
      )
    }

    const response = await this.graphFetch(
      `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(fileId)}/content`,
      accessToken,
      signal,
    )
    const bytes = await readResponseBytes(response, this.maxFileBytes)
    const contentType = metadata.file.mimeType || response.headers.get('content-type') || 'application/octet-stream'

    return {
      bytes,
      contentType,
      fileName: metadata.name,
    }
  }

  private validateResourceUrl(resourceUrl: string) {
    let url: URL
    try {
      url = new URL(resourceUrl)
    } catch {
      throw new SharePointResourceError(
        400,
        'invalid_sharepoint_url',
        'Enter a valid SharePoint file or folder link.',
      )
    }

    if (url.protocol !== 'https:' || !this.allowedHosts.has(url.hostname.toLowerCase())) {
      throw new SharePointResourceError(
        400,
        'unsupported_sharepoint_host',
        'The SharePoint link must belong to an approved SharePoint site.',
      )
    }

    return url
  }

  private async fetchSharedDriveItem(
    resourceUrl: URL,
    accessToken: string,
    signal?: AbortSignal,
  ) {
    const encodedSharingUrl = encodeSharingUrl(resourceUrl.toString())
    return this.graphJson(
      `/shares/${encodeURIComponent(encodedSharingUrl)}/driveItem?$select=id,name,size,file,folder,parentReference,webUrl,lastModifiedDateTime`,
      accessToken,
      signal,
    )
  }

  private async listFolderFiles(
    driveId: string,
    folderId: string,
    folderPath: string,
    accessToken: string,
    signal: AbortSignal | undefined,
    traversal: FolderTraversalState,
    depth = 0,
  ): Promise<SharePointFile[]> {
    if (depth > MAX_FOLDER_DEPTH) {
      throw new SharePointResourceError(
        413,
        'sharepoint_folder_too_deep',
        'The SharePoint folder is too deeply nested to scan.',
      )
    }

    const files: SharePointFile[] = []
    let nextUrl = `${GRAPH_BASE_URL}/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(folderId)}/children?$select=id,name,size,file,folder,parentReference,webUrl,lastModifiedDateTime&$top=${PAGE_SIZE}`

    while (nextUrl) {
      traversal.requestCount += 1
      if (traversal.requestCount > this.maxFolderRequests) {
        throw new SharePointResourceError(
          413,
          'sharepoint_folder_request_limit_exceeded',
          'The SharePoint folder requires too many requests to scan.',
        )
      }

      const data = await this.graphJsonUrl(nextUrl, accessToken, signal)
      const items = getArrayProperty(data, 'value')

      for (const value of items) {
        traversal.itemCount += 1
        if (traversal.itemCount > this.maxFolderItems) {
          throw new SharePointResourceError(
            413,
            'sharepoint_folder_item_limit_exceeded',
            'The SharePoint folder contains too many items to scan.',
          )
        }

        const item = parseDriveItem(value)
        if (item.folder) {
          files.push(...await this.listFolderFiles(
            driveId,
            item.id,
            joinPath(folderPath, item.name),
            accessToken,
            signal,
            traversal,
            depth + 1,
          ))
        } else if (item.file && isSupportedFileName(item.name)) {
          files.push(toSharePointFile(item, driveId, folderPath))
        }

        if (files.length > MAX_FILES) {
          throw new SharePointResourceError(
            413,
            'sharepoint_file_limit_exceeded',
            `The SharePoint folder contains more than ${MAX_FILES} files.`,
          )
        }
      }

      const nextLink = getStringProperty(data, '@odata.nextLink')
      if (!nextLink || !nextLink.startsWith(`${GRAPH_BASE_URL}/`)) {
        nextUrl = ''
      } else {
        nextUrl = nextLink
      }
    }

    return files
  }

  private async graphJson(path: string, accessToken: string, signal?: AbortSignal) {
    const response = await this.graphFetch(path, accessToken, signal)
    return response.json() as Promise<unknown>
  }

  private async graphJsonUrl(url: string, accessToken: string, signal?: AbortSignal) {
    const response = await this.graphFetch(url, accessToken, signal)
    return response.json() as Promise<unknown>
  }

  private async graphFetch(pathOrUrl: string, accessToken: string, signal?: AbortSignal) {
    const url = pathOrUrl.startsWith('https://') ? pathOrUrl : `${GRAPH_BASE_URL}${pathOrUrl}`
    const response = await fetch(url, {
      headers: {
        Accept: 'application/json',
        Authorization: `Bearer ${accessToken.replace(/^Bearer\s+/iu, '')}`,
      },
      redirect: 'follow',
      signal,
    })

    if (response.ok) {
      return response
    }

    await readResponseText(response, MAX_ERROR_BODY_BYTES)
    throw new SharePointResourceError(
      response.status === 401 || response.status === 403 ? response.status : 502,
      response.status === 401 || response.status === 403
        ? 'sharepoint_access_denied'
        : 'sharepoint_request_failed',
      response.status === 401 || response.status === 403
        ? 'Microsoft Graph denied access to this SharePoint resource.'
        : 'Microsoft Graph could not retrieve the SharePoint resource.',
    )
  }
}

type FolderTraversalState = {
  itemCount: number
  requestCount: number
}

async function readResponseBytes(response: Response, maxBytes: number) {
  const contentLength = response.headers.get('content-length')
  if (contentLength && /^\d+$/u.test(contentLength) && Number(contentLength) > maxBytes) {
    await response.body?.cancel().catch(() => undefined)
    throw new SharePointResourceError(
      413,
      'sharepoint_file_too_large',
      `The SharePoint file exceeds the ${maxBytes} byte download limit.`,
    )
  }

  const reader = response.body?.getReader()
  if (!reader) {
    const bytes = Buffer.from(await response.arrayBuffer())
    if (bytes.length > maxBytes) {
      throw new SharePointResourceError(
        413,
        'sharepoint_file_too_large',
        `The SharePoint file exceeds the ${maxBytes} byte download limit.`,
      )
    }
    return bytes
  }

  const chunks: Buffer[] = []
  let totalBytes = 0
  try {
    while (true) {
      const { done, value } = await reader.read()
      if (done) break

      totalBytes += value.byteLength
      if (totalBytes > maxBytes) {
        await reader.cancel().catch(() => undefined)
        throw new SharePointResourceError(
          413,
          'sharepoint_file_too_large',
          `The SharePoint file exceeds the ${maxBytes} byte download limit.`,
        )
      }
      chunks.push(Buffer.from(value))
    }
  } finally {
    reader.releaseLock()
  }

  return Buffer.concat(chunks, totalBytes)
}

async function readResponseText(response: Response, maxBytes: number) {
  const reader = response.body?.getReader()
  if (!reader) {
    return (await response.text()).slice(0, maxBytes)
  }

  const decoder = new TextDecoder()
  let totalBytes = 0
  let text = ''
  try {
    while (totalBytes < maxBytes) {
      const { done, value } = await reader.read()
      if (done) break

      const remaining = maxBytes - totalBytes
      const chunk = value.byteLength > remaining ? value.subarray(0, remaining) : value
      totalBytes += chunk.byteLength
      text += decoder.decode(chunk, { stream: totalBytes < maxBytes })
      if (chunk.byteLength < value.byteLength) {
        await reader.cancel().catch(() => undefined)
        break
      }
    }
  } finally {
    reader.releaseLock()
  }
  return text + decoder.decode()
}

function encodeSharingUrl(url: string) {
  return `u!${Buffer.from(url, 'utf8')
    .toString('base64')
    .replace(/=+$/u, '')
    .replace(/\//gu, '_')
    .replace(/\+/gu, '-')}`
}

function parseDriveItem(value: unknown) {
  const source = isRecord(value) && isRecord(value.driveItem) ? value.driveItem : value
  if (!isRecord(source)) {
    throw new SharePointResourceError(502, 'invalid_sharepoint_response', 'Microsoft Graph returned an invalid SharePoint item.')
  }

  const id = getRequiredString(source, 'id')
  const name = getRequiredString(source, 'name')
  const webUrl = getRequiredString(source, 'webUrl')
  const parentReference = isRecord(source.parentReference) ? source.parentReference : undefined
  const driveId = parentReference ? getStringProperty(parentReference, 'driveId') : undefined
  const file = isRecord(source.file)
    ? { mimeType: getStringProperty(source.file, 'mimeType') ?? 'application/octet-stream' }
    : undefined

  return {
    file,
    folder: source.folder !== undefined && source.folder !== null,
    id,
    lastModifiedDateTime: getStringProperty(source, 'lastModifiedDateTime') ?? '',
    name,
    parentReference: driveId ? { driveId } : undefined,
    size: getNonNegativeInteger(source, 'size'),
    webUrl,
  }
}

function toSharePointFile(
  item: ReturnType<typeof parseDriveItem>,
  driveId: string,
  path: string,
): SharePointFile {
  return {
    driveId,
    fileId: item.id,
    lastModifiedDateTime: item.lastModifiedDateTime,
    mimeType: item.file?.mimeType ?? 'application/octet-stream',
    name: item.name,
    path,
    size: item.size,
    webUrl: item.webUrl,
  }
}

function joinPath(parent: string, child: string) {
  return parent ? `${parent}/${child}` : child
}

function isSupportedFileName(name: string) {
  const extension = name.slice(name.lastIndexOf('.') + 1).toLowerCase()
  return extension !== name.toLowerCase() && SUPPORTED_EXTENSIONS.has(extension)
}

function validateGraphIdentifier(value: string, name: string) {
  if (!value || value.length > 512 || /[\r\n]/u.test(value)) {
    throw new SharePointResourceError(400, 'invalid_sharepoint_identifier', `The SharePoint ${name} is invalid.`)
  }
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function getArrayProperty(value: unknown, property: string) {
  if (!isRecord(value) || !Array.isArray(value[property])) {
    throw new SharePointResourceError(502, 'invalid_sharepoint_response', 'Microsoft Graph returned an invalid file list.')
  }
  return value[property]
}

function getRequiredString(value: Record<string, unknown>, property: string) {
  const result = getStringProperty(value, property)
  if (!result) {
    throw new SharePointResourceError(502, 'invalid_sharepoint_response', 'Microsoft Graph returned an incomplete SharePoint item.')
  }
  return result
}

function getStringProperty(value: unknown, property: string) {
  return isRecord(value) && typeof value[property] === 'string' ? value[property] : undefined
}

function getNonNegativeInteger(value: Record<string, unknown>, property: string) {
  const result = value[property]
  return typeof result === 'number' && Number.isSafeInteger(result) && result >= 0 ? result : 0
}
