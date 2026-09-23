import { Router, json, type Request, type RequestHandler, type Response } from 'express'

import { AuthError, getMicrosoftAccessToken } from '../auth/microsoftAuth'
import { ApiError } from '../errors'
import {
  SharePointResourceError,
  type SharePointResourceServiceLike,
} from '../services/SharePointResourceService'

const MAX_RESOLVE_BODY_BYTES = 16 * 1024

export type MicrosoftAccessTokenResolver = (request: Request) => Promise<string>

export function createSharePointRouter(
  service: SharePointResourceServiceLike,
  resolveAccessToken: MicrosoftAccessTokenResolver = getMicrosoftAccessToken,
) {
  const router = Router()

  router.post(
    '/resolve',
    json({ limit: MAX_RESOLVE_BODY_BYTES, strict: true, type: 'application/json' }),
    createResolveHandler(service, resolveAccessToken),
  )
  router.get(
    '/files/:driveId/:fileId/content',
    createDownloadHandler(service, resolveAccessToken),
  )
  router.all('/resolve', methodNotAllowed)
  router.all('/files/:driveId/:fileId/content', methodNotAllowed)

  return router
}

function createResolveHandler(
  service: SharePointResourceServiceLike,
  resolveAccessToken: MicrosoftAccessTokenResolver,
): RequestHandler {
  return async (request, response) => {
    if (!request.is('application/json')) {
      throw new ApiError(415, 'unsupported_media_type', 'Content-Type must be application/json.')
    }

    const url = parseResourceUrl(request.body)
    const token = await requireAccessToken(request, resolveAccessToken)
    const result = await runSharePointOperation(
      () => service.resolveResource(url, token),
    )

    response
      .status(200)
      .set('Cache-Control', 'private, no-store')
      .json(result)
  }
}

function createDownloadHandler(
  service: SharePointResourceServiceLike,
  resolveAccessToken: MicrosoftAccessTokenResolver,
): RequestHandler {
  return async (request, response) => {
    const driveId = parseIdentifier(request.params.driveId, 'driveId')
    const fileId = parseIdentifier(request.params.fileId, 'fileId')
    const token = await requireAccessToken(request, resolveAccessToken)
    const result = await runSharePointOperation(
      () => service.downloadFile(driveId, fileId, token),
    )

    response
      .status(200)
      .set({
        'Cache-Control': 'private, no-store',
        'Content-Disposition': `attachment; filename="${escapeHeaderValue(result.fileName)}"`,
        'Content-Length': String(result.bytes.length),
        'Content-Type': result.contentType,
        'X-Content-Type-Options': 'nosniff',
      })
      .send(result.bytes)
  }
}

async function requireAccessToken(
  request: Request,
  resolveAccessToken: MicrosoftAccessTokenResolver,
) {
  try {
    return await resolveAccessToken(request)
  } catch (error) {
    if (error instanceof AuthError) {
      throw new ApiError(error.statusCode, 'authentication_required', error.message)
    }
    throw error
  }
}

async function runSharePointOperation<Result>(operation: () => Promise<Result>) {
  try {
    return await operation()
  } catch (error) {
    if (error instanceof SharePointResourceError) {
      throw new ApiError(error.statusCode, error.code, error.message)
    }
    throw error
  }
}

function parseResourceUrl(value: unknown) {
  if (!isRecord(value) || typeof value.url !== 'string' || value.url.trim().length === 0) {
    throw new ApiError(400, 'invalid_sharepoint_request', 'A SharePoint file or folder URL is required.')
  }
  return value.url.trim()
}

function parseIdentifier(value: unknown, name: string) {
  if (typeof value !== 'string' || value.length === 0 || value.length > 512 || /[\r\n]/u.test(value)) {
    throw new ApiError(400, 'invalid_sharepoint_identifier', `The SharePoint ${name} is invalid.`)
  }
  return value
}

function escapeHeaderValue(value: string) {
  return value.replace(/[\r\n"\\]/gu, '_').slice(0, 180) || 'sharepoint-file'
}

function methodNotAllowed(_request: Request, response: Response) {
  response.status(405).end()
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}
