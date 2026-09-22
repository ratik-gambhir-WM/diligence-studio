import type { Request, RequestHandler, Response } from 'express'

import { API_V1_PATH } from '../apiPaths'
import { buildAppScopedPath, parseAppId } from '../appIdentity'
import { ApiError } from '../errors'
import type { TemplateKind } from '../repositories/TemplateRepository'
import {
  TEMPLATE_PREVIEW_PAGE_SIZE,
  type ImportService,
  type TemplateV2ReadOptions,
} from '../services/ImportTemplateService'

export function createImportHandlers(service: ImportService) {
  const create: RequestHandler = async (request, response) => {
    validatePowerPointRequest(request)

    const appId = parseAppId(request)
    const kind = parseTemplateKind(request.query.kind, true)
    const result = await runCancellableImport(
      request,
      response,
      (signal) => service.import(request.body, kind, signal, appId),
    )
    if (response.writableEnded) {
      return
    }
    response.status(201).set({
      Location: buildAppScopedPath(`${API_V1_PATH}/templates/${result.templateId}`, appId),
      'X-PowerPoint-Warning-Count': String(result.warnings.length),
      'X-Template-Id': result.templateId,
      'X-Template-Preview-Status': result.previewAvailable ? 'ready' : 'unavailable',
    })
    if (result.previewAvailable) {
      response.set('Link', `<${buildAppScopedPath(
        `${API_V1_PATH}/templates/${result.templateId}/preview`,
        appId,
      )}>; rel="preview"`)
    }
    response.json(result.templateJson)
  }

  const batchCreate: RequestHandler = async (request, response) => {
    validatePowerPointRequest(request)

    const appId = parseAppId(request)
    const kind = parseTemplateKind(request.query.kind, true)
    const result = await runCancellableImport(
      request,
      response,
      (signal) => service.batchImport(request.body, kind, signal, appId),
    )
    if (response.writableEnded) {
      return
    }
    response.status(201).set({
      'X-Imported-Template-Count': String(result.templates.length),
      'X-PowerPoint-Warning-Count': String(result.warnings.length),
    }).json(result)
  }

  const createV2: RequestHandler = async (request, response) => {
    validatePowerPointRequest(request)
    const appId = parseAppId(request)
    const kind = parseTemplateKind(request.query.kind, true)
    const result = await runCancellableImport(
      request,
      response,
      (signal) => service.importV2(request.body, kind, signal, appId),
    )
    if (response.writableEnded) return
    response.status(201).set({
      Location: buildAppScopedPath(`${API_V1_PATH}/templates/${result.templateId}`, appId),
      'X-PowerPoint-Warning-Count': String(result.warnings.length),
      'X-Template-Id': result.templateId,
      'X-Template-Preview-Status': result.previewAvailable ? 'ready' : 'unavailable',
    }).json(result)
  }

  const find: RequestHandler<{ templateId: string }> = (request, response) => {
    const template = service.find(request.params.templateId, parseAppId(request))
    if (!template) {
      throw new ApiError(404, 'template_not_found', 'The requested template does not exist.')
    }
    response.json(template.templateJson)
  }

  const findV2: RequestHandler<{ templateId: string }> = (request, response) => {
    const options = parseTemplateV2ReadOptions(request.query)
    const template = service.findV2(
      request.params.templateId,
      options,
      parseAppId(request),
    )
    if (!template) {
      throw new ApiError(404, 'template_not_found', 'The requested template does not exist.')
    }
    response.set('Cache-Control', 'private, no-store').json(template)
  }

  const findAsset: RequestHandler<{ assetId: string; templateId: string }> = (request, response) => {
    const asset = service.findAsset(
      request.params.templateId,
      request.params.assetId,
      parseAppId(request),
    )
    if (!asset) {
      throw new ApiError(404, 'template_asset_not_found', 'The requested template image does not exist.')
    }

    response
      .status(200)
      .set({
        'Cache-Control': 'no-store',
        'Content-Length': String(asset.bytes.length),
        'Content-Type': asset.contentType,
        'X-Content-Type-Options': 'nosniff',
      })
      .send(asset.bytes)
  }

  const list: RequestHandler = (request, response) => {
    response.json(service.list(
      parseTemplateKind(request.query.kind, false),
      parseAppId(request),
    ))
  }

  const listPreviews: RequestHandler = (request, response) => {
    response
      .set('Cache-Control', 'private, no-store')
      .json(service.listPreviews(parsePreviewPage(request.query.page), parseAppId(request)))
  }

  const findPreview: RequestHandler<{ templateId: string }> = (request, response) => {
    const appId = parseAppId(request)
    if (!service.find(request.params.templateId, appId)) {
      throw new ApiError(404, 'template_not_found', 'The requested template does not exist.')
    }
    const preview = service.findPreview(request.params.templateId, appId)
    if (!preview) {
      throw new ApiError(404, 'template_preview_not_found', 'The requested template preview does not exist.')
    }
    response.status(200).set({
      'Cache-Control': 'private, no-store',
      'Content-Length': String(preview.bytes.length),
      'Content-Type': preview.contentType,
      'X-Content-Type-Options': 'nosniff',
    }).send(preview.bytes)
  }

  const remove: RequestHandler<{ templateId: string }> = (request, response) => {
    if (!service.delete(request.params.templateId, parseAppId(request))) {
      throw new ApiError(404, 'template_not_found', 'The requested template does not exist.')
    }
    response.sendStatus(204)
  }

  return {
    batchCreate,
    create,
    createV2,
    find,
    findAsset,
    findPreview,
    findV2,
    list,
    listPreviews,
    remove,
  }
}

function validatePowerPointRequest(request: Request) {
  if (!request.is('application/vnd.openxmlformats-officedocument.presentationml.presentation')) {
    throw new ApiError(
      415,
      'unsupported_media_type',
      'Content-Type must be application/vnd.openxmlformats-officedocument.presentationml.presentation.',
    )
  }
  if (!Buffer.isBuffer(request.body) || request.body.length === 0) {
    throw new ApiError(400, 'missing_powerpoint', 'A PowerPoint file is required in the request body.')
  }
}

async function runCancellableImport<Result>(
  request: Request,
  response: Response,
  operation: (signal: AbortSignal) => Promise<Result>,
) {
  const abortController = new AbortController()
  const abort = () => abortController.abort()
  const abortOnResponseClose = () => {
    if (!response.writableEnded) abort()
  }
  request.once('aborted', abort)
  response.once('timeout', abort)
  response.once('close', abortOnResponseClose)

  try {
    return await operation(abortController.signal)
  } finally {
    request.off('aborted', abort)
    response.off('timeout', abort)
    response.off('close', abortOnResponseClose)
  }
}

function parseTemplateKind(value: unknown, useDefault: boolean): TemplateKind | undefined {
  if (value === undefined && useDefault) {
    return 'diagram'
  }
  if (value === undefined) {
    return undefined
  }
  if (value === 'diagram' || value === 'commentary') {
    return value
  }
  throw new ApiError(400, 'invalid_template_kind', 'Template kind must be diagram or commentary.')
}

function parsePreviewPage(value: unknown) {
  if (value === undefined) {
    return 1
  }
  if (typeof value !== 'string' || !/^[1-9]\d*$/u.test(value)) {
    throw new ApiError(400, 'invalid_preview_page', 'Preview page must be a positive integer.')
  }

  const page = Number(value)
  if (
    !Number.isSafeInteger(page)
    || !Number.isSafeInteger((page - 1) * TEMPLATE_PREVIEW_PAGE_SIZE)
  ) {
    throw new ApiError(400, 'invalid_preview_page', 'Preview page must be a positive integer.')
  }
  return page
}

function parseTemplateV2ReadOptions(query: Record<string, unknown>): TemplateV2ReadOptions {
  for (const key of Object.keys(query)) {
    if (key !== 'appId' && key !== 'embeddings' && key !== 'metadata') {
      throw new ApiError(400, 'invalid_template_v2_query', 'The template query is invalid.')
    }
  }

  return {
    includeEmbeddings: parseBooleanQuery(query.embeddings),
    includeMetadata: parseBooleanQuery(query.metadata),
  }
}

function parseBooleanQuery(value: unknown) {
  if (value === undefined || value === 'false') return false
  if (value === 'true') return true
  throw new ApiError(400, 'invalid_template_v2_query', 'The template query is invalid.')
}
