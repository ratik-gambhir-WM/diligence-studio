import { Router, raw } from 'express'

import { createImportHandlers } from '../handlers/importHandlers'
import type { ImportService } from '../services/ImportTemplateService'

const POWERPOINT_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.presentation'

export function createImportRouter(service: ImportService, maxUploadBytes: number) {
  const router = Router()
  const handlers = createImportHandlers(service)

  router.post(
    '/',
    raw({ inflate: false, limit: maxUploadBytes, type: POWERPOINT_CONTENT_TYPE }),
    handlers.create,
  )
  router.get('/:templateId/assets/:assetId', handlers.findAsset)
  router.all('/:templateId/assets/:assetId', (_request, response) => {
    response.status(405).end()
  })
  router.get('/:templateId', handlers.find)
  router.all('/:templateId', (_request, response) => {
    response.status(405).end()
  })
  router.all('/', (_request, response) => {
    response.status(405).end()
  })

  return router
}

export function createBatchImportRouter(service: ImportService, maxUploadBytes: number) {
  const router = Router()
  const handlers = createImportHandlers(service)

  router.post(
    '/',
    raw({ inflate: false, limit: maxUploadBytes, type: POWERPOINT_CONTENT_TYPE }),
    handlers.batchCreate,
  )
  router.all('/', (_request, response) => {
    response.status(405).end()
  })

  return router
}

/** Mount the sole public v2 endpoint with the existing bounded raw PowerPoint parser. */
export function createImportV2Router(service: ImportService, maxUploadBytes: number) {
  const router = Router()
  const handlers = createImportHandlers(service)
  router.post(
    '/',
    raw({ inflate: false, limit: maxUploadBytes, type: POWERPOINT_CONTENT_TYPE }),
    handlers.createV2,
  )
  router.all('/', (_request, response) => response.status(405).end())
  return router
}
