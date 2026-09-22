import { Router } from 'express'

import { createImportHandlers } from '../handlers/importHandlers'
import type { ImportService } from '../services/ImportTemplateService'

export function createTemplateRouter(service: ImportService) {
  const router = Router()
  const handlers = createImportHandlers(service)

  router.get('/', handlers.list)
  router.all('/', (_request, response) => {
    response.status(405).end()
  })
  router.get('/previews', handlers.listPreviews)
  router.all('/previews', (_request, response) => {
    response.status(405).end()
  })
  router.get('/:templateId/preview', handlers.findPreview)
  router.all('/:templateId/preview', (_request, response) => {
    response.status(405).end()
  })
  router.get('/:templateId', handlers.find)
  router.delete('/:templateId', handlers.remove)
  router.all('/:templateId', (_request, response) => {
    response.status(405).end()
  })

  return router
}

/** Expose the intentionally small v2 template inspection contract. */
export function createTemplateV2Router(service: ImportService) {
  const router = Router()
  const handlers = createImportHandlers(service)

  router.get('/:templateId', handlers.findV2)
  router.all('/:templateId', (_request, response) => {
    response.status(405).end()
  })
  router.all('/', (_request, response) => {
    response.status(405).end()
  })

  return router
}
