import { randomUUID } from 'node:crypto'

import express from 'express'

import { createApiLoggingMiddleware, recordApiError, type ApiLogger } from './apiLogging'
import { API_V1_PATH, API_V2_PATH } from './apiPaths'
import { ApiError, errorHandler, notFoundHandler } from './errors'
import { createExportRouter } from './routes/exportRoutes'
import { createBatchImportRouter, createImportRouter, createImportV2Router } from './routes/importRoutes'
import { createTemplateRouter, createTemplateV2Router } from './routes/templateRoutes'
import type { ExportPowerPointUseCase } from './services/ExportPowerPointService'
import type { ImportService } from './services/ImportTemplateService'

export type AppDependencies = {
  exportService: ExportPowerPointUseCase
  importService: ImportService
  maxExportJsonBytes: number
  maxUploadBytes: number
  logger?: ApiLogger
  requestTimeoutMs?: number
}

const DEFAULT_REQUEST_TIMEOUT_MS = 90_000

export function createApp(dependencies: AppDependencies) {
  const app = express()
  const apiV1 = express.Router()
  const apiV2 = express.Router()
  app.disable('x-powered-by')

  app.use((_request, response, next) => {
    const requestId = randomUUID()
    response.locals.requestId = requestId
    response.setHeader('X-Request-Id', requestId)
    next()
  })
  if (dependencies.logger) {
    app.use(createApiLoggingMiddleware(dependencies.logger))
  }
  app.use((request, response, next) => {
    const contentEncoding = request.get('Content-Encoding')
    if (contentEncoding && contentEncoding.toLowerCase() !== 'identity') {
      throw new ApiError(
        415,
        'unsupported_content_encoding',
        'Compressed request bodies are not supported.',
      )
    }

    response.setTimeout(dependencies.requestTimeoutMs ?? DEFAULT_REQUEST_TIMEOUT_MS, () => {
      if (!response.headersSent) {
        recordApiError(response, 'request_timeout')
        response.status(408).end()
      }
    })
    next()
  })
  apiV1.use('/export', createExportRouter(
    dependencies.exportService,
    dependencies.maxExportJsonBytes,
    dependencies.maxUploadBytes,
  ))
  apiV1.use(
    '/import',
    createImportRouter(dependencies.importService, dependencies.maxUploadBytes),
  )
  apiV1.use(
    '/batchImport',
    createBatchImportRouter(dependencies.importService, dependencies.maxUploadBytes),
  )
  apiV1.use('/templates', createTemplateRouter(dependencies.importService))
  app.use(API_V1_PATH, apiV1)
  apiV2.use('/import', createImportV2Router(dependencies.importService, dependencies.maxUploadBytes))
  apiV2.use('/templates', createTemplateV2Router(dependencies.importService))
  app.use(API_V2_PATH, apiV2)
  app.use(notFoundHandler)
  app.use(errorHandler)

  return app
}
