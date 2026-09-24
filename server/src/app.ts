import { randomUUID } from 'node:crypto'

import express, { type RequestHandler } from 'express'

import { createApiLoggingMiddleware, recordApiError, type ApiLogger } from './apiLogging'
import {
  AuthError,
  createMicrosoftAuthRouter,
  getMicrosoftAccessToken,
} from './auth/microsoftAuth'
import { API_V1_PATH } from './apiPaths'
import { ApiError, errorHandler, notFoundHandler } from './errors'
import { createExportRouter } from './routes/exportRoutes'
import { createBatchImportRouter, createImportRouter } from './routes/importRoutes'
import { createSharePointRouter } from './routes/sharePointRoutes'
import { createTemplateRouter } from './routes/templateRoutes'
import type { ExportPowerPointUseCase } from './services/ExportPowerPointService'
import type { ImportService } from './services/ImportTemplateService'
import type { SharePointResourceServiceLike } from './services/SharePointResourceService'

export type AppDependencies = {
  exportService: ExportPowerPointUseCase
  importService: ImportService
  maxExportJsonBytes: number
  maxUploadBytes: number
  logger?: ApiLogger
  microsoftCookieSecure?: boolean
  requireMicrosoftAuth?: boolean
  requestTimeoutMs?: number
  sharePointService?: SharePointResourceServiceLike
}

const DEFAULT_REQUEST_TIMEOUT_MS = 30_000

export function createApp(dependencies: AppDependencies) {
  const app = express()
  const apiV1 = express.Router()
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
  app.use('/api/auth', createMicrosoftAuthRouter({
    cookieSecure: dependencies.microsoftCookieSecure,
  }))
  if (dependencies.sharePointService) {
    app.use('/api/v1/sharepoint', createSharePointRouter(dependencies.sharePointService))
  }
  if (dependencies.requireMicrosoftAuth) {
    apiV1.use(createMicrosoftAuthGuard())
  }
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
  app.use(notFoundHandler)
  app.use(errorHandler)

  return app
}

function createMicrosoftAuthGuard(): RequestHandler {
  return async (request, _response, next) => {
    try {
      await getMicrosoftAccessToken(request)
      next()
    } catch (error) {
      if (error instanceof AuthError) {
        next(new ApiError(error.statusCode, 'authentication_required', error.message))
        return
      }
      next(error)
    }
  }
}
