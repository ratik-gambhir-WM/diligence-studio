import type { RequestHandler, Response } from 'express'

export type ApiLogLevel = 'error' | 'info' | 'warn'

export type ApiRequestLog = {
  durationMs: number
  errorCode?: string
  errorName?: string
  event: 'api_request_aborted' | 'api_request_completed'
  level: ApiLogLevel
  method: string
  outcome: 'failure' | 'success'
  path: string
  requestId: string
  statusCode: number
  timestamp: string
}

export type ApiLogger = (entry: ApiRequestLog) => void

const ERROR_CODE_LOCAL = 'apiErrorCode'
const ERROR_NAME_LOCAL = 'apiErrorName'

export function createApiLoggingMiddleware(logger: ApiLogger): RequestHandler {
  return (request, response, next) => {
    const startedAt = process.hrtime.bigint()

    const logRequest = (event: ApiRequestLog['event']) => {
      const statusCode = response.statusCode
      const failed = event === 'api_request_aborted' || statusCode >= 400
      const entry: ApiRequestLog = {
        durationMs: elapsedMilliseconds(startedAt),
        event,
        level: failed ? (statusCode >= 500 ? 'error' : 'warn') : 'info',
        method: request.method,
        outcome: failed ? 'failure' : 'success',
        path: pathWithoutQuery(request.originalUrl),
        requestId: response.locals.requestId as string,
        statusCode,
        timestamp: new Date().toISOString(),
      }
      const errorCode = response.locals[ERROR_CODE_LOCAL]
      const errorName = response.locals[ERROR_NAME_LOCAL]
      if (typeof errorCode === 'string') {
        entry.errorCode = errorCode
      }
      if (typeof errorName === 'string') {
        entry.errorName = errorName
      }
      try {
        logger(entry)
      } catch (error) {
        writeApiLogFailureToConsole(entry.requestId, error)
      }
    }

    const onFinish = () => {
      response.off('close', onClose)
      logRequest('api_request_completed')
    }
    const onClose = () => {
      response.off('finish', onFinish)
      if (!response.writableFinished) {
        logRequest('api_request_aborted')
      }
    }

    response.once('finish', onFinish)
    response.once('close', onClose)
    next()
  }
}

export function recordApiError(
  response: Response,
  errorCode: string,
  error?: unknown,
) {
  response.locals[ERROR_CODE_LOCAL] = errorCode
  if (error !== undefined) {
    response.locals[ERROR_NAME_LOCAL] = error instanceof Error ? error.name : 'NonErrorThrown'
  }
}

export const writeApiLogToConsole: ApiLogger = (entry) => {
  const serialized = JSON.stringify(entry)
  if (entry.level === 'error') {
    console.error(serialized)
  } else if (entry.level === 'warn') {
    console.warn(serialized)
  } else {
    console.log(serialized)
  }
}

function writeApiLogFailureToConsole(requestId: string, error: unknown) {
  console.error(JSON.stringify({
    errorName: error instanceof Error ? error.name : 'NonErrorThrown',
    event: 'api_log_write_failed',
    level: 'error',
    requestId,
    timestamp: new Date().toISOString(),
  }))
}

function elapsedMilliseconds(startedAt: bigint) {
  const nanoseconds = process.hrtime.bigint() - startedAt
  return Math.round(Number(nanoseconds) / 10_000) / 100
}

function pathWithoutQuery(originalUrl: string) {
  const queryStart = originalUrl.indexOf('?')
  return queryStart === -1 ? originalUrl : originalUrl.slice(0, queryStart)
}
