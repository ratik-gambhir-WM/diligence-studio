import type { Request } from 'express'

import { ApiError } from './errors'

export const DEFAULT_APP_ID = 'DiligenceStudio_WestMonroe'
export const APP_ID_HEADER = 'X-App-Id'

const APP_ID_PATTERN = /^[A-Za-z0-9][A-Za-z0-9_-]{0,127}$/u

export function parseAppId(request: Request) {
  const headerValue = request.get(APP_ID_HEADER)
  const queryValue = request.query.appId
  const value = headerValue ?? queryValue ?? DEFAULT_APP_ID
  if (typeof value !== 'string' || !APP_ID_PATTERN.test(value)) {
    throw new ApiError(
      400,
      'invalid_app_id',
      'App ID must be 1-128 letters, numbers, underscores, or hyphens.',
    )
  }
  if (headerValue !== undefined && queryValue !== undefined && headerValue !== queryValue) {
    throw new ApiError(400, 'conflicting_app_id', 'App ID values must match.')
  }

  return value
}

export function buildAppScopedPath(path: string, appId: string) {
  const separator = path.includes('?') ? '&' : '?'
  return `${path}${separator}appId=${encodeURIComponent(appId)}`
}
