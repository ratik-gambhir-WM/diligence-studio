import { createHash, randomBytes, timingSafeEqual } from 'node:crypto'

import { Router, type Request, type Response } from 'express'

const SESSION_COOKIE_NAME = 'diligence_studio_session'
const AUTH_STATE_COOKIE_NAME = 'diligence_studio_auth_state'
const SESSION_TTL_SECONDS = 8 * 60 * 60
const AUTH_STATE_TTL_MS = 10 * 60 * 1000
const MAX_PENDING_AUTHORIZATIONS = 1_000
const MAX_SESSIONS = 10_000
const MICROSOFT_SCOPES = 'openid profile email offline_access User.Read Files.Read Sites.Read.All'

type MicrosoftAuthConfig = {
  clientId: string
  clientSecret: string
  redirectUri: string
  tenantId: string
}

type PendingAuthorization = {
  codeVerifier: string
  expiresAt: number
}

export type MicrosoftUser = {
  email: string
  id: string
  name: string
}

type MicrosoftSession = {
  accessToken: string
  expiresAt: number
  refreshToken?: string
  sessionExpiresAt: number
  user: MicrosoftUser
}

export type MicrosoftAuthRouterOptions = {
  cookieSecure?: boolean
}

const pendingAuthorizations = new Map<string, PendingAuthorization>()
const sessions = new Map<string, MicrosoftSession>()

export function createMicrosoftAuthRouter(options: MicrosoftAuthRouterOptions = {}) {
  const router = Router()
  const cookieSecure = options.cookieSecure ?? getCookieSecureDefault()

  router.get('/login', (_request, response) => {
    try {
      beginLogin(response, cookieSecure)
    } catch (error) {
      if (error instanceof AuthError) {
        response.status(error.statusCode).json({ error: error.message })
        return
      }
      throw error
    }
  })

  router.get('/callback', async (request, response) => {
    await completeLogin(request, response, cookieSecure)
  })

  router.get('/me', async (request, response) => {
    await respondWithCurrentUser(request, response)
  })

  router.post('/logout', (request, response) => {
    logout(request, response, cookieSecure)
  })

  return router
}

export async function getMicrosoftAccessToken(request: Request, signal?: AbortSignal) {
  pruneExpiredSessions()
  const sessionId = getCookie(request, SESSION_COOKIE_NAME)
  const session = sessionId ? sessions.get(sessionId) : undefined

  if (!sessionId || !session || session.sessionExpiresAt <= Date.now()) {
    if (sessionId) sessions.delete(sessionId)
    throw new AuthError(401, 'Microsoft sign-in is required.')
  }

  if (session.expiresAt > Date.now() + 60_000) {
    return session.accessToken
  }

  if (!session.refreshToken) {
    sessions.delete(sessionId)
    throw new AuthError(401, 'Your Microsoft session has expired. Sign in again.')
  }

  let refreshed
  try {
    refreshed = await exchangeToken({
      grantType: 'refresh_token',
      refreshToken: session.refreshToken,
      signal,
    })
  } catch (error) {
    if (signal?.aborted) {
      throw error
    }
    sessions.delete(sessionId)
    throw new AuthError(401, 'Your Microsoft session has expired. Sign in again.')
  }
  session.accessToken = refreshed.accessToken
  session.expiresAt = refreshed.expiresAt
  session.refreshToken = refreshed.refreshToken ?? session.refreshToken
  return session.accessToken
}

function beginLogin(response: Response, cookieSecure: boolean) {
  pruneExpiredAuthorizations()
  if (pendingAuthorizations.size >= MAX_PENDING_AUTHORIZATIONS) {
    throw new AuthError(503, 'Microsoft sign-in is temporarily unavailable. Please try again.')
  }

  const config = getAuthConfig()
  const state = randomBytes(32).toString('hex')
  const codeVerifier = randomBytes(48).toString('base64url')
  const codeChallenge = createHash('sha256').update(codeVerifier).digest('base64url')

  pendingAuthorizations.set(state, {
    codeVerifier,
    expiresAt: Date.now() + AUTH_STATE_TTL_MS,
  })

  response.append('Set-Cookie', buildCookie(
    AUTH_STATE_COOKIE_NAME,
    state,
    AUTH_STATE_TTL_MS / 1000,
    cookieSecure,
    '/api/auth',
  ))

  const authorizationUrl = new URL(
    `https://login.microsoftonline.com/${encodeURIComponent(config.tenantId)}/oauth2/v2.0/authorize`,
  )
  authorizationUrl.search = new URLSearchParams({
    client_id: config.clientId,
    code_challenge: codeChallenge,
    code_challenge_method: 'S256',
    redirect_uri: config.redirectUri,
    response_mode: 'query',
    response_type: 'code',
    scope: MICROSOFT_SCOPES,
    state,
  }).toString()

  response.redirect(302, authorizationUrl.toString())
}

async function completeLogin(request: Request, response: Response, cookieSecure: boolean) {
  const error = request.query.error
  if (typeof error === 'string' && error) {
    clearAuthStateCookie(response, cookieSecure)
    redirectToLogin(response, 'Microsoft sign-in was cancelled or denied.')
    return
  }

  const state = typeof request.query.state === 'string' ? request.query.state : undefined
  const code = typeof request.query.code === 'string' ? request.query.code : undefined
  const authorization = state ? pendingAuthorizations.get(state) : undefined
  const stateCookie = getCookie(request, AUTH_STATE_COOKIE_NAME)

  if (
    !state
    || !code
    || !authorization
    || authorization.expiresAt < Date.now()
    || !stateCookie
    || !secureStringEqual(stateCookie, state)
  ) {
    if (state) pendingAuthorizations.delete(state)
    clearAuthStateCookie(response, cookieSecure)
    redirectToLogin(response, 'The Microsoft sign-in request expired. Please try again.')
    return
  }

  pendingAuthorizations.delete(state)

  try {
    const token = await exchangeToken({
      code,
      codeVerifier: authorization.codeVerifier,
      grantType: 'authorization_code',
    })
    const user = await fetchMicrosoftUser(token.accessToken)
    pruneExpiredSessions()
    if (sessions.size >= MAX_SESSIONS) {
      throw new AuthError(503, 'Microsoft sign-in is temporarily unavailable. Please try again.')
    }

    const sessionId = randomBytes(32).toString('hex')

    sessions.set(sessionId, {
      accessToken: token.accessToken,
      expiresAt: token.expiresAt,
      refreshToken: token.refreshToken,
      sessionExpiresAt: Date.now() + SESSION_TTL_SECONDS * 1000,
      user,
    })

    clearAuthStateCookie(response, cookieSecure)
    setSessionCookie(response, sessionId, cookieSecure)
    response.redirect(302, getFrontendOrigin())
  } catch {
    console.error('Microsoft sign-in failed.')
    clearAuthStateCookie(response, cookieSecure)
    redirectToLogin(response, 'Microsoft sign-in could not be completed.')
  }
}

async function exchangeToken({
  code,
  codeVerifier,
  grantType,
  refreshToken,
  signal,
}: {
  code?: string
  codeVerifier?: string
  grantType: 'authorization_code' | 'refresh_token'
  refreshToken?: string
  signal?: AbortSignal
}) {
  const config = getAuthConfig()
  const body = new URLSearchParams({
    client_id: config.clientId,
    client_secret: config.clientSecret,
    grant_type: grantType,
    scope: MICROSOFT_SCOPES,
  })

  if (grantType === 'authorization_code') {
    body.set('code', code ?? '')
    body.set('code_verifier', codeVerifier ?? '')
    body.set('redirect_uri', config.redirectUri)
  } else {
    body.set('refresh_token', refreshToken ?? '')
  }

  const response = await fetch(
    `https://login.microsoftonline.com/${encodeURIComponent(config.tenantId)}/oauth2/v2.0/token`,
    {
      body,
      headers: { 'content-type': 'application/x-www-form-urlencoded' },
      method: 'POST',
      signal,
    },
  )
  const payload = await response.json() as unknown
  const accessToken = getStringProperty(payload, 'access_token')

  if (!response.ok || !accessToken) {
    throw new Error('Microsoft token exchange failed.')
  }

  return {
    accessToken,
    expiresAt: Date.now() + (getNumberProperty(payload, 'expires_in') ?? 3600) * 1000,
    refreshToken: getStringProperty(payload, 'refresh_token'),
  }
}

async function fetchMicrosoftUser(accessToken: string): Promise<MicrosoftUser> {
  const response = await fetch(
    'https://graph.microsoft.com/v1.0/me?$select=id,displayName,mail,userPrincipalName',
    { headers: { authorization: `Bearer ${accessToken}` } },
  )
  const payload = await response.json() as unknown
  const id = getStringProperty(payload, 'id')
  const email = getStringProperty(payload, 'mail') ?? getStringProperty(payload, 'userPrincipalName')
  const name = getStringProperty(payload, 'displayName') ?? email ?? 'Microsoft user'

  if (!response.ok || !id || !email) {
    throw new Error('Microsoft Graph user lookup failed.')
  }

  return { email, id, name }
}

async function respondWithCurrentUser(request: Request, response: Response) {
  response.set('Cache-Control', 'private, no-store')

  try {
    await getMicrosoftAccessToken(request)
  } catch (error) {
    if (error instanceof AuthError) {
      response.status(error.statusCode).json({ authenticated: false })
      return
    }
    throw error
  }

  const sessionId = getCookie(request, SESSION_COOKIE_NAME)
  const session = sessionId ? sessions.get(sessionId) : undefined
  if (!session) {
    response.status(401).json({ authenticated: false })
    return
  }

  response.json({ authenticated: true, user: session.user })
}

function logout(request: Request, response: Response, cookieSecure: boolean) {
  const sessionId = getCookie(request, SESSION_COOKIE_NAME)
  if (sessionId) sessions.delete(sessionId)

  response
    .status(204)
    .set('Cache-Control', 'no-store')
    .append('Set-Cookie', buildCookie(SESSION_COOKIE_NAME, '', 0, cookieSecure, '/'))
    .append('Set-Cookie', buildCookie(AUTH_STATE_COOKIE_NAME, '', 0, cookieSecure, '/api/auth'))
    .end()
}

function getAuthConfig(): MicrosoftAuthConfig {
  const clientId = process.env.MICROSOFT_CLIENT_ID?.trim()
  const clientSecret = process.env.MICROSOFT_CLIENT_SECRET?.trim()
  const redirectUri = process.env.MICROSOFT_REDIRECT_URI?.trim()
    || 'http://localhost:43127/api/auth/callback'
  const tenantId = process.env.MICROSOFT_TENANT_ID?.trim() || 'organizations'

  if (!clientId || !clientSecret) {
    throw new AuthError(
      500,
      'Microsoft sign-in is not configured. Set MICROSOFT_CLIENT_ID and MICROSOFT_CLIENT_SECRET.',
    )
  }

  return { clientId, clientSecret, redirectUri, tenantId }
}

function setSessionCookie(response: Response, sessionId: string, cookieSecure: boolean) {
  response.append('Set-Cookie', buildCookie(
    SESSION_COOKIE_NAME,
    sessionId,
    SESSION_TTL_SECONDS,
    cookieSecure,
    '/',
  ))
}

function clearAuthStateCookie(response: Response, cookieSecure: boolean) {
  response.append('Set-Cookie', buildCookie(
    AUTH_STATE_COOKIE_NAME,
    '',
    0,
    cookieSecure,
    '/api/auth',
  ))
}

function buildCookie(name: string, value: string, maxAgeSeconds: number, secure: boolean, path: string) {
  const secureAttribute = secure ? '; Secure' : ''
  return `${name}=${encodeURIComponent(value)}; HttpOnly; Path=${path}; SameSite=Lax; Max-Age=${maxAgeSeconds}${secureAttribute}`
}

function redirectToLogin(response: Response, message: string) {
  const location = new URL('/login', getFrontendOrigin())
  location.searchParams.set('authError', message)
  response.redirect(302, location.toString())
}

function getFrontendOrigin() {
  return process.env.MICROSOFT_FRONTEND_ORIGIN?.trim() || 'http://localhost:5173'
}

function getCookie(request: Request, name: string) {
  const cookie = request.headers.cookie
    ?.split(';')
    .map((part) => part.trim())
    .find((part) => part.startsWith(`${name}=`))
  if (!cookie) return undefined

  try {
    return decodeURIComponent(cookie.slice(name.length + 1))
  } catch {
    return undefined
  }
}

function getCookieSecureDefault() {
  const configured = process.env.MICROSOFT_COOKIE_SECURE?.trim().toLowerCase()
  return configured === undefined ? true : configured === 'true'
}

function secureStringEqual(left: string, right: string) {
  const leftBytes = Buffer.from(left)
  const rightBytes = Buffer.from(right)
  return leftBytes.length === rightBytes.length && timingSafeEqual(leftBytes, rightBytes)
}

function pruneExpiredAuthorizations() {
  const now = Date.now()
  for (const [state, authorization] of pendingAuthorizations) {
    if (authorization.expiresAt <= now) {
      pendingAuthorizations.delete(state)
    }
  }
}

function pruneExpiredSessions() {
  const now = Date.now()
  for (const [sessionId, session] of sessions) {
    if (session.sessionExpiresAt <= now) {
      sessions.delete(sessionId)
    }
  }
}

function getStringProperty(value: unknown, property: string) {
  if (!isRecord(value) || typeof value[property] !== 'string') return undefined
  return value[property]
}

function getNumberProperty(value: unknown, property: string) {
  if (!isRecord(value) || typeof value[property] !== 'number' || !Number.isFinite(value[property])) {
    return undefined
  }
  return value[property]
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

export class AuthError extends Error {
  readonly statusCode: number

  constructor(statusCode: number, message: string) {
    super(message)
    this.name = 'AuthError'
    this.statusCode = statusCode
  }
}
