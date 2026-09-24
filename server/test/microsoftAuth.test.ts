// @vitest-environment node

import request from 'supertest'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

import { createApp } from '../src/app'
import { SqliteTemplateRepository } from '../src/repositories/SqliteTemplateRepository'
import { ExportPowerPointService } from '../src/services/ExportPowerPointService'
import { ImportTemplateService } from '../src/services/ImportTemplateService'
import { LibraryPowerPointConverter } from '../src/services/PowerPointConverter'

let templates: SqliteTemplateRepository

beforeEach(() => {
  templates = new SqliteTemplateRepository(':memory:')
  vi.stubEnv('MICROSOFT_CLIENT_ID', 'client-id')
  vi.stubEnv('MICROSOFT_CLIENT_SECRET', 'client-secret')
  vi.stubEnv('MICROSOFT_REDIRECT_URI', 'http://localhost:43127/api/auth/callback')
  vi.stubEnv('MICROSOFT_FRONTEND_ORIGIN', 'http://localhost:5173')
  vi.stubEnv('MICROSOFT_COOKIE_SECURE', 'false')
})

afterEach(() => {
  vi.restoreAllMocks()
  vi.unstubAllEnvs()
  templates.close()
})

describe('Microsoft authentication', () => {
  it('binds the OAuth callback to the browser that started login', async () => {
    const fetchMock = vi.spyOn(globalThis, 'fetch')
    const app = createTestApp()
    const login = await startLogin(app)
    const state = login.location.searchParams.get('state')

    expect(state).toBeTruthy()
    if (!state) throw new Error('The login redirect did not contain state.')

    const callback = await request(app)
      .get(`/api/auth/callback?code=attacker-code&state=${encodeURIComponent(state)}`)
      .expect(302)

    expect(callback.headers.location).toContain('/login?authError=')
    expect(fetchMock).not.toHaveBeenCalled()
  })

  it('creates a secure server session and exposes only the normalized user', async () => {
    const fetchMock = vi.spyOn(globalThis, 'fetch')
      .mockResolvedValueOnce(jsonResponse({
        access_token: 'access-token',
        expires_in: 3600,
        refresh_token: 'refresh-token',
      }))
      .mockResolvedValueOnce(jsonResponse({
        displayName: 'Test User',
        id: 'user-id',
        mail: 'test.user@example.com',
      }))
    const app = createTestApp()
    const login = await startLogin(app)
    const state = login.location.searchParams.get('state')

    expect(state).toBeTruthy()
    if (!state) throw new Error('The login redirect did not contain state.')

    const callback = await request(app)
      .get(`/api/auth/callback?code=code&state=${encodeURIComponent(state)}`)
      .set('Cookie', login.stateCookie)
      .expect(302)
    const setCookies = getSetCookies(callback)
    const sessionCookie = setCookies.find((cookie) => cookie.startsWith('diligence_studio_session='))

    expect(callback.headers.location).toBe('http://localhost:5173')
    expect(setCookies).toEqual(expect.arrayContaining([
      expect.stringContaining('diligence_studio_auth_state=;'),
      expect.stringContaining('diligence_studio_session='),
    ]))
    expect(sessionCookie).toBeDefined()
    expect(sessionCookie).toContain('HttpOnly')
    expect(sessionCookie).toContain('SameSite=Lax')
    expect(sessionCookie).not.toContain('Secure')

    if (!sessionCookie) throw new Error('The login callback did not set a session cookie.')
    const me = await request(app)
      .get('/api/auth/me')
      .set('Cookie', sessionCookie.split(';', 1)[0] ?? '')
      .expect(200)

    expect(me.headers['cache-control']).toBe('private, no-store')
    expect(me.body).toEqual({
      authenticated: true,
      user: {
        email: 'test.user@example.com',
        id: 'user-id',
        name: 'Test User',
      },
    })
    expect(fetchMock).toHaveBeenCalledTimes(2)
  })

  it('expires a session when refresh-token exchange fails', async () => {
    vi.spyOn(globalThis, 'fetch')
      .mockResolvedValueOnce(jsonResponse({
        access_token: 'access-token',
        expires_in: 1,
        refresh_token: 'refresh-token',
      }))
      .mockResolvedValueOnce(jsonResponse({
        displayName: 'Test User',
        id: 'user-id',
        userPrincipalName: 'test.user@example.com',
      }))
      .mockResolvedValueOnce(jsonResponse({ error: 'invalid_grant' }, { status: 400 }))
    const app = createTestApp()
    const login = await startLogin(app)
    const state = login.location.searchParams.get('state')

    expect(state).toBeTruthy()
    if (!state) throw new Error('The login redirect did not contain state.')

    const callback = await request(app)
      .get(`/api/auth/callback?code=code&state=${encodeURIComponent(state)}`)
      .set('Cookie', login.stateCookie)
      .expect(302)
    const sessionCookie = getSetCookies(callback)
      .find((cookie) => cookie.startsWith('diligence_studio_session='))

    expect(sessionCookie).toBeDefined()
    if (!sessionCookie) throw new Error('The login callback did not set a session cookie.')

    const me = await request(app)
      .get('/api/auth/me')
      .set('Cookie', sessionCookie.split(';', 1)[0] ?? '')
      .expect(401)

    expect(me.body).toEqual({ authenticated: false })
  })

  it('protects versioned API routes when Microsoft support is enabled', async () => {
    const response = await request(createTestApp(true))
      .get('/api/v1/templates')
      .expect(401)

    expect(response.body.error).toMatchObject({
      code: 'authentication_required',
      message: 'Microsoft sign-in is required.',
    })
  })
})

function createTestApp(requireMicrosoftAuth = false) {
  return createApp({
    exportService: new ExportPowerPointService(templates),
    importService: new ImportTemplateService(new LibraryPowerPointConverter(), templates),
    maxExportJsonBytes: 1024 * 1024,
    maxUploadBytes: 25 * 1024 * 1024,
    requireMicrosoftAuth,
  })
}

async function startLogin(app: ReturnType<typeof createApp>) {
  const response = await request(app).get('/api/auth/login').expect(302)
  const stateCookie = getSetCookies(response)
    .find((cookie) => cookie.startsWith('diligence_studio_auth_state='))
  if (!stateCookie) throw new Error('The login redirect did not set an OAuth state cookie.')

  return {
    location: new URL(response.headers.location),
    stateCookie: stateCookie.split(';', 1)[0] ?? '',
  }
}

function getSetCookies(response: { headers: Record<string, unknown> }) {
  const value = response.headers['set-cookie']
  if (Array.isArray(value)) {
    return value.filter((cookie): cookie is string => typeof cookie === 'string')
  }
  return typeof value === 'string' ? [value] : []
}

function jsonResponse(body: unknown, init: ResponseInit = {}) {
  const headers = new Headers(init.headers)
  headers.set('Content-Type', 'application/json')
  return new Response(JSON.stringify(body), { ...init, headers })
}
