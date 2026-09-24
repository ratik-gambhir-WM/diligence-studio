// @vitest-environment node

import path from 'node:path'
import { fileURLToPath } from 'node:url'

import { describe, expect, it } from 'vitest'

import { loadServerConfig } from '../src/config'

describe('server configuration', () => {
  it('uses explicit, validated values', () => {
    expect(loadServerConfig({
      MAX_EXPORT_JSON_BYTES: '8192',
      MAX_PPTX_UPLOAD_BYTES: '4096',
      MAX_TEMPLATE_PREVIEW_BYTES: '2048',
      HOST: '127.0.0.1',
      PORT: '4321',
      SQLITE_DB_PATH: './tmp/templates.sqlite',
      TEMPLATE_PREVIEW_PROVIDER: 'disabled',
      TEMPLATE_PREVIEW_RENDER_SIZE: '1200',
      TEMPLATE_PREVIEW_RENDER_URL: 'https://preview.example.test/_internal/template-preview',
      TEMPLATE_PREVIEW_TIMEOUT_MS: '5000',
    })).toEqual({
      databasePath: path.resolve('./tmp/templates.sqlite'),
      host: '127.0.0.1',
      maxExportJsonBytes: 8192,
      maxPreviewBytes: 2048,
      maxUploadBytes: 4096,
      microsoftCookieSecure: true,
      microsoftSupport: false,
      port: 4321,
      previewProvider: 'disabled',
      previewRenderSize: 1200,
      previewRenderUrl: 'https://preview.example.test/_internal/template-preview',
      previewTimeoutMs: 5000,
      requestTimeoutMs: 30_000,
      sharePointAllowedHosts: ['relentlessblue.sharepoint.com'],
    })
  })

  it('rejects invalid ports', () => {
    expect(() => loadServerConfig({ PORT: '0' })).toThrow('PORT must be between 1 and 65535.')
    expect(() => loadServerConfig({ PORT: '70000' })).toThrow('PORT must be between 1 and 65535.')
    expect(() => loadServerConfig({ PORT: '1e3' })).toThrow('PORT must be between 1 and 65535.')
  })

  it('uses the Node server bind defaults and validates HOST as an IP address', () => {
    expect(loadServerConfig({})).toMatchObject({
      databasePath: fileURLToPath(new URL('../data/templates.sqlite', import.meta.url)),
      host: '0.0.0.0',
      port: 43127,
      previewProvider: 'headless',
      previewRenderUrl: 'http://localhost:5173/_internal/template-preview',
      requestTimeoutMs: 30_000,
      microsoftCookieSecure: true,
      microsoftSupport: false,
    })
    expect(() => loadServerConfig({ HOST: 'localhost' })).toThrow('HOST must be an IP address.')
  })

  it('rejects unsafe preview render URLs and unsupported providers', () => {
    expect(() => loadServerConfig({ TEMPLATE_PREVIEW_RENDER_URL: 'file:///tmp/preview.html' }))
      .toThrow('TEMPLATE_PREVIEW_RENDER_URL must be a valid HTTP or HTTPS URL.')
    expect(() => loadServerConfig({ TEMPLATE_PREVIEW_RENDER_URL: 'https://user:secret@example.test' }))
      .toThrow('TEMPLATE_PREVIEW_RENDER_URL must be a valid HTTP or HTTPS URL.')
    expect(() => loadServerConfig({ TEMPLATE_PREVIEW_PROVIDER: 'browser' }))
      .toThrow('TEMPLATE_PREVIEW_PROVIDER must be headless, quicklook, or disabled.')
    expect(() => loadServerConfig({ TEMPLATE_PREVIEW_RENDER_SIZE: '5000' }))
      .toThrow('TEMPLATE_PREVIEW_RENDER_SIZE must be between 320 and 4096 pixels.')
  })

  it('validates SharePoint host allowlists', () => {
    expect(loadServerConfig({ SHAREPOINT_ALLOWED_HOSTS: ' Contoso.SharePoint.com,contoso.sharepoint.com ' }))
      .toMatchObject({ sharePointAllowedHosts: ['contoso.sharepoint.com'] })
    expect(() => loadServerConfig({ SHAREPOINT_ALLOWED_HOSTS: 'https://contoso.sharepoint.com' }))
      .toThrow('SHAREPOINT_ALLOWED_HOSTS must contain comma-separated hostnames.')
  })

  it('parses Microsoft support and secure-cookie settings', () => {
    expect(loadServerConfig({ MICROSOFT_SUPPORT: ' TRUE ', MICROSOFT_COOKIE_SECURE: 'false' }))
      .toMatchObject({ microsoftSupport: true, microsoftCookieSecure: false })
    expect(() => loadServerConfig({ MICROSOFT_SUPPORT: 'yes' }))
      .toThrow('MICROSOFT_SUPPORT must be true or false.')
  })
})
