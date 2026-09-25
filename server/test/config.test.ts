// @vitest-environment node

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
      REQUEST_TIMEOUT_MS: '120000',
      TEMPLATE_PREVIEW_RENDER_SIZE: '1200',
      QUARRY_TEMPLATE_PREVIEW_RENDER_URL: 'https://quarry-preview.example.test/_internal/template-preview',
      TEMPLATE_PREVIEW_TIMEOUT_MS: '5000',
    })).toEqual({
      host: '127.0.0.1',
      openaiApiKey: null,
      openaiSlideClassificationModel: null,
      openaiSlideEmbeddingDimensions: undefined,
      openaiSlideEmbeddingModel: null,
      maxExportJsonBytes: 8192,
      maxPreviewBytes: 2048,
      maxUploadBytes: 4096,
      port: 4321,
      previewRenderSize: 1200,
      quarryPreviewRenderUrl: 'https://quarry-preview.example.test/_internal/template-preview',
      previewTimeoutMs: 5000,
      requestTimeoutMs: 120_000,
      slideClassificationMaxImageBytes: 5 * 1024 * 1024,
      slideClassificationMaxTextChars: 12_000,
      slideClassificationProvider: 'disabled',
      slideClassificationTimeoutMs: 45_000,
      slideEmbeddingMaxTextBytes: 32_000,
      slideEmbeddingTimeoutMs: 20_000,
    })
  })

  it('rejects invalid ports', () => {
    expect(() => loadServerConfig({ PORT: '0' })).toThrow('PORT must be between 1 and 65535.')
    expect(() => loadServerConfig({ PORT: '70000' })).toThrow('PORT must be between 1 and 65535.')
    expect(() => loadServerConfig({ PORT: '1e3' })).toThrow('PORT must be between 1 and 65535.')
  })

  it('uses the Node server bind defaults and validates HOST as an IP address', () => {
    expect(loadServerConfig({})).toMatchObject({
      host: '0.0.0.0',
      port: 43127,
      quarryPreviewRenderUrl: 'http://localhost:1420/_internal/template-preview',
      requestTimeoutMs: 90_000,
    })
    expect(() => loadServerConfig({ HOST: 'localhost' })).toThrow('HOST must be an IP address.')
  })

  it('rejects unsafe Quarry preview render URLs', () => {
    expect(() => loadServerConfig({ QUARRY_TEMPLATE_PREVIEW_RENDER_URL: 'file:///tmp/preview.html' }))
      .toThrow('QUARRY_TEMPLATE_PREVIEW_RENDER_URL must be a valid HTTP or HTTPS URL.')
    expect(() => loadServerConfig({ TEMPLATE_PREVIEW_RENDER_SIZE: '5000' }))
      .toThrow('TEMPLATE_PREVIEW_RENDER_SIZE must be between 320 and 4096 pixels.')
  })

  it('validates the complete OpenAI classification configuration without exposing values', () => {
    expect(() => loadServerConfig({ SLIDE_CLASSIFICATION_PROVIDER: 'openai' }))
      .toThrow('OPENAI_API_KEY is required')
    expect(() => loadServerConfig({
      SLIDE_CLASSIFICATION_PROVIDER: 'openai',
      OPENAI_API_KEY: 'secret-value',
    })).toThrow('OPENAI_SLIDE_CLASSIFICATION_MODEL is required')
    expect(() => loadServerConfig({
      SLIDE_CLASSIFICATION_PROVIDER: 'openai',
      OPENAI_API_KEY: 'secret-value',
      OPENAI_SLIDE_CLASSIFICATION_MODEL: 'configured-model',
    })).toThrow('OPENAI_SLIDE_EMBEDDING_MODEL is required')
    expect(loadServerConfig({
      SLIDE_CLASSIFICATION_PROVIDER: 'openai',
      OPENAI_API_KEY: 'secret-value',
      OPENAI_SLIDE_CLASSIFICATION_MODEL: 'configured-model',
      OPENAI_SLIDE_EMBEDDING_DIMENSIONS: '512',
      OPENAI_SLIDE_EMBEDDING_MODEL: 'text-embedding-3-small',
    })).toMatchObject({
      openaiApiKey: 'secret-value',
      openaiSlideClassificationModel: 'configured-model',
      openaiSlideEmbeddingDimensions: 512,
      openaiSlideEmbeddingModel: 'text-embedding-3-small',
      slideClassificationProvider: 'openai',
    })
  })
})
