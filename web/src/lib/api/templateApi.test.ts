import { afterEach, describe, expect, it, vi } from 'vitest'

import {
  batchImportTemplates,
  deleteTemplate,
  exportPresentation,
  getTemplate,
  importTemplate,
  insertPresentation,
  listTemplates,
  TemplateApiError,
} from './templateApi'

afterEach(() => {
  vi.unstubAllGlobals()
})

describe('template API client', () => {
  it('validates and resolves the catalog preview URLs under the API base', async () => {
    const fetchMock = vi.fn(async () => jsonResponse({
      templates: [{
        templateId: 'template-1',
        kind: 'diagram',
        title: 'Template one',
        description: 'Description',
        slideCount: 1,
        elementCount: 3,
        previewUrl: '/templates/template-1/preview?appId=DiligenceStudio_WestMonroe',
      }],
    }))
    vi.stubGlobal('fetch', fetchMock)

    await expect(listTemplates('diagram')).resolves.toEqual({
      templates: [expect.objectContaining({
        previewUrl: '/api/v1/templates/template-1/preview?appId=DiligenceStudio_WestMonroe',
        templateId: 'template-1',
      })],
    })
    expect(fetchMock).toHaveBeenCalledWith('/api/v1/templates?kind=diagram', {
      headers: { 'X-App-Id': 'DiligenceStudio_WestMonroe' },
      signal: undefined,
    })
  })

  it('rejects malformed catalog and canvas responses', async () => {
    vi.stubGlobal('fetch', vi.fn()
      .mockResolvedValueOnce(jsonResponse({ templates: [{ templateId: 4 }] }))
      .mockResolvedValueOnce(jsonResponse({ presentation: { slides: [] } })))

    await expect(listTemplates('diagram')).rejects.toBeInstanceOf(TemplateApiError)
    await expect(getTemplate('template-1')).rejects.toMatchObject({ code: 'invalid_api_response' })
    expect(fetch).toHaveBeenNthCalledWith(2, '/api/v1/templates/template-1', {
      headers: { 'X-App-Id': 'DiligenceStudio_WestMonroe' },
      signal: undefined,
    })
  })

  it('returns the imported template ID and preview status from headers', async () => {
    const fetchMock = vi.fn(async () => jsonResponse(
      { presentation: { slides: [{}] } },
      {
        status: 201,
        headers: {
          'X-Template-Id': 'imported-template',
          'X-Template-Preview-Status': 'ready',
        },
      },
    ))
    vi.stubGlobal('fetch', fetchMock)
    const file = new File(['pptx'], 'template.pptx', {
      type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    })

    await expect(importTemplate(file, 'commentary')).resolves.toMatchObject({
      previewStatus: 'ready',
      templateId: 'imported-template',
    })
    expect(fetchMock).toHaveBeenCalledWith('/api/v1/import?kind=commentary', {
      body: file,
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        'X-App-Id': 'DiligenceStudio_WestMonroe',
      },
      method: 'POST',
      signal: undefined,
    })
  })

  it('validates batch-import results and sends the deck to the batch endpoint', async () => {
    const fetchMock = vi.fn(async () => jsonResponse({
      templates: [
        {
          previewAvailable: true,
          templateId: 'batch-template-1',
          templateJson: { presentation: { slides: [{}] } },
        },
        {
          previewAvailable: false,
          templateId: 'batch-template-2',
          templateJson: { presentation: { slides: [{}] } },
        },
      ],
      warnings: ['Slide 2 preview is unavailable.'],
    }, { status: 201 }))
    vi.stubGlobal('fetch', fetchMock)
    const file = new File(['pptx'], 'templates.pptx', {
      type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    })

    await expect(batchImportTemplates(file, 'diagram')).resolves.toMatchObject({
      templates: [
        { previewStatus: 'ready', templateId: 'batch-template-1' },
        { previewStatus: 'unavailable', templateId: 'batch-template-2' },
      ],
      warnings: ['Slide 2 preview is unavailable.'],
    })
    expect(fetchMock).toHaveBeenCalledWith('/api/v1/batchImport?kind=diagram', {
      body: file,
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        'X-App-Id': 'DiligenceStudio_WestMonroe',
      },
      method: 'POST',
      signal: undefined,
    })
  })

  it('rejects malformed batch-import responses', async () => {
    vi.stubGlobal('fetch', vi.fn(async () => jsonResponse({ templates: [], warnings: [] }, { status: 201 })))
    const file = new File(['pptx'], 'templates.pptx')

    await expect(batchImportTemplates(file, 'commentary')).rejects.toMatchObject({
      code: 'invalid_api_response',
    })
  })

  it('deletes an encoded template ID and requires the 204 response contract', async () => {
    const fetchMock = vi.fn()
      .mockResolvedValueOnce(new Response(null, { status: 204 }))
      .mockResolvedValueOnce(new Response(null, { status: 200 }))
    vi.stubGlobal('fetch', fetchMock)

    await expect(deleteTemplate('template/one')).resolves.toBeUndefined()
    expect(fetchMock).toHaveBeenNthCalledWith(1, '/api/v1/templates/template%2Fone', {
      headers: { 'X-App-Id': 'DiligenceStudio_WestMonroe' },
      method: 'DELETE',
      signal: undefined,
    })
    await expect(deleteTemplate('template-two')).rejects.toMatchObject({
      code: 'invalid_api_response',
    })
  })

  it('uses versioned URLs for PowerPoint export and insertion', async () => {
    const fetchMock = vi.fn()
      .mockResolvedValueOnce(powerPointResponse())
      .mockResolvedValueOnce(powerPointResponse())
    vi.stubGlobal('fetch', fetchMock)
    const presentation = { presentation: { slides: [{}] } }
    const target = new File(['pptx'], 'target.pptx')

    await exportPresentation(presentation)
    await insertPresentation(presentation, target, 2)

    expect(fetchMock).toHaveBeenNthCalledWith(1, '/api/v1/export', {
      body: JSON.stringify(presentation),
      headers: { 'Content-Type': 'application/json', 'X-App-Id': 'DiligenceStudio_WestMonroe' },
      method: 'POST',
      signal: undefined,
    })
    expect(fetchMock).toHaveBeenNthCalledWith(2, '/api/v1/export/insert', {
      body: expect.any(FormData),
      headers: { 'X-App-Id': 'DiligenceStudio_WestMonroe' },
      method: 'POST',
      signal: undefined,
    })
  })
})

function jsonResponse(body: unknown, init: ResponseInit = {}) {
  const headers = new Headers(init.headers)
  headers.set('Content-Type', 'application/json')
  return new Response(JSON.stringify(body), { ...init, headers })
}

function powerPointResponse() {
  return new Response(new Uint8Array([0x50, 0x4b]), {
    headers: {
      'Content-Disposition': 'attachment; filename="presentation.pptx"',
      'Content-Type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    },
  })
}
