// @vitest-environment node

import { afterEach, describe, expect, it, vi } from 'vitest'

import {
  SharePointResourceService,
} from '../src/services/SharePointResourceService'

const RESOURCE_URL = 'https://relentlessblue.sharepoint.com/:b:/s/TestSite/test-file?e=abc'

describe('SharePointResourceService', () => {
  afterEach(() => {
    vi.restoreAllMocks()
  })

  it('resolves a shared file using the Graph shares endpoint', async () => {
    const fetchMock = vi.spyOn(globalThis, 'fetch').mockResolvedValue(
      jsonResponse({
        driveItem: {
          file: { mimeType: 'application/pdf' },
          id: 'file-1',
          lastModifiedDateTime: '2026-09-23T00:00:00Z',
          name: 'Test diagram.pdf',
          parentReference: { driveId: 'drive-1' },
          size: 123,
          webUrl: 'https://relentlessblue.sharepoint.com/sites/TestSite/test-file',
        },
      }),
    )
    const service = new SharePointResourceService()

    const result = await service.resolveResource(RESOURCE_URL, 'access-token')

    expect(result).toMatchObject({ kind: 'file', name: 'Test diagram.pdf' })
    expect(result.files).toEqual([expect.objectContaining({
      driveId: 'drive-1',
      fileId: 'file-1',
      mimeType: 'application/pdf',
      name: 'Test diagram.pdf',
    })])
    expect(fetchMock).toHaveBeenCalledWith(
      expect.stringMatching(/^https:\/\/graph\.microsoft\.com\/v1\.0\/shares\/u!.*\/driveItem\?/u),
      expect.objectContaining({
        headers: expect.objectContaining({ Authorization: 'Bearer access-token' }),
      }),
    )
  })

  it('lists nested files for a shared folder', async () => {
    const fetchMock = vi.spyOn(globalThis, 'fetch')
      .mockResolvedValueOnce(jsonResponse({
        driveItem: {
          folder: {},
          id: 'folder-1',
          name: 'Architecture',
          parentReference: { driveId: 'drive-1' },
          webUrl: 'https://relentlessblue.sharepoint.com/sites/TestSite/Architecture',
        },
      }))
      .mockResolvedValueOnce(jsonResponse({
        value: [
          {
            file: { mimeType: 'application/pdf' },
            id: 'file-1',
            lastModifiedDateTime: '2026-09-23T00:00:00Z',
            name: 'diagram.pdf',
            parentReference: { driveId: 'drive-1' },
            size: 123,
            webUrl: 'https://relentlessblue.sharepoint.com/sites/TestSite/Architecture/diagram.pdf',
          },
          {
            folder: {},
            id: 'nested-1',
            name: 'Archive',
            parentReference: { driveId: 'drive-1' },
            webUrl: 'https://relentlessblue.sharepoint.com/sites/TestSite/Architecture/Archive',
          },
        ],
      }))
      .mockResolvedValueOnce(jsonResponse({
        value: [{
          file: { mimeType: 'text/plain' },
          id: 'file-2',
          lastModifiedDateTime: '2026-09-23T00:00:00Z',
          name: 'notes.txt',
          parentReference: { driveId: 'drive-1' },
          size: 10,
          webUrl: 'https://relentlessblue.sharepoint.com/sites/TestSite/Architecture/Archive/notes.txt',
        }],
      }))
    const service = new SharePointResourceService()

    const result = await service.resolveResource(
      'https://relentlessblue.sharepoint.com/:f:/s/TestSite/test-folder?e=abc',
      'access-token',
    )

    expect(result.kind).toBe('folder')
    expect(result.files.map((file) => file.path)).toEqual(['Architecture', 'Architecture/Archive'])
    expect(fetchMock).toHaveBeenCalledTimes(3)
  })

  it('downloads file content through Graph without exposing the Graph URL', async () => {
    const fetchMock = vi.spyOn(globalThis, 'fetch')
      .mockResolvedValueOnce(jsonResponse({
        file: { mimeType: 'application/pdf' },
        id: 'file-1',
        name: 'diagram.pdf',
        parentReference: { driveId: 'drive-1' },
        size: 4,
        webUrl: 'https://relentlessblue.sharepoint.com/sites/TestSite/diagram.pdf',
      }))
      .mockResolvedValueOnce(new Response(new Uint8Array([1, 2, 3, 4]), {
        headers: { 'Content-Type': 'application/pdf' },
        status: 200,
      }))
    const service = new SharePointResourceService()

    const result = await service.downloadFile('drive-1', 'file-1', 'Bearer access-token')

    expect(result.fileName).toBe('diagram.pdf')
    expect(result.contentType).toBe('application/pdf')
    expect([...result.bytes]).toEqual([1, 2, 3, 4])
    expect(fetchMock.mock.calls[1]?.[0]).toBe(
      'https://graph.microsoft.com/v1.0/drives/drive-1/items/file-1/content',
    )
    expect(fetchMock.mock.calls[1]?.[1]).toMatchObject({
      headers: expect.objectContaining({ Authorization: 'Bearer access-token' }),
    })
  })

  it('rejects links outside the configured SharePoint hosts', async () => {
    const service = new SharePointResourceService()

    await expect(service.resolveResource('https://example.com/file.pdf', 'access-token'))
      .rejects.toMatchObject({
        code: 'unsupported_sharepoint_host',
        statusCode: 400,
      })
  })
})

function jsonResponse(value: unknown) {
  return new Response(JSON.stringify(value), {
    headers: { 'Content-Type': 'application/json' },
    status: 200,
  })
}
