// @vitest-environment node

import { mkdtemp, mkdir, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import path from 'node:path'

import { afterEach, describe, expect, it } from 'vitest'

import { seedBuiltinTemplates } from '../src/catalog/seedBuiltinTemplates'
import type { PowerPointCanvasJson } from '../src/lib/import/PowerpointImportTypes'
import { SqliteTemplateRepository } from '../src/repositories/SqliteTemplateRepository'
import {
  HeadlessTemplatePreviewGenerator,
  QuickLookTemplatePreviewGenerator,
  readPngDimensions,
} from '../src/services/TemplatePreview'

const ONE_PIXEL_PNG = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=',
  'base64',
)
const EMPTY_TEMPLATE: PowerPointCanvasJson = {
  presentation: {
    preserveElementOrder: true,
    showBranding: false,
    slides: [{
      backgroundColor: 'FFFFFF',
      elements: [],
      height: 720,
      id: 'slide-1',
      name: 'Preview test',
      width: 1280,
    }],
    title: 'Preview test',
  },
}

const repositories: SqliteTemplateRepository[] = []
const temporaryDirectories: string[] = []

afterEach(async () => {
  repositories.splice(0).forEach((repository) => repository.close())
  await Promise.all(
    temporaryDirectories.splice(0).map((directory) => rm(directory, { force: true, recursive: true })),
  )
})

describe('built-in catalog', () => {
  it('seeds stable diagram and commentary templates idempotently with previews', async () => {
    const repository = new SqliteTemplateRepository(':memory:')
    repositories.push(repository)

    await seedBuiltinTemplates(repository)
    await seedBuiltinTemplates(repository)

    expect(repository.list('diagram')).toHaveLength(6)
    expect(repository.list('commentary')).toHaveLength(3)
    expect(repository.findPreview('layered-platform')).toMatchObject({
      contentType: 'image/png',
      templateId: 'layered-platform',
    })
    const commentary = repository.findById('security-ssa')
    expect(JSON.stringify(commentary)).not.toContain('element-903000.jpg')
    expect(JSON.stringify(commentary)).not.toContain('element-5.png')
  })

  it('registers app metadata and keeps built-in access independently removable', async () => {
    const repository = new SqliteTemplateRepository(':memory:')
    repositories.push(repository)
    await seedBuiltinTemplates(repository)

    repository.ensureApp('Quarry_WestMonroe', {
      displayName: 'Quarry_WestMonroe',
      metadata: { projectPath: '/Users/rgambhir/Quarry' },
    })

    expect(repository.findApp('Quarry_WestMonroe')).toMatchObject({
      appId: 'Quarry_WestMonroe',
      displayName: 'Quarry_WestMonroe',
      metadata: { projectPath: '/Users/rgambhir/Quarry' },
    })
    expect(repository.list('diagram', 'Quarry_WestMonroe')).toHaveLength(6)
    expect(repository.delete('layered-platform', 'Quarry_WestMonroe')).toBe(true)
    repository.ensureApp('Quarry_WestMonroe')
    expect(repository.findById('layered-platform', 'Quarry_WestMonroe')).toBeUndefined()
    expect(repository.findById('layered-platform')).toBeDefined()
  })
})

describe('template preview validation', () => {
  it('accepts the single PNG created by the injected Quick Look runner', async () => {
    const directory = await mkdtemp(path.join(tmpdir(), 'tts-preview-test-'))
    temporaryDirectories.push(directory)
    const outputDirectory = path.join(directory, 'preview')
    await mkdir(outputDirectory)
    const generator = new QuickLookTemplatePreviewGenerator(
      { maxBytes: 1024, maxDimension: 100, renderSize: 1600, timeoutMs: 1000 },
      async (inputPath, receivedOutputDirectory, options) => {
        expect(inputPath).toBe(path.join(directory, 'upload.pptx'))
        expect(receivedOutputDirectory).toBe(outputDirectory)
        expect(options.renderSize).toBe(1600)
        await writeFile(path.join(outputDirectory, 'upload.pptx.png'), ONE_PIXEL_PNG)
      },
    )

    await expect(generator.generate({
      inputPath: path.join(directory, 'upload.pptx'),
      outputDirectory,
      templateJson: EMPTY_TEMPLATE,
    }))
      .resolves.toMatchObject({ bytes: ONE_PIXEL_PNG, height: 1, width: 1 })
  })

  it('rejects malformed PNG data and extra output files', async () => {
    expect(readPngDimensions(Buffer.from('not png'))).toBeUndefined()
    const directory = await mkdtemp(path.join(tmpdir(), 'tts-preview-test-'))
    temporaryDirectories.push(directory)
    const outputDirectory = path.join(directory, 'preview')
    await mkdir(outputDirectory)
    const generator = new QuickLookTemplatePreviewGenerator(
      { maxBytes: 1024, renderSize: 1600, timeoutMs: 1000 },
      async () => {
        await writeFile(path.join(outputDirectory, 'one.png'), ONE_PIXEL_PNG)
        await writeFile(path.join(outputDirectory, 'two.png'), ONE_PIXEL_PNG)
      },
    )

    await expect(generator.generate({
      inputPath: path.join(directory, 'upload.pptx'),
      outputDirectory,
      templateJson: EMPTY_TEMPLATE,
    }))
      .resolves.toBeUndefined()
  })

  it('accepts a PNG rendered from normalized JSON by the injected headless runner', async () => {
    const generator = new HeadlessTemplatePreviewGenerator(
      {
        maxBytes: 1024,
        maxDimension: 100,
        renderSize: 80,
        renderUrl: 'http://127.0.0.1:5173/_internal/template-preview',
        timeoutMs: 1000,
      },
      async (templateJson, options) => {
        expect(templateJson).toBe(EMPTY_TEMPLATE)
        expect(options.renderSize).toBe(80)
        return ONE_PIXEL_PNG
      },
    )

    await expect(generator.generate({
      inputPath: '/unused/upload.pptx',
      outputDirectory: '/unused/preview',
      templateJson: EMPTY_TEMPLATE,
    })).resolves.toMatchObject({ bytes: ONE_PIXEL_PNG, height: 1, width: 1 })
  })

  it('serializes headless preview jobs to bound browser resource use', async () => {
    let releaseFirst: (() => void) | undefined
    const firstCanFinish = new Promise<void>((resolve) => {
      releaseFirst = resolve
    })
    let calls = 0
    const generator = new HeadlessTemplatePreviewGenerator(
      {
        maxBytes: 1024,
        renderSize: 80,
        renderUrl: 'http://127.0.0.1:5173/_internal/template-preview',
        timeoutMs: 1000,
      },
      async () => {
        calls += 1
        if (calls === 1) await firstCanFinish
        return ONE_PIXEL_PNG
      },
    )
    const request = {
      inputPath: '/unused/upload.pptx',
      outputDirectory: '/unused/preview',
      templateJson: EMPTY_TEMPLATE,
    }

    const first = generator.generate(request)
    const second = generator.generate(request)
    await Promise.resolve()
    await Promise.resolve()
    expect(calls).toBe(1)
    releaseFirst?.()
    await Promise.all([first, second])
    expect(calls).toBe(2)
  })
})
