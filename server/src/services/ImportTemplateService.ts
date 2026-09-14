import { randomUUID } from 'node:crypto'
import { mkdir, mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import path from 'node:path'

import { buildAppScopedPath, DEFAULT_APP_ID } from '../appIdentity'
import { ApiError } from '../errors'
import type {
  PowerPointCanvasElement,
  PowerPointCanvasJson,
} from '../lib/import/PowerpointImportTypes'
import {
  buildSlideClassificationInputFingerprint,
  checksumBytes,
} from '../lib/retrieval/SlideClassificationInput'
import {
  SLIDE_CLASSIFICATION_PROMPT_VERSION,
  SLIDE_CLASSIFICATION_SCHEMA_VERSION,
} from '../lib/retrieval/SlideRetrievalMetadata'
import type {
  TemplateInsert,
  TemplateKind,
  TemplateRepository,
} from '../repositories/TemplateRepository'
import type { PowerPointConverter } from './PowerPointConverter'
import {
  DisabledTemplatePreviewGenerator,
  type TemplatePreviewGenerator,
} from './TemplatePreview'
import {
  externalizeTemplateAssets,
  hydrateCanvasTemplateAssetSources,
} from './TemplateAssets'

export type TemplateListItem = {
  description: string
  elementCount: number
  kind: TemplateKind
  previewUrl: string | null
  slideCount: number
  templateId: string
  title: string
}

export type TemplateListResponse = {
  templates: TemplateListItem[]
}

export type TemplatePreviewListResponse = {
  pagination: {
    hasNextPage: boolean
    hasPreviousPage: boolean
    page: number
    pageSize: number
    totalItems: number
    totalPages: number
  }
  previews: Array<{
    contentType: 'image/png'
    dataUrl: string
    height: number
    previewUrl: string
    templateId: string
    width: number
  }>
}

export type BatchImportResponse = {
  templates: Array<{
    previewAvailable: boolean
    templateId: string
    templateJson: PowerPointCanvasJson
  }>
  warnings: string[]
}

export type ImportV2Response = {
  previewAvailable: boolean
  retrieval: { status: 'pending' }
  templateId: string
  templateJson: PowerPointCanvasJson
  warnings: string[]
}

export const TEMPLATE_PREVIEW_PAGE_SIZE = 10

export class ImportService {
  constructor(
    private readonly converter: PowerPointConverter,
    private readonly templates: TemplateRepository,
    private readonly createTemplateId: () => string = randomUUID,
    private readonly createAssetId: () => string = randomUUID,
    private readonly previewGenerator: TemplatePreviewGenerator = new DisabledTemplatePreviewGenerator(),
    private readonly notifyClassificationWork: () => void = () => undefined,
  ) {}

  async import(
    source: Buffer,
    kind: TemplateKind = 'diagram',
    signal?: AbortSignal,
    appId: string = DEFAULT_APP_ID,
  ) {
    return this.#importSingle(source, kind, signal, appId, false)
  }

  /** Import one slide and atomically persist its template plus durable classification work. */
  async importV2(
    source: Buffer,
    kind: TemplateKind = 'diagram',
    signal?: AbortSignal,
    appId: string = DEFAULT_APP_ID,
  ): Promise<ImportV2Response> {
    const result = await this.#importSingle(source, kind, signal, appId, true)
    try {
      this.notifyClassificationWork()
    } catch {
      // Notification is best-effort; durable polling will recover the committed job.
    }
    return { ...result, retrieval: { status: 'pending' } }
  }

  async #importSingle(
    source: Buffer,
    kind: TemplateKind,
    signal: AbortSignal | undefined,
    appId: string,
    createClassification: boolean,
  ) {
    this.templates.ensureApp(appId)
    const workingDirectory = await mkdtemp(path.join(tmpdir(), 'diligence-studio-import-'))
    const inputPath = path.join(workingDirectory, 'upload.pptx')
    const outputPath = path.join(workingDirectory, 'upload.canvas.json')
    const previewDirectory = path.join(workingDirectory, 'preview')

    try {
      await writeFile(inputPath, source)
      await mkdir(previewDirectory)
      const conversion = await this.converter.convertFile(inputPath, outputPath)
      if (conversion.templateJson.presentation.slides.length !== 1) {
        throw new ApiError(
          422,
          'template_must_have_one_slide',
          'Template PowerPoint files must contain exactly one slide.',
        )
      }
      const { templateJson: repairedTemplateJson } = repairCanvasDimensions(conversion.templateJson)
      const preview = await this.previewGenerator.generate({
        inputPath,
        outputDirectory: previewDirectory,
        templateJson: repairedTemplateJson,
      }, signal)
      if (signal?.aborted) {
        throw new ApiError(499, 'request_cancelled', 'The template import was cancelled.')
      }
      const warnings = preview
        ? conversion.warnings
        : [...conversion.warnings, 'A preview image could not be generated for this template.']
      const templateId = this.createTemplateId()
      const externalized = externalizeTemplateAssets(
        templateId,
        repairedTemplateJson,
        this.createAssetId,
      )
      const template = {
        templateId,
        templateJson: externalized.templateJson,
      }
      const record: TemplateInsert = {
        assets: externalized.assets,
        metadata: {
          checksum: null,
          createdAt: new Date().toISOString(),
          description: 'Imported PowerPoint template',
          kind,
          source: 'import',
          templateId,
        },
        preview: preview ? { ...preview, templateId } : undefined,
        template,
      }
      if (createClassification) {
        this.templates.insertWithPendingClassification(record, {
          inputFingerprint: buildSlideClassificationInputFingerprint(
            externalized.templateJson,
            preview ? checksumBytes(preview.bytes) : null,
          ),
          promptVersion: SLIDE_CLASSIFICATION_PROMPT_VERSION,
          schemaVersion: SLIDE_CLASSIFICATION_SCHEMA_VERSION,
        }, appId)
      } else {
        this.templates.insert(
          template,
          externalized.assets,
          record.metadata,
          record.preview,
          appId,
        )
      }

      return {
        previewAvailable: preview !== undefined,
        templateId,
        templateJson: hydrateCanvasTemplateAssetSources(
          externalized.templateJson,
          externalized.assets,
        ),
        warnings,
      }
    } finally {
      await rm(workingDirectory, { force: true, recursive: true })
    }
  }

  async batchImport(
    source: Buffer,
    kind: TemplateKind = 'diagram',
    signal?: AbortSignal,
    appId: string = DEFAULT_APP_ID,
  ): Promise<BatchImportResponse> {
    this.templates.ensureApp(appId)
    const workingDirectory = await mkdtemp(path.join(tmpdir(), 'diligence-studio-batch-import-'))
    const inputPath = path.join(workingDirectory, 'upload.pptx')
    const outputPath = path.join(workingDirectory, 'upload.canvas.json')

    try {
      await writeFile(inputPath, source)
      const conversion = await this.converter.convertFile(inputPath, outputPath)
      if (conversion.templateJson.presentation.slides.length === 0) {
        throw new ApiError(
          422,
          'powerpoint_has_no_slides',
          'The PowerPoint file must contain at least one slide.',
        )
      }

      const warnings = [...conversion.warnings]
      const records: TemplateInsert[] = []
      const templates: BatchImportResponse['templates'] = []
      const createdAt = new Date().toISOString()

      for (const [index, slide] of conversion.templateJson.presentation.slides.entries()) {
        if (signal?.aborted) {
          throw new ApiError(499, 'request_cancelled', 'The template import was cancelled.')
        }

        const templateJson: PowerPointCanvasJson = {
          presentation: {
            ...conversion.templateJson.presentation,
            title: slide.name,
            slides: [slide],
          },
        }
        const { templateJson: repairedTemplateJson } = repairCanvasDimensions(templateJson)
        const previewDirectory = path.join(workingDirectory, `preview-${index + 1}`)
        await mkdir(previewDirectory)
        const preview = index > 0 && this.previewGenerator.supportsIndependentSlides === false
          ? undefined
          : await this.previewGenerator.generate({
              inputPath,
              outputDirectory: previewDirectory,
              templateJson: repairedTemplateJson,
            }, signal)
        if (signal?.aborted) {
          throw new ApiError(499, 'request_cancelled', 'The template import was cancelled.')
        }
        if (!preview) {
          warnings.push(`Slide ${index + 1}: A preview image could not be generated for this template.`)
        }

        const templateId = this.createTemplateId()
        const externalized = externalizeTemplateAssets(
          templateId,
          repairedTemplateJson,
          this.createAssetId,
        )
        records.push({
          assets: externalized.assets,
          metadata: {
            checksum: null,
            createdAt,
            description: 'Imported PowerPoint template',
            kind,
            source: 'import',
            templateId,
          },
          preview: preview ? { ...preview, templateId } : undefined,
          template: { templateId, templateJson: externalized.templateJson },
        })
        templates.push({
          previewAvailable: preview !== undefined,
          templateId,
          templateJson: hydrateCanvasTemplateAssetSources(
            externalized.templateJson,
            externalized.assets,
          ),
        })
      }

      this.templates.insertMany(records, appId)
      return { templates, warnings }
    } finally {
      await rm(workingDirectory, { force: true, recursive: true })
    }
  }

  find(templateId: string, appId: string = DEFAULT_APP_ID) {
    this.templates.ensureApp(appId)
    const storedTemplate = this.templates.findByIdWithAssets(templateId, appId)
    if (!storedTemplate) {
      return undefined
    }

    const externalized = externalizeTemplateAssets(
      templateId,
      storedTemplate.templateJson,
      this.createAssetId,
    )
    const repaired = repairCanvasDimensions(externalized.templateJson)
    let templateJson = repaired.templateJson
    let assets = storedTemplate.assets
    if (repaired.repaired || externalized.assets.length > 0) {
      assets = [...assets, ...externalized.assets]
      this.templates.update({ templateId, templateJson }, externalized.assets, appId)
    }

    return {
      templateId,
      templateJson: hydrateCanvasTemplateAssetSources(templateJson, assets),
    }
  }

  findAsset(templateId: string, assetId: string, appId: string = DEFAULT_APP_ID) {
    this.templates.ensureApp(appId)
    return this.templates.findAsset(templateId, assetId, appId)
  }

  findPreview(templateId: string, appId: string = DEFAULT_APP_ID) {
    this.templates.ensureApp(appId)
    return this.templates.findPreview(templateId, appId)
  }

  list(kind?: TemplateKind, appId: string = DEFAULT_APP_ID): TemplateListResponse {
    this.templates.ensureApp(appId)
    return {
      templates: this.templates.list(kind, appId).map(({ metadata, previewAvailable, templateId, templateJson }) => ({
        description: metadata.description,
        kind: metadata.kind,
        previewUrl: previewAvailable
          ? buildAppScopedPath(`/templates/${templateId}/preview`, appId)
          : null,
        templateId,
        title: templateJson.presentation.title,
        slideCount: templateJson.presentation.slides.length,
        elementCount: templateJson.presentation.slides.reduce(
          (count, slide) => count + slide.elements.length,
          0,
        ),
      })),
    }
  }

  listPreviews(page: number, appId: string = DEFAULT_APP_ID): TemplatePreviewListResponse {
    this.templates.ensureApp(appId)
    const offset = (page - 1) * TEMPLATE_PREVIEW_PAGE_SIZE
    const { previews, total } = this.templates.listPreviews(
      TEMPLATE_PREVIEW_PAGE_SIZE,
      offset,
      appId,
    )

    return {
      pagination: {
        hasNextPage: offset + previews.length < total,
        hasPreviousPage: page > 1 && total > 0,
        page,
        pageSize: TEMPLATE_PREVIEW_PAGE_SIZE,
        totalItems: total,
        totalPages: Math.ceil(total / TEMPLATE_PREVIEW_PAGE_SIZE),
      },
      previews: previews.map(({ bytes, contentType, height, templateId, width }) => ({
        contentType,
        dataUrl: `data:${contentType};base64,${bytes.toString('base64')}`,
        height,
        previewUrl: buildAppScopedPath(`/templates/${templateId}/preview`, appId),
        templateId,
        width,
      })),
    }
  }

  delete(templateId: string, appId: string = DEFAULT_APP_ID) {
    this.templates.ensureApp(appId)
    return this.templates.delete(templateId, appId)
  }
}

export { ImportService as ImportTemplateService }

function repairCanvasDimensions(templateJson: PowerPointCanvasJson): {
  repaired: boolean
  templateJson: PowerPointCanvasJson
} {
  let repaired = false
  const positiveDimension = (value: number) => {
    if (value > 0) {
      return value
    }
    repaired = true
    return 1
  }
  const repairElement = (element: PowerPointCanvasElement): PowerPointCanvasElement => {
    if (element.type === 'line') {
      return element
    }

    const w = positiveDimension(element.w)
    const h = positiveDimension(element.h)
    return w === element.w && h === element.h ? element : { ...element, w, h }
  }

  const slides = templateJson.presentation.slides.map((slide) => {
    const width = positiveDimension(slide.width)
    const height = positiveDimension(slide.height)
    const elements = slide.elements.map(repairElement)
    return width === slide.width
      && height === slide.height
      && elements.every((element, index) => element === slide.elements[index])
      ? slide
      : { ...slide, width, height, elements }
  })

  return {
    repaired,
    templateJson: repaired
      ? {
          ...templateJson,
          presentation: { ...templateJson.presentation, slides },
        }
      : templateJson,
  }
}
