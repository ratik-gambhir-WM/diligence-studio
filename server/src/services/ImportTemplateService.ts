import { randomUUID } from 'node:crypto'

import { buildAppScopedPath, DEFAULT_APP_ID } from '../appIdentity'
import { ApiError } from '../errors'
import { SlideProviderError } from '../integrations/SlideProvider'
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
import type { SlideClassificationService } from './SlideClassificationService'
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
  retrieval: { status: 'ready' }
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
    private readonly classificationService?: SlideClassificationService,
  ) {}

  async import(
    source: Buffer,
    kind: TemplateKind = 'diagram',
    signal?: AbortSignal,
    appId: string = DEFAULT_APP_ID,
  ) {
    return this.#importSingle(source, kind, signal, appId, false)
  }

  /** Import one slide and return only after classification and embedding are ready in memory. */
  async importV2(
    source: Buffer,
    kind: TemplateKind = 'diagram',
    signal?: AbortSignal,
    appId: string = DEFAULT_APP_ID,
  ): Promise<ImportV2Response> {
    if (!this.classificationService) {
      throw new ApiError(
        503,
        'slide_classification_unavailable',
        'Synchronous slide classification is not configured.',
      )
    }
    const result = await this.#importSingle(source, kind, signal, appId, true)
    try {
      if (!result.classificationJob) {
        throw new Error('The v2 import did not create synchronous classification input.')
      }
      await this.classificationService.process(result.classificationJob, signal)
    } catch (error) {
      const errorCode = signal?.aborted
        ? 'request_cancelled'
        : error instanceof SlideProviderError ? error.code : 'classification_internal'
      this.templates.recordSlideClassificationFailure(result.templateId, errorCode, null)
      throw toSynchronousProcessingError(error, signal)
    }
    return {
      previewAvailable: result.previewAvailable,
      retrieval: { status: 'ready' },
      templateId: result.templateId,
      templateJson: result.templateJson,
      warnings: result.warnings,
    }
  }

  async #importSingle(
    source: Buffer,
    kind: TemplateKind,
    signal: AbortSignal | undefined,
    appId: string,
    createClassification: boolean,
  ) {
    this.templates.ensureApp(appId)
    const conversion = await this.converter.convert(source)
    if (conversion.templateJson.presentation.slides.length !== 1) {
      throw new ApiError(
        422,
        'template_must_have_one_slide',
        'Template PowerPoint files must contain exactly one slide.',
      )
    }
    const { templateJson: repairedTemplateJson } = repairCanvasDimensions(conversion.templateJson)
    const preview = await this.previewGenerator.generate({
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
      const inputFingerprint = buildSlideClassificationInputFingerprint(
        externalized.templateJson,
        preview ? checksumBytes(preview.bytes) : null,
      )
      this.templates.insertWithProcessingClassification(record, {
        inputFingerprint,
        promptVersion: SLIDE_CLASSIFICATION_PROMPT_VERSION,
        schemaVersion: SLIDE_CLASSIFICATION_SCHEMA_VERSION,
      }, appId)
      return {
        classificationJob: {
          attemptCount: 1,
          inputFingerprint,
          kind,
          preview: record.preview,
          template,
          title: externalized.templateJson.presentation.title,
        },
        previewAvailable: preview !== undefined,
        templateId,
        templateJson: hydrateCanvasTemplateAssetSources(
          externalized.templateJson,
          externalized.assets,
        ),
        warnings,
      }
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
      classificationJob: undefined,
      previewAvailable: preview !== undefined,
      templateId,
      templateJson: hydrateCanvasTemplateAssetSources(
        externalized.templateJson,
        externalized.assets,
      ),
      warnings,
    }
  }

  async batchImport(
    source: Buffer,
    kind: TemplateKind = 'diagram',
    signal?: AbortSignal,
    appId: string = DEFAULT_APP_ID,
  ): Promise<BatchImportResponse> {
    this.templates.ensureApp(appId)
    const conversion = await this.converter.convert(source)
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

    // TODO: Add Concurrency here
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
      const preview = await this.previewGenerator.generate({
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

function toSynchronousProcessingError(error: unknown, signal?: AbortSignal) {
  if (signal?.aborted) {
    return new ApiError(499, 'request_cancelled', 'The template import was cancelled.')
  }
  if (!(error instanceof SlideProviderError)) return error
  if (error.code.endsWith('_timeout')) {
    return new ApiError(504, error.code, 'Slide retrieval processing timed out.')
  }
  if (error.code.endsWith('_rate_limited') || error.code.endsWith('_unavailable')) {
    return new ApiError(503, error.code, 'Slide retrieval processing is temporarily unavailable.')
  }
  return new ApiError(502, error.code, 'Slide retrieval processing failed.')
}

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
