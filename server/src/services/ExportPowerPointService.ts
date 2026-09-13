import {
  buildSuggestedFileName,
  buildThemedPptxBytes,
  normalizePresentationSpec,
} from '../lib/export/PowerpointGenerator'
import { DEFAULT_APP_ID } from '../appIdentity'
import type {
  NormalizedPresentation,
} from '../lib/shared/PowerpointTypes'
import { ApiError } from '../errors'
import { insertPptxBytesIntoExistingDeck } from '../lib/export/PowerpointInsertion'
import type { TemplateRepository } from '../repositories/TemplateRepository'
import { decodeDataImage, hydrateTemplateAssetSources } from './TemplateAssets'
import { normalizePresentation } from './NormalizePresentation'
import { validatePowerPointPackage } from './PowerPointConverter'

const MAX_JSON_DEPTH = 100
const MAX_JSON_NODES = 100_000
const MAX_EMBEDDED_IMAGE_BYTES = 40 * 1024 * 1024

export type ExportedPowerPoint = {
  bytes: Uint8Array
  fileName: string
  warnings: string[]
}

export interface ExportPowerPointUseCase {
  export(input: unknown, appId?: string): Promise<ExportedPowerPoint>
  insert(
    input: unknown,
    targetDeck: Buffer,
    targetFileName: string,
    insertAfterSlide: number,
    appId?: string,
  ): Promise<ExportedPowerPoint>
}

export class ExportService implements ExportPowerPointUseCase {
  constructor(private readonly templates: TemplateRepository) {}

  async export(input: unknown, appId: string = DEFAULT_APP_ID): Promise<ExportedPowerPoint> {
    this.templates.ensureApp(appId)
    assertBoundedJsonValue(input)
    const normalized = normalizePresentation(input)
    const { presentation, issues } = normalizePresentationSpec(normalized.templateJson)
    const errors = issues.filter((issue) => issue.level === 'error')

    if (!presentation || errors.length > 0) {
      throw new ApiError(
        422,
        'invalid_presentation_json',
        'The JSON could not be converted into a PowerPoint presentation.',
      )
    }

    const hydratedPresentation = hydrateTemplateAssetSources(presentation, this.templates, appId)
    validateImageSources(hydratedPresentation)
    const bytes = await buildThemedPptxBytes(hydratedPresentation)

    return {
      bytes,
      fileName: buildSuggestedFileName(hydratedPresentation),
      warnings: normalized.warnings,
    }
  }

  async insert(
    input: unknown,
    targetDeck: Buffer,
    targetFileName: string,
    insertAfterSlide: number,
    appId: string = DEFAULT_APP_ID,
  ): Promise<ExportedPowerPoint> {
    if (!Number.isSafeInteger(insertAfterSlide) || insertAfterSlide < 0) {
      throw new ApiError(
        400,
        'invalid_insert_position',
        'insertAfterSlide must be a non-negative whole number.',
      )
    }
    try {
      validatePowerPointPackage(targetDeck)
      const generated = await this.export(input, appId)
      const inserted = await insertPptxBytesIntoExistingDeck(
        generated.bytes,
        targetDeck,
        targetFileName,
        insertAfterSlide,
      )
      return { ...inserted, warnings: generated.warnings }
    } catch (error) {
      if (error instanceof ApiError) {
        throw error
      }
      throw new ApiError(
        422,
        'invalid_target_powerpoint',
        'The target file is not a supported PowerPoint presentation.',
        { cause: error },
      )
    }
  }
}

function assertBoundedJsonValue(input: unknown): void {
  const pending: Array<{ depth: number; value: unknown }> = [{ depth: 0, value: input }]
  let nodeCount = 0

  while (pending.length > 0) {
    const current = pending.pop()
    if (!current) {
      continue
    }

    nodeCount += 1
    if (nodeCount > MAX_JSON_NODES || current.depth > MAX_JSON_DEPTH) {
      throw new ApiError(
        422,
        'presentation_too_complex',
        'The presentation JSON exceeds the processing complexity limit.',
      )
    }

    const value = current.value
    if (
      value === null
      || typeof value === 'string'
      || typeof value === 'boolean'
      || (typeof value === 'number' && Number.isFinite(value))
    ) {
      continue
    }

    if (Array.isArray(value)) {
      for (const entry of value) {
        pending.push({ depth: current.depth + 1, value: entry })
      }
      continue
    }

    if (isPlainRecord(value)) {
      for (const entry of Object.values(value)) {
        pending.push({ depth: current.depth + 1, value: entry })
      }
      continue
    }

    throw new ApiError(
      400,
      'invalid_json_value',
      'The request body contains an unsupported JSON value.',
    )
  }
}

export { ExportService as ExportPowerPointService }

function validateImageSources(presentation: NormalizedPresentation) {
  let embeddedBytes = 0

  for (const slide of presentation.slides) {
    for (const element of slide.elements) {
      if (element.kind !== 'image') {
        continue
      }

      let imageBytes: number
      try {
        imageBytes = decodeDataImage(element.src).bytes.length
      } catch {
        throw new ApiError(
          422,
          'unsupported_image_source',
          'Exported images must use an embedded base64 data URI.',
        )
      }

      embeddedBytes += imageBytes
      if (embeddedBytes > MAX_EMBEDDED_IMAGE_BYTES) {
        throw new ApiError(
          413,
          'embedded_images_too_large',
          'The embedded presentation images exceed the processing limit.',
        )
      }
    }
  }
}

function isPlainRecord(value: unknown): value is Record<string, unknown> {
  if (typeof value !== 'object' || value === null) {
    return false
  }
  const prototype = Object.getPrototypeOf(value)
  return prototype === Object.prototype || prototype === null
}
