import type {
  PowerPointCanvasImageElement,
  PowerPointCanvasJson,
} from '../lib/import/PowerpointImportTypes'
import type { NormalizedPresentation } from '../lib/shared/PowerpointTypes'
import { ApiError } from '../errors'
import type { TemplateAsset, TemplateRepository } from '../repositories/TemplateRepository'

const TEMPLATE_ASSET_PATH_PATTERN =
  /^\/import\/([A-Za-z0-9_-]{1,128})\/assets\/([A-Za-z0-9_-]{1,128})$/
const SUPPORTED_IMAGE_CONTENT_TYPES = new Set([
  'image/png',
  'image/jpeg',
  'image/jpg',
  'image/gif',
  'image/svg+xml',
  'image/emf',
])

export function externalizeTemplateAssets(
  templateId: string,
  templateJson: PowerPointCanvasJson,
  createAssetId: () => string,
) {
  const assets: TemplateAsset[] = []
  const assetBySource = new Map<string, TemplateAsset>()

  const slides = templateJson.presentation.slides.map((slide) => ({
    ...slide,
    elements: slide.elements.map((element) => {
      if (element.type !== 'image') {
        return element
      }
      if (parseTemplateAssetPath(element.src)) {
        return element
      }

      const existingAsset = assetBySource.get(element.src)
      const asset = existingAsset ?? createTemplateAsset(templateId, element, createAssetId())
      if (!existingAsset) {
        assetBySource.set(element.src, asset)
        assets.push(asset)
      }

      return {
        ...element,
        src: buildTemplateAssetPath(asset.templateId, asset.assetId),
      }
    }),
  }))

  return {
    assets,
    templateJson: {
      ...templateJson,
      presentation: {
        ...templateJson.presentation,
        slides,
      },
    },
  }
}

export function hydrateTemplateAssetSources(
  presentation: NormalizedPresentation,
  templates: TemplateRepository,
  appId: string,
) {
  return {
    ...presentation,
    slides: presentation.slides.map((slide) => ({
      ...slide,
      elements: slide.elements.map((element) => {
        if (element.kind !== 'image') {
          return element
        }

        const reference = parseTemplateAssetPath(element.src)
        if (!reference) {
          return element
        }

        const asset = templates.findAsset(reference.templateId, reference.assetId, appId)
        if (!asset) {
          throw new ApiError(
            422,
            'template_asset_not_found',
            'A referenced template image does not exist.',
          )
        }

        return { ...element, src: templateAssetDataUri(asset) }
      }),
    })),
  }
}

export function hydrateCanvasTemplateAssetSources(
  templateJson: PowerPointCanvasJson,
  assets: readonly TemplateAsset[],
): PowerPointCanvasJson {
  const assetsByReference = new Map(
    assets.map((asset) => [templateAssetKey(asset.templateId, asset.assetId), asset]),
  )

  return {
    ...templateJson,
    presentation: {
      ...templateJson.presentation,
      slides: templateJson.presentation.slides.map((slide) => ({
        ...slide,
        elements: slide.elements.map((element) => {
          if (element.type !== 'image') {
            return element
          }

          const reference = parseTemplateAssetPath(element.src)
          if (!reference) {
            return element
          }

          const asset = assetsByReference.get(
            templateAssetKey(reference.templateId, reference.assetId),
          )
          if (!asset) {
            throw new ApiError(
              422,
              'template_asset_not_found',
              'A referenced template image does not exist.',
            )
          }

          return { ...element, src: templateAssetDataUri(asset) }
        }),
      })),
    },
  }
}

function createTemplateAsset(
  templateId: string,
  element: PowerPointCanvasImageElement,
  assetId: string,
): TemplateAsset {
  let decoded: ReturnType<typeof decodeDataImage>
  try {
    decoded = decodeDataImage(element.src)
  } catch {
    throw new ApiError(
      422,
      'unsupported_imported_image_source',
      'An imported image could not be stored safely.',
    )
  }

  if (decoded.bytes.length === 0) {
    throw new ApiError(
      422,
      'invalid_imported_image',
      'An imported image did not contain valid image data.',
    )
  }

  return {
    assetId,
    bytes: decoded.bytes,
    contentType: decoded.contentType,
    templateId,
  }
}

export function decodeDataImage(source: string) {
  if (!source.startsWith('data:')) {
    throw new Error('The image is not a supported base64 data URI.')
  }
  const separatorIndex = source.indexOf(';base64,', 5)
  if (separatorIndex < 0) {
    throw new Error('The image is not a supported base64 data URI.')
  }
  const contentType = source.slice(5, separatorIndex).toLowerCase()
  const encoded = source.slice(separatorIndex + ';base64,'.length)
  if (
    !SUPPORTED_IMAGE_CONTENT_TYPES.has(contentType)
    || encoded.length % 4 !== 0
    || !/^[A-Za-z0-9+/]*={0,2}$/.test(encoded)
  ) {
    throw new Error('The image is not a supported base64 data URI.')
  }

  const bytes = Buffer.from(encoded, 'base64')
  if (bytes.toString('base64') !== encoded) {
    throw new Error('The image data is not valid canonical base64.')
  }
  return { bytes, contentType }
}

function buildTemplateAssetPath(templateId: string, assetId: string) {
  return `/import/${templateId}/assets/${assetId}`
}

function parseTemplateAssetPath(source: string) {
  const match = TEMPLATE_ASSET_PATH_PATTERN.exec(source)
  const templateId = match?.[1]
  const assetId = match?.[2]
  return templateId && assetId ? { assetId, templateId } : undefined
}

function templateAssetDataUri(asset: TemplateAsset) {
  return `data:${asset.contentType};base64,${asset.bytes.toString('base64')}`
}

function templateAssetKey(templateId: string, assetId: string) {
  return `${templateId}\u0000${assetId}`
}
