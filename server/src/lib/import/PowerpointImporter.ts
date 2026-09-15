import path from 'node:path'
import JSZip from 'jszip'
import {
  BACKGROUND_ELEMENT_ID_BASE,
  BACKGROUND_Z_INDEX_OFFSET,
  EMU_PER_INCH,
  GEOMETRY_DECIMAL_PLACES,
  INHERITED_LAYOUT_Z_INDEX,
  INHERITED_MASTER_Z_INDEX,
  PX_PER_INCH,
  SLIDE_BACKGROUND_Z_INDEX,
  SLIDE_BOUNDS_TOLERANCE_PX,
} from '../shared/PowerpointConstants'
import { normalizePresentationSpec } from '../shared/PowerpointNormalizer'
import type {
  ExtractedRelationship,
  NormalizedPresentation,
  XmlNode,
} from '../shared/PowerpointTypes'
import { coerceNumber } from '../shared/PowerpointUtils'
import { compactPresentation } from './PowerpointCompactWriter'
import {
  composeTransform,
  extractElementTransform,
  groupTransform,
  identityTransform,
} from './PowerpointGeometry'
import type {
  ExtractedElementRecord,
  ImportedPowerPoint,
  ImportPowerPointBytesOptions,
  PlaceholderSourceIndex,
  SlideSize,
  TransformMatrix,
} from './PowerpointImportTypes'
import {
  firstDefined,
  round,
} from './PowerpointImportUtils'
import {
  collectSupportParts,
  extractSlidePaths,
  extractSlideSize,
  maybeReadZipText,
  parseRelationships,
  readZipText,
  relationshipPathFor,
} from './PowerpointOoxml'
import { extractNonVisual } from './PowerpointShapeUtils'
import { extractTableElements } from './PowerpointTableExtractor'
import {
  applyExtractedTypography,
  cloneNode,
  extractBackgroundColor,
  extractText,
} from './PowerpointText'
import {
  child,
  collectRelationshipIds,
  descendants,
  findDescendant,
  paragraphText,
  parseXml,
} from './PowerpointXml'

export type {
  ImportedPowerPoint,
  ImportPowerPointBytesOptions,
  PowerPointCanvasJson,
} from './PowerpointImportTypes'

/**
 * Converts PowerPoint OOXML directly into the native object consumed by
 * <SvgSlideCanvas input={template.jsonSpec} />. The compact output is also
 * accepted by the existing JSON-to-PowerPoint exporter.
 */

/** Import PowerPoint OOXML bytes and keep images embedded in the returned JSON. */
export async function importPowerPointBytes(
  source: Buffer,
  options: ImportPowerPointBytesOptions = {},
): Promise<ImportedPowerPoint> {
  validateSlideNumber(options.slide)
  if (!Buffer.isBuffer(source) || source.length === 0) {
    throw new Error('PowerPoint source bytes are required.')
  }
  const sourceName = normalizeSourceName(options.sourceName)
  return importPowerPointSource(source, {
    slide: options.slide,
    sourcePptx: sourceName,
  })
}

type ImportPowerPointSourceOptions = {
  slide?: number
  sourcePptx: string
}

async function importPowerPointSource(
  source: Buffer,
  options: ImportPowerPointSourceOptions,
): Promise<ImportedPowerPoint> {
  const zip = await JSZip.loadAsync(source)
  const presentationXml = await readZipText(zip, 'ppt/presentation.xml')
  const presentationRels = parseRelationships(
    await readZipText(zip, 'ppt/_rels/presentation.xml.rels'),
    'ppt',
  )
  const slideSize = extractSlideSize(presentationXml)
  const slidePaths = extractSlidePaths(presentationXml, presentationRels)

  if (!slidePaths.length) {
    throw new Error('No slides were found in the PowerPoint deck.')
  }

  const selectedSlides = options.slide
    ? slidePaths.slice(options.slide - 1, options.slide)
    : slidePaths

  if (options.slide && selectedSlides.length === 0) {
    throw new Error(`Slide ${options.slide} does not exist. The deck has ${slidePaths.length} slide(s).`)
  }
  const warnings: string[] = []
  const normalizedSlides: NormalizedPresentation['slides'] = []
  let deckTitle = path.basename(options.sourcePptx, path.extname(options.sourcePptx))

  for (const slidePath of selectedSlides) {
    const slideNumber = slidePaths.indexOf(slidePath) + 1
    const extracted = await extractSlide(zip, options.sourcePptx, slidePath, slideNumber, slideSize)
    const { presentation, issues } = normalizePresentationSpec(extracted)
    const errors = issues.filter((issue) => issue.level === 'error')

    if (!presentation || errors.length) {
      const formatted = issues
        .map((issue) => `${issue.level.toUpperCase()} ${issue.path}: ${issue.message}`)
        .join('\n')
      throw new Error(`Slide ${slideNumber} could not be normalized.\n${formatted}`)
    }

    warnings.push(
      ...issues
        .filter((issue) => issue.level === 'warning')
        .map((issue) => `slide ${slideNumber}: ${issue.path}: ${issue.message}`),
    )

    applyExtractedTypography(presentation, extracted)
    normalizedSlides.push(...presentation.slides)
    if (normalizedSlides.length === 1) {
      deckTitle = presentation.meta.title || deckTitle
    }
  }

  const compact = compactPresentation({
    meta: {
      title: deckTitle,
      width: slideSize.widthPx,
      height: slideSize.heightPx,
      preserveElementOrder: true,
      showBranding: false,
      sourceType: 'native-presentation',
    },
    slides: normalizedSlides,
  })
  const roundTrip = normalizePresentationSpec(compact)
  const roundTripErrors = roundTrip.issues.filter((issue) => issue.level === 'error')
  if (!roundTrip.presentation || roundTripErrors.length) {
    throw new Error(
      `Generated JSON failed the TemplateCanvas contract:\n${roundTripErrors
        .map((issue) => `${issue.path}: ${issue.message}`)
        .join('\n')}`,
    )
  }
  return {
    jsonSpec: compact,
    warnings,
    sourceSlideCount: slidePaths.length,
    importedSlideCount: normalizedSlides.length,
  }
}

function validateSlideNumber(slide: number | undefined) {
  if (slide !== undefined && (!Number.isInteger(slide) || slide < 1)) {
    throw new Error('slide must be a positive whole slide number.')
  }
}

function normalizeSourceName(value: string | undefined) {
  const sourceName = value ?? 'upload.pptx'
  if (!sourceName.trim()) throw new Error('sourceName must not be blank.')
  return path.basename(sourceName)
}

async function extractSlide(
  zip: JSZip,
  sourcePptx: string,
  slidePath: string,
  slideNumber: number,
  slideSize: SlideSize,
) {
  const slideXml = await readZipText(zip, slidePath)
  const relationshipsPath = relationshipPathFor(slidePath)
  const rawRelationshipsXml = await maybeReadZipText(zip, relationshipsPath)
  const relationships = rawRelationshipsXml
    ? parseRelationships(rawRelationshipsXml, path.posix.dirname(slidePath))
    : []
  const slideAst = parseXml(slideXml)
  const inherited = await extractInheritedSlideParts(zip, relationships, slideSize)
  const showMasterShapes = findDescendant(slideAst, 'p:sld')?.attributes?.showMasterSp !== '0'
  const slideBackgroundElement = extractBackgroundImageElement(
    slideAst,
    slidePath,
    'Slide',
    SLIDE_BACKGROUND_Z_INDEX,
    slideSize,
  )
  const shapeTree = findDescendant(slideAst, 'p:spTree')
  const slideElements = shapeTree
    ? extractShapeTreeElements(
        shapeTree,
        identityTransform(),
        slidePath,
        inherited.placeholderSources,
      )
    : []
  const baseElements = [
    ...(showMasterShapes ? inherited.masterElements : []),
    ...inherited.layoutElements,
    ...(slideBackgroundElement ? [slideBackgroundElement] : []),
  ]
  const elements = [
    ...baseElements,
    ...slideElements.map((element) => ({
      ...element,
      zIndex: coerceNumber(element.zIndex, 0) + baseElements.length,
    })),
  ]
  const allRelationships = [...relationships, ...inherited.relationships]

  return {
    sourcePptx,
    slideNumber,
    slidePath,
    relationshipsPath,
    relationships: allRelationships,
    rawRelationshipsXml,
    relationshipIds: Array.from(new Set(elements.flatMap((element) => element.relationshipIds ?? []))),
    slideSize: {
      cx: slideSize.cx,
      cy: slideSize.cy,
      widthInches: round(slideSize.cx / EMU_PER_INCH, GEOMETRY_DECIMAL_PLACES),
      heightInches: round(slideSize.cy / EMU_PER_INCH, GEOMETRY_DECIMAL_PLACES),
      widthPx: slideSize.widthPx,
      heightPx: slideSize.heightPx,
    },
    shapeTree: {
      elements,
    },
    supportParts: await collectSupportParts(zip, allRelationships),
    backgroundColor: extractBackgroundColor(slideAst) || inherited.backgroundColor,
    rawXml: slideXml,
    summary: {
      totalDrawableElements: elements.length,
      textElementCount: elements.filter((element) => !!paragraphText(element.text)).length,
    },
  }
}

async function extractInheritedSlideParts(
  zip: JSZip,
  slideRelationships: ExtractedRelationship[],
  slideSize: SlideSize,
) {
  const layoutRelationship = slideRelationships.find((relationship) =>
    relationship.Type?.includes('/slideLayout'),
  )
  const layout = layoutRelationship?.resolvedTarget
    ? await extractRelatedDrawablePart(
        zip,
        layoutRelationship.resolvedTarget,
        'layout',
        INHERITED_LAYOUT_Z_INDEX,
        slideSize,
      )
    : emptyDrawablePart()
  const masterRelationship = layout.rawRelationships.find((relationship) =>
    relationship.Type?.includes('/slideMaster'),
  )
  const master = masterRelationship?.resolvedTarget
    ? await extractRelatedDrawablePart(
        zip,
        masterRelationship.resolvedTarget,
        'master',
        INHERITED_MASTER_Z_INDEX,
        slideSize,
      )
    : emptyDrawablePart()

  return {
    masterElements: layout.showMasterShapes ? master.elements : [],
    layoutElements: layout.elements,
    relationships: [...layout.relationships, ...master.relationships],
    backgroundColor: layout.backgroundColor || master.backgroundColor,
    placeholderSources: {
      layout: layout.placeholderNodes,
      master: master.placeholderNodes,
      masterTextStyles: master.textStyles,
    },
  }
}

async function extractRelatedDrawablePart(
  zip: JSZip,
  partPath: string,
  label: string,
  zOffset: number,
  slideSize: SlideSize,
) {
  const rawXml = await maybeReadZipText(zip, partPath)
  if (!rawXml) {
    return emptyDrawablePart()
  }

  const relationshipsPath = relationshipPathFor(partPath)
  const rawRelationshipsXml = await maybeReadZipText(zip, relationshipsPath)
  const relationships = rawRelationshipsXml
    ? parseRelationships(rawRelationshipsXml, path.posix.dirname(partPath))
    : []
  const scopedRelationships = relationships.map((relationship) => ({
    ...relationship,
    Id: scopedRelationshipId(partPath, relationship.Id),
  }))
  const partAst = parseXml(rawXml)
  const backgroundElement = extractBackgroundImageElement(
    partAst,
    `${partPath}#${label}`,
    `${label[0]?.toUpperCase() ?? ''}${label.slice(1)}`,
    zOffset + BACKGROUND_Z_INDEX_OFFSET,
    slideSize,
    partPath,
  )
  const shapeTree = findDescendant(partAst, 'p:spTree')
  const placeholderNodes = shapeTree
    ? descendants(shapeTree, 'p:sp').filter((node) => !!extractNonVisual(node).placeholder)
    : []
  const drawableElements = shapeTree
    ? extractShapeTreeElements(shapeTree, identityTransform(), `${partPath}#${label}`).map((element) => ({
        ...element,
        relationshipIds: element.relationshipIds?.map((relationshipId) =>
          scopedRelationshipId(partPath, relationshipId),
        ),
        zIndex: zOffset + coerceNumber(element.zIndex, 0),
      }))
    : []
  const elements = [
    ...(backgroundElement ? [backgroundElement] : []),
    ...drawableElements,
  ].filter((element) => shouldKeepInheritedElement(element, scopedRelationships, slideSize))

  return {
    elements,
    relationships: scopedRelationships,
    rawRelationships: relationships,
    backgroundColor: extractBackgroundColor(partAst),
    placeholderNodes,
    textStyles: {
      title: findDescendant(partAst, 'p:titleStyle'),
      body: findDescendant(partAst, 'p:bodyStyle'),
      other: findDescendant(partAst, 'p:otherStyle'),
    },
    showMasterShapes: findDescendant(partAst, 'p:sldLayout')?.attributes?.showMasterSp !== '0',
  }
}

function extractBackgroundImageElement(
  root: XmlNode,
  pathLabel: string,
  label: string,
  zIndex: number,
  slideSize: SlideSize,
  relationshipScope?: string,
): ExtractedElementRecord | undefined {
  const background = findDescendant(root, 'p:bg')
  const blipFill = findDescendant(background, 'a:blipFill')
  const blip = findDescendant(blipFill, 'a:blip')
  const relationshipId = blip?.attributes?.['r:embed'] || blip?.attributes?.['r:link']

  if (!relationshipId) {
    return undefined
  }

  const scopedId = relationshipScope ? scopedRelationshipId(relationshipScope, relationshipId) : relationshipId
  const id = BACKGROUND_ELEMENT_ID_BASE + Math.abs(Math.trunc(zIndex))

  return {
    path: `${pathLabel}#background-image`,
    zIndex,
    tag: 'p:bg',
    kind: 'graphicFrame',
    nonVisual: {
      id,
      name: `${label} Background Image`,
    },
    transform: {
      xPx: 0,
      yPx: 0,
      widthPx: slideSize.widthPx,
      heightPx: slideSize.heightPx,
      rotation: 0,
      xInches: 0,
      yInches: 0,
      widthInches: round(slideSize.widthPx / PX_PER_INCH, GEOMETRY_DECIMAL_PLACES),
      heightInches: round(slideSize.heightPx / PX_PER_INCH, GEOMETRY_DECIMAL_PLACES),
    },
    relationshipIds: [scopedId],
    xmlAst: background,
  }
}

function emptyDrawablePart() {
  return {
    elements: [] as ExtractedElementRecord[],
    relationships: [] as ExtractedRelationship[],
    rawRelationships: [] as ExtractedRelationship[],
    backgroundColor: undefined as string | undefined,
    placeholderNodes: [] as XmlNode[],
    textStyles: {} as PlaceholderSourceIndex['masterTextStyles'],
    showMasterShapes: true,
  }
}

function scopedRelationshipId(partPath: string, relationshipId: string) {
  return `${partPath}:${relationshipId}`
}

function shouldKeepInheritedElement(
  element: ExtractedElementRecord,
  relationships: ExtractedRelationship[],
  slideSize?: SlideSize,
) {
  const nonVisual = element.nonVisual

  if (nonVisual?.placeholder) {
    return false
  }

  const text = paragraphText(element.text).replace(/\s+/g, ' ').trim().toLowerCase()
  if (
    text.includes('click to edit master') ||
    text.includes('edit master text styles') ||
    text.includes('master title style')
  ) {
    return false
  }

  if (element.kind === 'graphicFrame') {
    return (element.relationshipIds ?? []).every((relationshipId) => {
      const relationship = relationships.find((entry) => entry.Id === relationshipId)
      const extension = (relationship?.resolvedTarget || relationship?.Target || '').split('.').pop()?.toLowerCase()
      return extension !== 'emf' && extension !== 'wmf'
    })
  }

  if (slideSize && isCompletelyOutsideSlide(element, slideSize)) {
    return false
  }

  return true
}

function isCompletelyOutsideSlide(element: ExtractedElementRecord, slideSize: SlideSize) {
  const transform = element.transform
  const x = coerceNumber(transform?.xPx, 0)
  const y = coerceNumber(transform?.yPx, 0)
  const width = coerceNumber(transform?.widthPx, 0)
  const height = coerceNumber(transform?.heightPx, 0)
  return (
    x + width < -SLIDE_BOUNDS_TOLERANCE_PX ||
    y + height < -SLIDE_BOUNDS_TOLERANCE_PX ||
    x > slideSize.widthPx + SLIDE_BOUNDS_TOLERANCE_PX ||
    y > slideSize.heightPx + SLIDE_BOUNDS_TOLERANCE_PX
  )
}

function extractShapeTreeElements(
  parent: XmlNode,
  matrix: TransformMatrix,
  pathLabel: string,
  placeholderIndex?: PlaceholderSourceIndex,
) {
  const elements: ExtractedElementRecord[] = []

  for (const [nodeIndex, node] of (parent.children ?? []).entries()) {
    if (node.tag === 'p:grpSp') {
      const groupMatrix = composeTransform(matrix, groupTransform(node))
      const groupPath = `${pathLabel}#p:grpSp[${nodeIndex + 1}]`
      elements.push(...extractShapeTreeElements(node, groupMatrix, groupPath, placeholderIndex))
      continue
    }

    if (!['p:sp', 'p:cxnSp', 'p:pic', 'p:graphicFrame'].includes(node.tag)) {
      continue
    }

    if (node.tag === 'p:graphicFrame' && findDescendant(node, 'a:tbl')) {
      elements.push(...extractTableElements(node, matrix, pathLabel, elements.length))
      continue
    }

    const nonVisual = extractNonVisual(node)
    const placeholderSources = findPlaceholderSources(nonVisual.placeholder, placeholderIndex)
    const transform = extractElementTransform(node, matrix, placeholderSources)
    const shapeProperties = mergeShapeProperties(
      child(node, 'p:spPr'),
      placeholderSources.map((source) => child(source, 'p:spPr')),
    )
    const style = child(node, 'p:style') ?? firstDefined(
      placeholderSources.map((source) => child(source, 'p:style')),
    )
    const text = extractText(
      child(node, 'p:txBody'),
      [
        findPlaceholderTextStyle(nonVisual.placeholder, placeholderIndex),
        ...placeholderSources
          .slice()
          .reverse()
          .map((source) => child(source, 'p:txBody')),
      ],
    )
    const preset = child(shapeProperties, 'a:prstGeom')?.attributes?.prst
    const relationshipIds = Array.from(new Set(collectRelationshipIds(node)))

    elements.push({
      path: `${pathLabel}#${node.tag}[${elements.length + 1}]`,
      zIndex: elements.length,
      tag: node.tag,
      kind:
        node.tag === 'p:cxnSp'
          ? 'connector'
          : node.tag === 'p:graphicFrame' || node.tag === 'p:pic'
            ? 'graphicFrame'
            : 'shape',
      nonVisual,
      transform,
      presetGeometry: preset
        ? {
            preset,
            xmlAst: child(shapeProperties, 'a:prstGeom'),
          }
        : undefined,
      relationshipIds,
      xmlAst: node,
      text,
      shapeProperties,
      style,
    })
  }

  return elements
}

function findPlaceholderTextStyle(
  placeholder: { type?: string; idx?: string } | undefined,
  index: PlaceholderSourceIndex | undefined,
) {
  if (!placeholder || !index) {
    return undefined
  }

  if (placeholder.type === 'title' || placeholder.type === 'ctrTitle') {
    return index.masterTextStyles.title
  }

  if (
    placeholder.type === 'body'
    || placeholder.type === 'subTitle'
    || placeholder.type === 'obj'
  ) {
    return index.masterTextStyles.body
  }

  return index.masterTextStyles.other
}

function findPlaceholderSources(
  placeholder: { type?: string; idx?: string } | undefined,
  index: PlaceholderSourceIndex | undefined,
) {
  if (!placeholder || !index) {
    return []
  }

  return [
    findMatchingPlaceholder(index.layout, placeholder),
    findMatchingPlaceholder(index.master, placeholder),
  ].filter((node): node is XmlNode => !!node)
}

function findMatchingPlaceholder(
  candidates: XmlNode[],
  placeholder: { type?: string; idx?: string },
) {
  return candidates.find((candidate) => {
    const candidatePlaceholder = extractNonVisual(candidate).placeholder
    if (!candidatePlaceholder) {
      return false
    }

    if (placeholder.idx && candidatePlaceholder.idx) {
      return placeholder.idx === candidatePlaceholder.idx
    }

    const expectedType = placeholder.type || 'body'
    const candidateType = candidatePlaceholder.type || 'body'
    return expectedType === candidateType
  })
}

function mergeShapeProperties(
  primary: XmlNode | undefined,
  fallbacks: Array<XmlNode | undefined>,
) {
  if (!primary && !fallbacks.some(Boolean)) {
    return undefined
  }

  const merged = cloneNode(primary) ?? { tag: 'p:spPr', children: [] }
  const existingTags = new Set((merged.children ?? []).map((node) => node.tag))

  for (const fallback of fallbacks) {
    for (const fallbackChild of fallback?.children ?? []) {
      if (!existingTags.has(fallbackChild.tag)) {
        const cloned = cloneNode(fallbackChild)
        if (cloned) {
          merged.children?.push(cloned)
          existingTags.add(fallbackChild.tag)
        }
      }
    }
  }

  return merged
}
