// @vitest-environment node

import JSZip from 'jszip'
import PptxGenJS from 'pptxgenjs'
import { describe, expect, it } from 'vitest'

import type {
  PowerPointCanvasElement,
  PowerPointCanvasLineElement,
  PowerPointCanvasShapeElement,
  PowerPointCanvasTextElement,
} from '../src/lib/import/PowerpointImportTypes'
import { extractElementTransform, identityTransform } from '../src/lib/import/PowerpointGeometry'
import type { ExtractedSlideRecord } from '../src/lib/import/PowerpointImportTypes'
import { collectSupportParts } from '../src/lib/import/PowerpointOoxml'
import { extractTableElements } from '../src/lib/import/PowerpointTableExtractor'
import { applyExtractedTypography, extractText } from '../src/lib/import/PowerpointText'
import { findDescendant, parseXml } from '../src/lib/import/PowerpointXml'
import { normalizeExtractedPresentation } from '../src/lib/shared/PowerpointExtractedNormalizer'
import type { NormalizedElement, NormalizedShapeElement } from '../src/lib/shared/PowerpointTypes'
import { parseColor } from '../src/lib/shared/PowerpointUtils'
import { ExportPowerPointService } from '../src/services/ExportPowerPointService'
import { LibraryPowerPointConverter } from '../src/services/PowerPointConverter'
import { SqliteTemplateRepository } from '../src/repositories/SqliteTemplateRepository'
import { normalizePresentation } from '../src/services/NormalizePresentation'

const THEME_XML = `
  <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
    <a:themeElements>
      <a:clrScheme name="Custom">
        <a:dk1><a:srgbClr val="070154"/></a:dk1>
        <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
        <a:accent3><a:srgbClr val="CED7E6"/></a:accent3>
      </a:clrScheme>
      <a:fontScheme name="Custom">
        <a:majorFont><a:latin typeface="Arial"/></a:majorFont>
        <a:minorFont><a:latin typeface="Arial"/></a:minorFont>
      </a:fontScheme>
    </a:themeElements>
  </a:theme>
`

const TABLE_XML = `
  <p:graphicFrame xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
    <p:xfrm><a:off x="0" y="0"/><a:ext cx="3810000" cy="1905000"/></p:xfrm>
    <a:graphic><a:graphicData><a:tbl>
      <a:tblGrid>
        <a:gridCol w="952500"/><a:gridCol w="952500"/>
        <a:gridCol w="952500"/><a:gridCol w="952500"/>
      </a:tblGrid>
      <a:tr h="952500">
        <a:tc rowSpan="2">
          <a:txBody><a:p><a:pPr algn="ctr"/><a:r><a:rPr sz="1200" b="1"><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:latin typeface="+mn-lt"/></a:rPr><a:t>Category</a:t></a:r></a:p></a:txBody>
          <a:tcPr anchor="ctr"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:tcPr>
        </a:tc>
        <a:tc gridSpan="2">
          <a:txBody><a:p><a:pPr algn="ctr"/><a:r><a:rPr sz="1200"><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></a:rPr><a:t>Target Position</a:t></a:r></a:p></a:txBody>
          <a:tcPr><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:tcPr>
        </a:tc>
        <a:tc hMerge="1"><a:txBody><a:p/></a:txBody><a:tcPr/></a:tc>
        <a:tc rowSpan="2">
          <a:txBody><a:p><a:r><a:rPr sz="1200"><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></a:rPr><a:t>Commentary</a:t></a:r></a:p></a:txBody>
          <a:tcPr><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:tcPr>
        </a:tc>
      </a:tr>
      <a:tr h="0">
        <a:tc vMerge="1"><a:txBody><a:p/></a:txBody><a:tcPr/></a:tc>
        <a:tc>
          <a:txBody><a:p><a:r><a:rPr sz="1100"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:rPr><a:t>Value</a:t></a:r></a:p></a:txBody>
          <a:tcPr><a:solidFill><a:schemeClr val="accent3"/></a:solidFill></a:tcPr>
        </a:tc>
        <a:tc>
          <a:txBody><a:p><a:r><a:rPr sz="1100"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:rPr><a:t>Risk</a:t></a:r></a:p></a:txBody>
          <a:tcPr><a:solidFill><a:schemeClr val="accent3"/></a:solidFill></a:tcPr>
        </a:tc>
        <a:tc vMerge="1"><a:txBody><a:p/></a:txBody><a:tcPr/></a:tc>
      </a:tr>
    </a:tbl></a:graphicData></a:graphic>
  </p:graphicFrame>
`

describe('Rust PowerPoint parity', () => {
  it('inherits master title typography without replacing explicit run properties', () => {
    const slideTextBody = findDescendant(parseXml(`
      <p:txBody>
        <a:bodyPr/>
        <a:lstStyle/>
        <a:p>
          <a:r><a:rPr sz="2000"/><a:t>Title</a:t></a:r>
        </a:p>
      </p:txBody>
    `), 'p:txBody')
    const masterTitleStyle = findDescendant(parseXml(`
      <p:titleStyle>
        <a:lvl1pPr algn="l">
          <a:defRPr sz="2200" b="1">
            <a:solidFill><a:schemeClr val="tx1"/></a:solidFill>
            <a:latin typeface="+mj-lt"/>
          </a:defRPr>
        </a:lvl1pPr>
      </p:titleStyle>
    `), 'p:titleStyle')

    expect(extractText(slideTextBody, [masterTitleStyle])).toMatchObject({
      paragraphs: [
        {
          properties: { algn: 'l' },
          runs: [
            {
              text: 'Title',
              properties: {
                sz: '2000',
                b: '1',
                fontSchemeColor: 'tx1',
                fontFace: '+mj-lt',
              },
            },
          ],
        },
      ],
    })
  })

  it('round trips styled text boxes without losing fill, geometry, or run typography', async () => {
    const source = new PptxGenJS()
    source.layout = 'LAYOUT_WIDE'
    const slide = source.addSlide()
    slide.addText(
      [
        {
          text: 'Architecture',
          options: {
            bold: true,
            color: 'FFFFFF',
            fontFace: 'Arial',
            fontSize: 19.5,
          },
        },
        {
          text: ' overview',
          options: {
            color: 'FFFFFF',
            fontFace: 'Aptos',
            fontSize: 11,
            italic: true,
          },
        },
      ],
      {
        x: 1,
        y: 1.25,
        w: 3.5,
        h: 0.625,
        fill: { color: '070154', transparency: 20 },
        line: { color: '0047FF', transparency: 30, width: 2 },
        margin: 0,
      },
    )
    slide.addText('Database', {
      x: 5,
      y: 1.25,
      w: 1.5,
      h: 1,
      shape: source.ShapeType.flowChartMagneticDisk,
      fill: { color: 'E8EEF8' },
      line: { color: '070154', width: 1 },
      color: '070154',
      fontFace: 'Arial',
      fontSize: 12,
      margin: 0,
    })
    const sourceBytes = await source.write({ outputType: 'nodebuffer' })
    if (!Buffer.isBuffer(sourceBytes)) {
      throw new Error('Expected PptxGenJS to return a Node.js buffer.')
    }
    const sourceArchive = await JSZip.loadAsync(sourceBytes)
    const sourceSlideXml = await sourceArchive.file('ppt/slides/slide1.xml')?.async('text')
    if (!sourceSlideXml) {
      throw new Error('Expected the source deck to contain slide XML.')
    }
    sourceArchive.file(
      'ppt/slides/slide1.xml',
      sourceSlideXml.replace('<p:cNvSpPr/>', '<p:cNvSpPr txBox="1"/>'),
    )
    const textBoxSourceBytes = await sourceArchive.generateAsync({ type: 'nodebuffer' })

    const converter = new LibraryPowerPointConverter()
    const imported = await converter.convert(textBoxSourceBytes)
    const importedElement = textElementWithText(
      imported.templateJson.presentation.slides[0]?.elements,
      'Architecture overview',
    )
    expect(imported.templateJson.presentation.slides[0]).toMatchObject({
      width: 1280,
      height: 720,
    })
    expect(importedElement).toMatchObject({
      type: 'text',
      x: 96,
      y: 120,
      w: 336,
      h: 60,
      fill: '070154',
      fillOpacity: 0.8,
      stroke: '0047FF',
      strokeOpacity: 0.7,
      strokeWidth: 2,
      fontFace: 'Arial',
      fontSize: 19.5,
      runs: [
        expect.objectContaining({
          text: 'Architecture',
          bold: true,
          color: 'FFFFFF',
          fontFace: 'Arial',
          fontSize: 19.5,
        }),
        expect.objectContaining({
          text: ' overview',
          color: 'FFFFFF',
          fontFace: 'Aptos',
          fontSize: 11,
          italic: true,
        }),
      ],
    })
    expect(shapeElementWithText(
      imported.templateJson.presentation.slides[0]?.elements,
      'Database',
    )).toMatchObject({
      type: 'shape',
      shape: 'flowChartMagneticDisk',
      x: 480,
      y: 120,
      w: 144,
      h: 96,
      fill: 'E8EEF8',
      stroke: '070154',
      strokeWidth: 1,
      fontFace: 'Arial',
      fontSize: 12,
    })

    const templates = new SqliteTemplateRepository()
    try {
      const exported = await new ExportPowerPointService(templates).export(imported.templateJson)
      const roundTrip = await converter.convert(Buffer.from(exported.bytes))
      expect(textElementWithText(
        roundTrip.templateJson.presentation.slides[0]?.elements,
        'Architecture overview',
      )).toMatchObject({
        x: 96,
        y: 120,
        w: 336,
        h: 60,
        fill: '070154',
        fillOpacity: 0.8,
        stroke: '0047FF',
        strokeOpacity: 0.7,
        strokeWidth: 2,
        fontFace: 'Arial',
        fontSize: 19.5,
        runs: [
          expect.objectContaining({ fontFace: 'Arial', fontSize: 19.5 }),
          expect.objectContaining({ fontFace: 'Aptos', fontSize: 11 }),
        ],
      })
      expect(shapeElementWithText(
        roundTrip.templateJson.presentation.slides[0]?.elements,
        'Database',
      )).toMatchObject({
        type: 'shape',
        shape: 'flowChartMagneticDisk',
        x: 480,
        y: 120,
        w: 144,
        h: 96,
        fill: 'E8EEF8',
        stroke: '070154',
        strokeWidth: 1,
        fontFace: 'Arial',
        fontSize: 12,
      })
    } finally {
      templates.close()
    }
  })

  it('imports and exports straight and elbow connectors without flattening their routing mode', async () => {
    const source = new PptxGenJS()
    source.layout = 'LAYOUT_WIDE'
    const slide = source.addSlide()
    slide.addShape(source.ShapeType.line, {
      x: 1,
      y: 1,
      w: 2,
      h: 0,
      line: { color: '070154', endArrowType: 'triangle', width: 2 },
    })
    slide.addShape('bentConnector2' as PptxGenJS.ShapeType, {
      x: 4,
      y: 1,
      w: 2.5,
      h: 1,
      line: { color: '0047FF', endArrowType: 'triangle', width: 2 },
    })
    const sourceBytes = await source.write({ outputType: 'nodebuffer' })
    if (!Buffer.isBuffer(sourceBytes)) {
      throw new Error('Expected PptxGenJS to return a Node.js buffer.')
    }

    const converter = new LibraryPowerPointConverter()
    const imported = await converter.convert(sourceBytes)
    expect(lineElements(imported.templateJson.presentation.slides[0]?.elements)).toEqual([
      expect.objectContaining({ lineType: 'straight', endArrow: 'triangle' }),
      expect.objectContaining({
        lineType: 'elbow',
        elbowDirection: 'horizontal-first',
        endArrow: 'triangle',
      }),
    ])

    const templates = new SqliteTemplateRepository()
    try {
      const exported = await new ExportPowerPointService(templates).export(imported.templateJson)
      const archive = await JSZip.loadAsync(exported.bytes)
      const slideXml = await archive.file('ppt/slides/slide1.xml')?.async('text')
      expect(slideXml).toContain('prst="bentConnector2"')

      const roundTrip = await converter.convert(Buffer.from(exported.bytes))
      expect(lineElements(roundTrip.templateJson.presentation.slides[0]?.elements)).toEqual([
        expect.objectContaining({ lineType: 'straight' }),
        expect.objectContaining({
          lineType: 'elbow',
          elbowDirection: 'horizontal-first',
          endArrow: 'triangle',
        }),
      ])
    } finally {
      templates.close()
    }
  })

  it('exports vertical-first elbows without moving either visible arrowhead', async () => {
    const normalized = normalizePresentation({
      presentation: {
        showBranding: false,
        slides: [{
          elements: [{
            type: 'line',
            id: 'vertical-first',
            lineType: 'elbow',
            elbowDirection: 'vertical-first',
            x1: 100,
            y1: 100,
            x2: 300,
            y2: 180,
            beginArrow: 'diamond',
            endArrow: 'triangle',
          }],
        }],
      },
    })
    const templates = new SqliteTemplateRepository()

    try {
      const exported = await new ExportPowerPointService(templates).export(normalized.templateJson)
      const converter = new LibraryPowerPointConverter()
      const roundTrip = await converter.convert(Buffer.from(exported.bytes))

      expect(lineElements(roundTrip.templateJson.presentation.slides[0]?.elements)[0]).toMatchObject({
        lineType: 'elbow',
        elbowDirection: 'horizontal-first',
        x1: 300,
        y1: 180,
        x2: 100,
        y2: 100,
        beginArrow: 'triangle',
        endArrow: 'diamond',
      })
    } finally {
      templates.close()
    }
  })

  it('infers legacy connector routing from arbitrary endpoint coordinates', () => {
    const normalized = normalizePresentation({
      presentation: {
        showBranding: false,
        slides: [{
          elements: [
            { type: 'line', id: 'horizontal', x1: 10, y1: 20, x2: 200, y2: 20 },
            { type: 'line', id: 'vertical', x1: 10, y1: 20, x2: 10, y2: 200 },
            { type: 'line', id: 'diagonal', x1: 10, y1: 20, x2: 200, y2: 80 },
            {
              type: 'line',
              id: 'explicit-straight',
              lineType: 'straight',
              x1: 10,
              y1: 20,
              x2: 200,
              y2: 80,
            },
          ],
        }],
      },
    })

    expect(lineElements(normalized.templateJson.presentation.slides[0]?.elements)
      .map((line) => line.lineType)).toEqual([
      'straight',
      'straight',
      'elbow',
      'straight',
    ])
  })

  it('preserves theme colors, fonts, and merged table cell assignments', () => {
    const table = findDescendant(parseXml(TABLE_XML), 'p:graphicFrame')
    if (!table) {
      throw new Error('Expected the table fixture to contain a graphic frame.')
    }
    const elements = extractTableElements(table, identityTransform(), 'slide', 6)
    expect(extractElementTransform(table, identityTransform())).toMatchObject({
      widthPx: 400,
      heightPx: 200,
    })

    const extracted: ExtractedSlideRecord = {
      sourcePptx: 'fixture.pptx',
      slideNumber: 1,
      slidePath: 'ppt/slides/slide1.xml',
      relationshipsPath: 'ppt/slides/_rels/slide1.xml.rels',
      relationships: [],
      relationshipIds: [],
      rawXml: TABLE_XML,
      shapeTree: { elements },
      slideSize: {
        cx: 12_192_000,
        cy: 6_858_000,
        widthInches: 13.3333,
        heightInches: 7.5,
        widthPx: 1280,
        heightPx: 720,
      },
      supportParts: {
        'ppt/theme/custom-theme.xml': {
          path: 'ppt/theme/custom-theme.xml',
          size: Buffer.byteLength(THEME_XML),
          relationshipType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
          contentTypeHint: 'application/xml',
          rawXml: THEME_XML,
        },
      },
      summary: { totalDrawableElements: elements.length, textElementCount: elements.length },
    }
    const presentation = normalizeExtractedPresentation(extracted, [])
    applyExtractedTypography(presentation, extracted)
    expect(presentation.slides[0]?.elements).toHaveLength(5)

    const category = shapeWithText(presentation.slides[0]?.elements, 'Category')
    expect(category).toMatchObject({
      x: 0,
      w: 100,
      h: 200,
      fill: '070154',
      textColor: 'FFFFFF',
      fontFace: 'Arial',
      bold: true,
      align: 'center',
    })
    expect(shapeWithText(presentation.slides[0]?.elements, 'Target Position')).toMatchObject({
      x: 100,
      w: 200,
      h: 100,
      fill: '070154',
    })
    expect(shapeWithText(presentation.slides[0]?.elements, 'Commentary')).toMatchObject({
      x: 300,
      w: 100,
      h: 200,
    })
    expect(shapeWithText(presentation.slides[0]?.elements, 'Value')).toMatchObject({
      x: 100,
      y: 100,
      fill: 'CED7E6',
    })
    expect(shapeWithText(presentation.slides[0]?.elements, 'Risk')).toMatchObject({
      x: 200,
      y: 100,
    })
  })

  it('applies color transforms after resolving the theme color', () => {
    const fill = findDescendant(
      parseXml('<a:solidFill><a:schemeClr val="dk1"><a:tint val="50000"/></a:schemeClr></a:solidFill>'),
      'a:solidFill',
    )
    expect(parseColor(fill, { dk1: '070154' }, '111827')).toBe('8380AA')
  })

  it('falls back to an arbitrarily named theme part when no relationship resolves one', async () => {
    const zip = new JSZip()
    zip.file('ppt/theme/custom-theme.xml', THEME_XML)

    const supportParts = await collectSupportParts(zip, [])

    expect(supportParts['ppt/theme/custom-theme.xml']).toMatchObject({
      relationshipType:
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
      rawXml: THEME_XML,
    })
  })

  it('matches the Rust JSON aliases, defaults, coercions, and warnings', () => {
    const normalized = normalizePresentation({
      title: 'Normalization parity',
      preserveElementOrder: 'yes',
      showBranding: 1.5,
      canvas: { width: null, w: 640 },
      slides: [
        null,
        {
          backgroundColor: '##abcdef',
          elements: [
            {
              type: 'text',
              x: '50%',
              y: '1in',
              w: '2in',
              h: 80,
              text: 'Parity',
              bold: true,
              fontFace: 'Arial',
              fontSize: ' ',
              italic: true,
            },
            { type: 'unsupported' },
          ],
        },
      ],
    })

    expect(normalized.templateJson.presentation).toMatchObject({
      title: 'Normalization parity',
      preserveElementOrder: true,
      showBranding: false,
      slides: [
        {
          id: 'slide-2',
          width: 1280,
          height: 720,
          backgroundColor: 'ABCDEF',
          elements: [
            expect.objectContaining({
              type: 'text',
              x: 640,
              y: 96,
              w: 192,
              fontSize: 18,
              fontFace: 'Arial',
              bold: true,
              italic: true,
              runs: [expect.objectContaining({
                bold: true,
                fontFace: 'Arial',
                fontSize: 18,
                italic: true,
              })],
            }),
          ],
        },
      ],
    })
    expect(normalized.warnings).toEqual([
      'slides[0]: Skipped a non-object slide entry.',
      'slides[1].elements[1]: Unsupported element type "unsupported".',
    ])
  })
})

function shapeWithText(
  elements: NormalizedElement[] | undefined,
  text: string,
): NormalizedShapeElement {
  const element = elements?.find(
    (candidate): candidate is NormalizedShapeElement => candidate.kind === 'shape' && candidate.label === text,
  )
  if (!element) {
    throw new Error(`Missing table cell ${text}.`)
  }
  return element
}

function textElementWithText(
  elements: PowerPointCanvasElement[] | undefined,
  text: string,
): PowerPointCanvasTextElement {
  const element = elements?.find(
    (candidate): candidate is PowerPointCanvasTextElement =>
      candidate.type === 'text' && candidate.text === text,
  )
  if (!element) {
    throw new Error(`Missing text element ${text}.`)
  }
  return element
}

function shapeElementWithText(
  elements: PowerPointCanvasElement[] | undefined,
  text: string,
): PowerPointCanvasShapeElement {
  const element = elements?.find(
    (candidate): candidate is PowerPointCanvasShapeElement =>
      candidate.type === 'shape' && candidate.text === text,
  )
  if (!element) {
    throw new Error(`Missing shape element ${text}.`)
  }
  return element
}

function lineElements(
  elements: PowerPointCanvasElement[] | undefined,
): PowerPointCanvasLineElement[] {
  return elements?.filter(
    (element): element is PowerPointCanvasLineElement => element.type === 'line',
  ) ?? []
}
