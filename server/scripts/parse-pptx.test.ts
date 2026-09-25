import JSZip from 'jszip'
import PptxGenJS from 'pptxgenjs'
import { beforeAll, describe, expect, it } from 'vitest'

import { buildThemedPptxBytes } from '../src/lib/export/PowerpointGenerator'
import { importPowerPointBytes } from '../src/lib/import/PowerpointImporter'
import type {
  PowerPointCanvasElement,
  PowerPointCanvasShapeElement,
  PowerPointCanvasTextElement,
} from '../src/lib/import/PowerpointImportTypes'
import { normalizePresentationSpec } from '../src/lib/shared/PowerpointNormalizer'

let source: Buffer<ArrayBufferLike> = Buffer.alloc(0)

beforeAll(async () => {
  const pptx = new PptxGenJS()
  pptx.layout = 'LAYOUT_WIDE'
  const first = pptx.addSlide()
  first.addText('First slide', { x: 0.5, y: 0.5, w: 3, h: 0.5 })

  const second = pptx.addSlide()
  second.background = { color: 'F4F7FB' }
  second.addText('Imported title', {
    x: 0.6,
    y: 0.4,
    w: 4,
    h: 0.6,
    fontFace: 'Aptos Display',
    fontSize: 28,
    bold: true,
    color: '070154',
    margin: 0,
  })
  second.addShape(pptx.ShapeType.ellipse, {
    x: 1,
    y: 1.5,
    w: 2.4,
    h: 1.4,
    fill: { color: '1DD566', transparency: 25 },
    line: { color: '0047FF', width: 2 },
  })
  second.addImage({
    data: 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAYAAABytg0kAAAAGElEQVR42mNkYGD4z8DAwMgAI0AMYBQAAD0AAgFIr6QAAAAASUVORK5CYII=',
    x: 4,
    y: 1.5,
    w: 2,
    h: 1.5,
    sizing: { type: 'crop', x: 0.25, y: 0.1, w: 1.5, h: 1.1 },
  })
  second.addShape(pptx.ShapeType.line, {
    x: 6.5,
    y: 1.75,
    w: 1.5,
    h: 0,
    line: { color: '070154', endArrowType: 'triangle', width: 1.5 },
  })
  second.addShape('bentConnector2' as PptxGenJS.ShapeType, {
    x: 6.5,
    y: 2.25,
    w: 1.5,
    h: 0.75,
    line: { color: '0047FF', endArrowType: 'triangle', width: 1.5 },
  })

  const third = pptx.addSlide()
  const headerCell = (text: string) => ({
    text,
    options: {
      bold: true,
      color: 'FFFFFF',
      fill: { color: '070154' },
      fontSize: 11,
    },
  })
  const bodyCell = (text: string) => ({ text, options: { color: '070154', fontSize: 10.5 } })
  third.addTable(
    [
      [headerCell('Approach'), headerCell('Details')],
      [bodyCell('Hosting Provider'), bodyCell('AWS')],
      [bodyCell('Hosting Methodology'), bodyCell('Mostly PaaS (EKS, RDS)\nSome IaaS (EC2 - Linux)')],
      [
        bodyCell('Availability'),
        bodyCell(
          'Multiple availability zones\nSingle production region (US-East-1)\nSingle development region (US-West-1)\nApp: no demand-based auto-scaling',
        ),
      ],
      [bodyCell('Disaster Recovery'), bodyCell('Data replication to EU-West-1, annual tabletop exercises')],
    ],
    {
      x: 8.4,
      y: 3.4,
      w: 4.4,
      h: 2.4,
      colW: [1.65, 2.75],
      border: { type: 'solid', color: 'CED7E6', pt: 0.5 },
      margin: 0.08,
      valign: 'middle',
    },
  )
  const generated = await pptx.write({ outputType: 'nodebuffer', compression: true })

  // PowerPoint commonly stores content-driven rows as h="0". Reproduce that
  // OOXML rather than relying on PptxGenJS's equal-height row defaults.
  const zip = await JSZip.loadAsync(generated)
  const tableSlidePath = 'ppt/slides/slide3.xml'
  const tableSlide = await zip.file(tableSlidePath)?.async('text')
  if (!tableSlide) {
    throw new Error('Missing generated table slide')
  }
  let rowIndex = 0
  const autoSizedRows = tableSlide.replace(/<a:tr h="[^"]+">/gu, (match) => {
    rowIndex += 1
    return rowIndex === 1 ? match : '<a:tr h="0">'
  })
  zip.file(tableSlidePath, autoSizedRows)
  source = await zip.generateAsync({ type: 'nodebuffer' })
})

describe('in-memory PowerPoint XML to TemplateCanvas JSON import', () => {
  it('accepts PowerPoint bytes and keeps image bytes embedded', async () => {
    const result = await importPowerPointBytes(source, {
      sourceName: 'source.pptx',
      slide: 2,
    })

    expect(result).toMatchObject({ importedSlideCount: 1, sourceSlideCount: 3 })
    expect(result.jsonSpec.presentation.slides[0]?.elements).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ src: expect.stringMatching(/^data:image\/png;base64,/u) }),
      ]),
    )
  })

  it('emits a selected slide in the native template.jsonSpec contract', async () => {
    const result = await importPowerPointBytes(source, {
      sourceName: 'source.pptx',
      slide: 2,
    })
    const json = result.jsonSpec as {
      presentation: {
        preserveElementOrder: boolean
        showBranding: boolean
        slides: Array<{
          backgroundColor: string
          elements: PowerPointCanvasElement[]
          id: string
        }>
      }
    }

    expect(result).toMatchObject({ sourceSlideCount: 3, importedSlideCount: 1 })

    expect(json.presentation).toMatchObject({
      preserveElementOrder: true,
      showBranding: false,
    })
    expect(json.presentation.slides).toHaveLength(1)
    expect(json.presentation.slides[0]).toMatchObject({
      id: 'slide-2',
      backgroundColor: 'F4F7FB',
    })
    expect(json.presentation.slides[0].elements).toEqual([
      expect.objectContaining({
        type: 'text',
        text: 'Imported title',
        fontFace: 'Aptos Display',
        fontSize: 28,
        textColor: '070154',
        bold: true,
      }),
      expect.objectContaining({
        type: 'shape',
        shape: 'ellipse',
        fill: '1DD566',
        fillOpacity: 0.75,
        stroke: '0047FF',
      }),
      expect.objectContaining({
        type: 'image',
        crop: {
          left: 0.13,
          right: 0.13,
          top: 0.07,
          bottom: 0.2,
        },
      }),
      expect.objectContaining({
        type: 'line',
        lineType: 'straight',
        endArrow: 'triangle',
      }),
      expect.objectContaining({
        type: 'line',
        lineType: 'elbow',
        elbowDirection: 'horizontal-first',
        endArrow: 'triangle',
      }),
    ])

    const normalized = normalizePresentationSpec(result.jsonSpec).presentation
    if (!normalized) throw new Error('The imported fixture did not normalize for round-trip export.')
    const roundTripPptx = await buildThemedPptxBytes(normalized)
    const roundTripJson = (await importPowerPointBytes(Buffer.from(roundTripPptx), {
      sourceName: 'round-trip.pptx',
    })).jsonSpec as typeof json
    expect(roundTripJson.presentation.slides[0].elements).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          type: 'image',
          crop: {
            left: 0.13,
            right: 0.13,
            top: 0.07,
            bottom: 0.2,
          },
        }),
        expect.objectContaining({
          type: 'line',
          lineType: 'elbow',
          elbowDirection: 'horizontal-first',
          endArrow: 'triangle',
        }),
      ]),
    )
  })

  it('keeps top-level typography distinct from shapes inside groups', async () => {
    const zip = await JSZip.loadAsync(source)
    const titleSlidePath = 'ppt/slides/slide2.xml'
    const titleSlide = await zip.file(titleSlidePath)?.async('text')
    if (!titleSlide) {
      throw new Error('Missing generated title slide')
    }
    zip.file(
      titleSlidePath,
      titleSlide.replace('</p:spTree>', `${GROUP_WITH_CONFLICTING_LOCAL_PATH_XML}</p:spTree>`),
    )
    const groupedSource = await zip.generateAsync({ type: 'nodebuffer' })
    const result = await importPowerPointBytes(groupedSource, {
      sourceName: 'grouped-source.pptx',
      slide: 2,
    })
    const elements = result.jsonSpec.presentation.slides[0]?.elements ?? []
    const importedTitle = elements.find(
      (element): element is PowerPointCanvasTextElement =>
        element.type === 'text' && element.text === 'Imported title',
    )
    const groupedAnnotation = elements.find(
      (element): element is PowerPointCanvasTextElement =>
        element.type === 'text' && element.text === 'Grouped annotation',
    )

    expect(importedTitle).toMatchObject({
      fontFace: 'Aptos Display',
      fontSize: 28,
      textColor: '070154',
      bold: true,
    })
    expect(groupedAnnotation).toMatchObject({
      fontFace: 'Arial',
      fontSize: 12,
      textColor: 'FFFFFF',
    })
  })

  it('distributes PowerPoint auto-sized table rows by their rendered text height', async () => {
    const json = (await importPowerPointBytes(source, {
      sourceName: 'source.pptx',
      slide: 3,
    })).jsonSpec as {
      presentation: {
        slides: Array<{
          elements: PowerPointCanvasElement[]
        }>
      }
    }
    const elements = json.presentation.slides[0].elements
    const cell = (text: string) => elements.find(
      (element): element is PowerPointCanvasShapeElement | PowerPointCanvasTextElement =>
        'text' in element && element.text === text,
    )
    const header = cell('Approach')
    const provider = cell('Hosting Provider')
    const methodology = cell('Hosting Methodology')
    const availability = cell('Availability')
    const recovery = cell('Disaster Recovery')

    expect(header).toBeDefined()
    expect(provider).toBeDefined()
    expect(methodology).toBeDefined()
    expect(availability).toBeDefined()
    expect(recovery).toBeDefined()
    expect(header?.h).toBeCloseTo(46.08, 1)
    expect(provider?.h).toBeGreaterThan(20)
    expect(methodology?.h).toBeGreaterThan(provider?.h ?? 0)
    expect(availability?.h).toBeGreaterThan(methodology?.h ?? 0)
    expect(recovery?.h).toBeCloseTo(methodology?.h ?? 0, 1)
    expect(provider?.y).toBeCloseTo((header?.y ?? 0) + (header?.h ?? 0), 1)
    expect(methodology?.y).toBeCloseTo((provider?.y ?? 0) + (provider?.h ?? 0), 1)
    expect(availability?.y).toBeCloseTo((methodology?.y ?? 0) + (methodology?.h ?? 0), 1)
    expect(recovery?.y).toBeCloseTo((availability?.y ?? 0) + (availability?.h ?? 0), 1)
    expect((recovery?.y ?? 0) + (recovery?.h ?? 0)).toBeCloseTo(556.8, 1)
  })
})

const GROUP_WITH_CONFLICTING_LOCAL_PATH_XML = `
  <p:grpSp>
    <p:nvGrpSpPr>
      <p:cNvPr id="900" name="Typography collision group"/>
      <p:cNvGrpSpPr/>
      <p:nvPr/>
    </p:nvGrpSpPr>
    <p:grpSpPr>
      <a:xfrm>
        <a:off x="0" y="0"/>
        <a:ext cx="914400" cy="457200"/>
        <a:chOff x="0" y="0"/>
        <a:chExt cx="914400" cy="457200"/>
      </a:xfrm>
    </p:grpSpPr>
    <p:sp>
      <p:nvSpPr>
        <p:cNvPr id="901" name="Grouped annotation"/>
        <p:cNvSpPr txBox="1"/>
        <p:nvPr/>
      </p:nvSpPr>
      <p:spPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="914400" cy="457200"/>
        </a:xfrm>
        <a:noFill/>
        <a:ln><a:noFill/></a:ln>
      </p:spPr>
      <p:txBody>
        <a:bodyPr/>
        <a:lstStyle/>
        <a:p>
          <a:r>
            <a:rPr lang="en-US" sz="1200">
              <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
              <a:latin typeface="Arial"/>
            </a:rPr>
            <a:t>Grouped annotation</a:t>
          </a:r>
        </a:p>
      </p:txBody>
    </p:sp>
  </p:grpSp>
`
