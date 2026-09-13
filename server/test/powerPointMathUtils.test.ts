// @vitest-environment node

import { describe, expect, it } from 'vitest'

import { opacityToTransparency, pxToInches } from '../src/lib/export/PowerpointUtils'
import { emuToPx, round } from '../src/lib/import/PowerpointImportUtils'
import { getConnectorAwareElementOrder } from '../src/lib/shared/PowerpointLayering'
import { normalizePresentationSpec } from '../src/lib/shared/PowerpointNormalizer'
import type {
  JsonObject,
  NormalizedImageElement,
  NormalizedLineElement,
  NormalizedShapeElement,
  XmlNode,
} from '../src/lib/shared/PowerpointTypes'
import {
  emuLineWidthToPoints,
  emuToPoints,
  inchesToPx,
  parseColorOpacity,
} from '../src/lib/shared/PowerpointUtils'

function shapeElement(id: string, w: number, h: number): NormalizedShapeElement {
  return {
    kind: 'shape',
    id,
    sourcePath: id,
    opacity: 1,
    rotate: 0,
    valign: 'middle',
    x: 0,
    y: 0,
    w,
    h,
    shape: 'rect',
    label: '',
    fill: 'FFFFFF',
    stroke: '000000',
    strokeWidth: 1,
    borderRadius: 0,
    padding: 0,
    align: 'left',
    textColor: '000000',
    fontSize: 12,
    fontFace: 'Arial',
    bold: false,
    textRuns: [],
  }
}

function lineElement(id: string): NormalizedLineElement {
  return {
    kind: 'line',
    id,
    sourcePath: id,
    opacity: 1,
    rotate: 0,
    valign: 'middle',
    lineType: 'straight',
    x1: 0,
    y1: 0,
    x2: 100,
    y2: 100,
    stroke: '000000',
    strokeWidth: 1,
    dash: 'solid',
    endArrow: 'none',
    occlusionRects: [],
  }
}

function normalizeImage(crop: JsonObject, opacity = 1): NormalizedImageElement {
  const { presentation, issues } = normalizePresentationSpec({
    presentation: {
      showBranding: false,
      slides: [{
        elements: [{
          type: 'image',
          src: 'data:image/png;base64,AA==',
          x: 0,
          y: 0,
          w: 100,
          h: 100,
          crop,
          opacity,
        }],
      }],
    },
  })
  expect(issues.filter((issue) => issue.level === 'error')).toEqual([])
  const element = presentation?.slides[0]?.elements[0]
  if (element?.kind !== 'image') {
    throw new Error('Expected normalization to return an image element.')
  }
  return element
}

function solidFill(modifiers: XmlNode[]): XmlNode {
  return {
    tag: 'a:solidFill',
    children: [{ tag: 'a:srgbClr', attributes: { val: 'FFFFFF' }, children: modifiers }],
  }
}

describe('PowerPoint crop and opacity boundaries', () => {
  it.each([
    {
      name: 'drops an all-zero crop',
      input: { top: 0, right: 0, bottom: 0, left: 0 },
      expected: undefined,
    },
    {
      name: 'clamps crop edges to the supported zero-to-one range',
      input: { top: -0.25, right: 0.4, bottom: 1.5, left: 0 },
      expected: { top: 0, right: 0.4, bottom: 1, left: 0 },
    },
    {
      name: 'falls back to zero for non-numeric crop edges',
      input: { top: 'invalid', right: 0.25, bottom: null, left: false },
      expected: { top: 0, right: 0.25, bottom: 0, left: 0 },
    },
  ])('$name', ({ input, expected }) => {
    expect(normalizeImage(input).crop).toEqual(expected)
  })

  it.each([
    { opacity: -0.2, expected: 0 },
    { opacity: 0.35, expected: 0.35 },
    { opacity: 1.2, expected: 1 },
  ])('clamps normalized image opacity $opacity to $expected', ({ opacity, expected }) => {
    expect(normalizeImage({}, opacity).opacity).toBe(expected)
  })

  it.each([
    { name: 'defaults missing color data to opaque', fill: undefined, expected: 1 },
    {
      name: 'combines alpha, alpha modulation, and alpha offset in document order',
      fill: solidFill([
        { tag: 'a:alpha', attributes: { val: '50000' }, children: [] },
        { tag: 'a:alphaMod', attributes: { val: '50000' }, children: [] },
        { tag: 'a:alphaOff', attributes: { val: '10000' }, children: [] },
      ]),
      expected: 0.35,
    },
    {
      name: 'clamps alpha above the supported range',
      fill: solidFill([{ tag: 'a:alpha', attributes: { val: '120000' }, children: [] }]),
      expected: 1,
    },
    {
      name: 'clamps alpha below the supported range',
      fill: solidFill([{ tag: 'a:alpha', attributes: { val: '-100' }, children: [] }]),
      expected: 0,
    },
  ])('$name', ({ fill, expected }) => {
    expect(parseColorOpacity(fill)).toBeCloseTo(expected)
  })

  it.each([
    { opacity: -0.2, expected: 100 },
    { opacity: 0, expected: 100 },
    { opacity: 0.335, expected: 67 },
    { opacity: 1, expected: 0 },
    { opacity: 1.2, expected: 0 },
  ])('maps opacity $opacity to $expected percent transparency', ({ opacity, expected }) => {
    expect(opacityToTransparency(opacity)).toBe(expected)
  })
})

describe('connector-aware element layering', () => {
  it.each([
    { name: 'large-area threshold', w: 400, h: 200, isContainer: true },
    { name: 'just below large-area threshold', w: 399, h: 200, isContainer: false },
    { name: 'tall-lane threshold', w: 120, h: 450, isContainer: true },
    { name: 'just below tall-lane width threshold', w: 119, h: 450, isContainer: false },
  ])('classifies a rectangle at the $name', ({ w, h, isContainer }) => {
    const content = shapeElement('content', 10, 10)
    const candidate = shapeElement('candidate', w, h)
    const connector = lineElement('connector')
    const orderedIds = getConnectorAwareElementOrder({
      width: 1000,
      height: 1000,
      elements: [content, candidate, connector],
    }).map((element) => element.id)

    expect(orderedIds).toEqual(
      isContainer
        ? ['candidate', 'connector', 'content']
        : ['connector', 'content', 'candidate'],
    )
  })

  it('keeps source order stable within every computed layer', () => {
    const elements = [
      shapeElement('content-a', 10, 10),
      lineElement('line-a'),
      shapeElement('container-a', 400, 200),
      shapeElement('west-monroe-dot-a', 1, 1),
      shapeElement('content-b', 10, 10),
      lineElement('line-b'),
      shapeElement('container-b', 400, 200),
      shapeElement('west-monroe-dot-b', 1, 1),
    ]

    const orderedIds = getConnectorAwareElementOrder({
      width: 1000,
      height: 1000,
      elements,
    }).map((element) => element.id)

    expect(orderedIds).toEqual([
      'west-monroe-dot-a',
      'west-monroe-dot-b',
      'container-a',
      'container-b',
      'line-a',
      'line-b',
      'content-a',
      'content-b',
    ])
  })
})

describe('PowerPoint unit conversions and rounding', () => {
  it.each([
    { emu: 914_400, expected: 96 },
    { emu: 457_200, expected: 48 },
    { emu: -457_200, expected: -48 },
  ])('maps $emu EMUs to $expected pixels', ({ emu, expected }) => {
    expect(emuToPx(emu)).toBe(expected)
  })

  it.each([
    { emu: '12700', expected: 1 },
    { emu: '914400', expected: 72 },
    { emu: '0', expected: 0 },
    { emu: 'invalid', expected: 0 },
  ])('maps $emu EMUs to $expected points', ({ emu, expected }) => {
    expect(emuToPoints(emu)).toBe(expected)
  })

  it.each([
    { emu: '12700', expected: 1 },
    { emu: '25400', expected: 2 },
    { emu: '0', expected: 1 },
    { emu: 'invalid', expected: 1 },
  ])('maps line width $emu EMUs to $expected points', ({ emu, expected }) => {
    expect(emuLineWidthToPoints(emu)).toBe(expected)
  })

  it.each([
    { inches: 0, pixels: 0 },
    { inches: 1, pixels: 96 },
    { inches: 2.5, pixels: 240 },
    { inches: -0.5, pixels: -48 },
  ])('round trips $inches inches and $pixels pixels', ({ inches, pixels }) => {
    expect(inchesToPx(inches)).toBe(pixels)
    expect(pxToInches(pixels)).toBe(inches)
  })

  it.each([
    { value: 1.23456, digits: 4, expected: 1.2346 },
    { value: -1.23456, digits: 3, expected: -1.235 },
    { value: 12.4, digits: 0, expected: 12 },
  ])('rounds $value to $digits decimal places', ({ value, digits, expected }) => {
    expect(round(value, digits)).toBe(expected)
  })
})
