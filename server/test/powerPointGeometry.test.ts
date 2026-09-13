// @vitest-environment node

import { describe, expect, it } from 'vitest'

import {
  composeTransform,
  extractElementTransform,
  groupTransform,
  identityTransform,
} from '../src/lib/import/PowerpointGeometry'
import type { TransformMatrix } from '../src/lib/import/PowerpointImportTypes'
import { findDescendant, parseXml } from '../src/lib/import/PowerpointXml'
import type { XmlNode } from '../src/lib/shared/PowerpointTypes'

const EMU_PER_INCH = 914_400

function requiredNode(xml: string, tag: string): XmlNode {
  const node = findDescendant(parseXml(xml), tag)
  if (!node) {
    throw new Error(`Expected the fixture to contain ${tag}.`)
  }
  return node
}

function expectMatrix(actual: TransformMatrix, expected: TransformMatrix) {
  for (const key of ['a', 'b', 'c', 'd', 'e', 'f'] as const) {
    expect(actual[key]).toBeCloseTo(expected[key], 8)
  }
}

describe('PowerPoint transform geometry', () => {
  it.each([
    {
      name: 'applies a nested translation before the parent scale and translation',
      parent: { a: 2, b: 0, c: 0, d: 3, e: 10, f: -20 },
      child: { a: 1, b: 0, c: 0, d: 1, e: -5, f: 4 },
      expected: { a: 2, b: 0, c: 0, d: 3, e: 0, f: -8 },
    },
    {
      name: 'composes a nested translation through a 90 degree parent rotation',
      parent: { a: 0, b: 1, c: -1, d: 0, e: 100, f: 0 },
      child: { a: 1, b: 0, c: 0, d: 1, e: -20, f: 30 },
      expected: { a: 0, b: 1, c: -1, d: 0, e: 70, f: -20 },
    },
  ])('$name', ({ parent, child, expected }) => {
    expectMatrix(composeTransform(parent, child), expected)
  })

  it.each([
    {
      name: 'rotates 90 degrees around the group center',
      attributes: 'rot="5400000"',
      expected: { a: 0, b: 1, c: -1, d: 0, e: EMU_PER_INCH, f: 0 },
    },
    {
      name: 'reflects horizontally around the group center',
      attributes: 'flipH="1"',
      expected: { a: -1, b: 0, c: 0, d: 1, e: EMU_PER_INCH, f: 0 },
    },
  ])('$name', ({ attributes, expected }) => {
    const group = requiredNode(
      `<p:grpSp>
        <p:grpSpPr>
          <a:xfrm ${attributes}>
            <a:off x="0" y="0"/>
            <a:ext cx="${EMU_PER_INCH}" cy="${EMU_PER_INCH}"/>
            <a:chOff x="0" y="0"/>
            <a:chExt cx="${EMU_PER_INCH}" cy="${EMU_PER_INCH}"/>
          </a:xfrm>
        </p:grpSpPr>
      </p:grpSp>`,
      'p:grpSp',
    )

    expectMatrix(groupTransform(group), expected)
  })

  it.each([
    {
      name: 'preserves negative coordinates and local flips',
      matrix: identityTransform(),
      shapeAttributes: 'flipH="1" flipV="1"',
      offset: { x: -EMU_PER_INCH, y: -EMU_PER_INCH / 2 },
      extent: { w: EMU_PER_INCH * 2, h: EMU_PER_INCH },
      expected: {
        xPx: -96,
        yPx: -48,
        widthPx: 192,
        heightPx: 96,
        rotation: 0,
        flipH: true,
        flipV: true,
      },
    },
    {
      name: 'positions and rotates a child through a 90 degree parent transform',
      matrix: { a: 0, b: 1, c: -1, d: 0, e: EMU_PER_INCH, f: 0 },
      shapeAttributes: '',
      offset: { x: 0, y: 0 },
      extent: { w: EMU_PER_INCH / 2, h: EMU_PER_INCH / 2 },
      expected: {
        xPx: 48,
        yPx: 0,
        widthPx: 48,
        heightPx: 48,
        rotation: 90,
        flipH: false,
        flipV: false,
      },
    },
  ])('$name', ({ matrix, shapeAttributes, offset, extent, expected }) => {
    const shape = requiredNode(
      `<p:sp>
        <p:spPr>
          <a:xfrm ${shapeAttributes}>
            <a:off x="${offset.x}" y="${offset.y}"/>
            <a:ext cx="${extent.w}" cy="${extent.h}"/>
          </a:xfrm>
        </p:spPr>
      </p:sp>`,
      'p:sp',
    )

    expect(extractElementTransform(shape, matrix)).toMatchObject(expected)
  })
})
