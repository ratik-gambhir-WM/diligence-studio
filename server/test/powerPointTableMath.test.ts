// @vitest-environment node

import { describe, expect, it } from 'vitest'

import { identityTransform } from '../src/lib/import/PowerpointGeometry'
import { extractTableElements } from '../src/lib/import/PowerpointTableExtractor'
import { findDescendant, parseXml } from '../src/lib/import/PowerpointXml'

const MIXED_SIZE_TABLE_XML = `
  <p:graphicFrame>
    <p:xfrm>
      <a:off x="0" y="0"/>
      <a:ext cx="5486400" cy="3657600"/>
    </p:xfrm>
    <a:graphic><a:graphicData><a:tbl>
      <a:tblGrid>
        <a:gridCol w="914400"/>
        <a:gridCol w="0"/>
        <a:gridCol w="1828800"/>
      </a:tblGrid>
      <a:tr h="914400">
        <a:tc><a:txBody><a:p><a:r><a:t>A1</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
        <a:tc><a:txBody><a:p><a:r><a:t>B1</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
        <a:tc><a:txBody><a:p><a:r><a:t>C1</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
      </a:tr>
      <a:tr h="0">
        <a:tc><a:txBody><a:p><a:r><a:t>A2</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
        <a:tc><a:txBody><a:p><a:r><a:t>B2</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
        <a:tc><a:txBody><a:p><a:r><a:t>C2</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
      </a:tr>
      <a:tr h="914400">
        <a:tc><a:txBody><a:p><a:r><a:t>A3</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
        <a:tc><a:txBody><a:p><a:r><a:t>B3</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
        <a:tc><a:txBody><a:p><a:r><a:t>C3</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
      </a:tr>
    </a:tbl></a:graphicData></a:graphic>
  </p:graphicFrame>
`

describe('PowerPoint table sizing', () => {
  it('distributes remaining space to automatic rows and columns while preserving explicit sizes', () => {
    const frame = findDescendant(parseXml(MIXED_SIZE_TABLE_XML), 'p:graphicFrame')
    if (!frame) {
      throw new Error('Expected the fixture to contain a graphic frame.')
    }

    const elements = extractTableElements(frame, identityTransform(), 'slide', 0)
    expect(elements).toHaveLength(9)

    const expectedCells = [
      { index: 0, xPx: 0, yPx: 0, widthPx: 96, heightPx: 96 },
      { index: 1, xPx: 96, yPx: 0, widthPx: 288, heightPx: 96 },
      { index: 2, xPx: 384, yPx: 0, widthPx: 192, heightPx: 96 },
      { index: 3, xPx: 0, yPx: 96, widthPx: 96, heightPx: 192 },
      { index: 4, xPx: 96, yPx: 96, widthPx: 288, heightPx: 192 },
      { index: 5, xPx: 384, yPx: 96, widthPx: 192, heightPx: 192 },
      { index: 6, xPx: 0, yPx: 288, widthPx: 96, heightPx: 96 },
      { index: 7, xPx: 96, yPx: 288, widthPx: 288, heightPx: 96 },
      { index: 8, xPx: 384, yPx: 288, widthPx: 192, heightPx: 96 },
    ]

    for (const { index, ...expectedTransform } of expectedCells) {
      expect(elements[index]?.transform).toMatchObject(expectedTransform)
    }
  })
})
