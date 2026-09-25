import PptxGenJS from 'pptxgenjs'
import {
  getWestMonroeBrandFrameLayout,
  WEST_MONROE_BRAND_COLOR,
} from '../WestMonroeBrandFrame'
import { DEFAULT_FONT_FACE, DEFAULT_OPACITY } from '../shared/PowerpointConstants'
import type {
  NormalizedImageElement,
  NormalizedLineElement,
  NormalizedPresentation,
  NormalizedTextRun,
} from '../shared/PowerpointTypes'
import { cleanHex, normalizeElbowDirection } from '../shared/PowerpointUtils'
import {
  PPTX_DEFAULT_BACKGROUND_COLOR,
  PPTX_DEFAULT_LINE_COLOR,
  PPTX_DEFAULT_TEXT_COLOR,
  PPTX_FULL_TRANSPARENCY,
  PPTX_LINE_SPACING_MULTIPLE,
  PPTX_MIN_VISIBLE_CROP_FRACTION,
} from './PowerpointConstants'
import {
  opacityToTransparency,
  pxToInches,
  toPptxShapeName,
  toPptxVerticalAlign,
} from './PowerpointUtils'
import { getConnectorAwareElementOrder } from '../shared/PowerpointLayering'
import { WEST_MONROE_LOGO_DATA } from './WestMonroeLogo'

// PptxGenJS writes arbitrary OOXML preset names correctly, but its public enum omits connectors.
const PPTX_ELBOW_CONNECTOR_SHAPE = 'bentConnector2' as PptxGenJS.ShapeType
const PPTX_STRAIGHT_CONNECTOR_SHAPE = 'line' as PptxGenJS.ShapeType

export function buildPptxPresentation(presentation: NormalizedPresentation) {
  const pptx = new PptxGenJS()
  const widthInches = pxToInches(presentation.meta.width)
  const heightInches = pxToInches(presentation.meta.height)
  const brandFrame = getWestMonroeBrandFrameLayout(
    presentation.meta.width,
    presentation.meta.height,
  )

  pptx.defineLayout({
    name: 'JSON_LAYOUT',
    width: widthInches,
    height: heightInches,
  })
  pptx.layout = 'JSON_LAYOUT'
  pptx.author = 'OpenAI Codex'
  pptx.company = 'OpenAI'
  pptx.subject = 'JSON to PowerPoint'
  pptx.title = presentation.meta.title
  pptx.theme = {
    headFontFace: DEFAULT_FONT_FACE,
    bodyFontFace: DEFAULT_FONT_FACE,
  }
  if (presentation.meta.showBranding) {
    pptx.defineSlideMaster({
      title: 'WEST_MONROE_BRANDED_FRAME',
      objects: [
        {
          rect: {
            x: pxToInches(brandFrame.footer.x),
            y: pxToInches(brandFrame.footer.y),
            w: pxToInches(brandFrame.footer.w),
            h: pxToInches(brandFrame.footer.h),
            fill: { color: WEST_MONROE_BRAND_COLOR },
            line: { transparency: PPTX_FULL_TRANSPARENCY },
          },
        },
        ...brandFrame.dots.map((dot) => ({
          rect: {
            x: pxToInches(dot.x),
            y: pxToInches(dot.y),
            w: pxToInches(dot.w),
            h: pxToInches(dot.h),
            fill: { color: WEST_MONROE_BRAND_COLOR },
            line: { transparency: PPTX_FULL_TRANSPARENCY },
          },
        })),
        {
          image: {
            data: WEST_MONROE_LOGO_DATA,
            x: pxToInches(brandFrame.logo.x),
            y: pxToInches(brandFrame.logo.y),
            w: pxToInches(brandFrame.logo.w),
            h: pxToInches(brandFrame.logo.h),
          },
        },
      ],
    })
  }
  for (const slideSpec of presentation.slides) {
    const slide = presentation.meta.showBranding
      ? pptx.addSlide('WEST_MONROE_BRANDED_FRAME')
      : pptx.addSlide()
    slide.background = { color: cleanHex(slideSpec.backgroundColor, PPTX_DEFAULT_BACKGROUND_COLOR) }

    const elements = slideSpec.preserveElementOrder
      ? slideSpec.elements
      : getConnectorAwareElementOrder(slideSpec)
    for (const element of elements) {
      if (element.kind === 'line') {
        const lineShape = toPptxLineShape(element)
        slide.addShape(lineShape.shape, {
          x: pxToInches(lineShape.geometry.x),
          y: pxToInches(lineShape.geometry.y),
          w: pxToInches(lineShape.geometry.w),
          h: pxToInches(lineShape.geometry.h),
          flipH: lineShape.geometry.flipH,
          flipV: lineShape.geometry.flipV,
          rotate: element.rotate,
          line: {
            color: cleanHex(element.stroke, PPTX_DEFAULT_LINE_COLOR),
            width: element.strokeWidth,
            transparency: opacityToTransparency(
              combinedOpacity(element.strokeOpacity, element.opacity),
            ),
            dashType:
              element.dash === 'solid' ? 'solid' : element.dash === 'dot' ? 'sysDot' : 'dash',
            beginArrowType: lineShape.beginArrow,
            endArrowType: lineShape.endArrow,
          },
        })
        continue
      }

      if (element.kind === 'image') {
        const imageOptions = buildImageOptions(element)
        slide.addImage(imageOptions)
        continue
      }

      if (element.kind === 'text') {
        slide.addText(toPptxTextRuns(element.runs), {
          x: pxToInches(element.x),
          y: pxToInches(element.y),
          w: pxToInches(element.w),
          h: pxToInches(element.h),
          margin: [element.padding, element.padding, element.padding, element.padding],
          fontFace: element.fontFace,
          fontSize: element.fontSize,
          color: cleanHex(element.color, PPTX_DEFAULT_TEXT_COLOR),
          bold: element.runs.length ? element.runs.every((run) => run.bold) : element.bold,
          italic: element.runs.length ? element.runs.every((run) => run.italic) : element.italic,
          align: element.align,
          valign: toPptxVerticalAlign(element.valign),
          lineSpacingMultiple: PPTX_LINE_SPACING_MULTIPLE,
          paraSpaceAfter: 0,
          paraSpaceBefore: 0,
          fill: colorToFill(element.fill, combinedOpacity(element.fillOpacity, element.opacity)),
          line: colorToLine(
            element.stroke,
            element.strokeWidth,
            combinedOpacity(element.strokeOpacity, element.opacity),
          ),
          rotate: element.rotate,
          flipH: element.flipH,
          flipV: element.flipV,
          fit: 'shrink',
          isTextBox: true,
          shape: element.borderRadius > 0 ? 'roundRect' : 'rect',
          rectRadius: element.borderRadius > 0 ? pxToInches(element.borderRadius) : undefined,
        })
        continue
      }

      if (element.label.trim().length > 0) {
        slide.addText(toPptxTextRuns(element.textRuns), {
          x: pxToInches(element.x),
          y: pxToInches(element.y),
          w: pxToInches(element.w),
          h: pxToInches(element.h),
          margin: [element.padding, element.padding, element.padding, element.padding],
          fontFace: element.fontFace,
          fontSize: element.fontSize,
          color: cleanHex(element.textColor, PPTX_DEFAULT_TEXT_COLOR),
          bold: element.bold,
          align: element.align,
          valign: toPptxVerticalAlign(element.valign),
          rotate: element.rotate,
          flipH: element.flipH,
          flipV: element.flipV,
          fit: 'shrink',
          fill: colorToFill(element.fill, combinedOpacity(element.fillOpacity, element.opacity)),
          line: colorToLine(
            element.stroke,
            element.strokeWidth,
            combinedOpacity(element.strokeOpacity, element.opacity),
          ),
          shape:
            element.shape === 'rect' && element.borderRadius > 0
              ? 'roundRect'
              : toPptxShapeName(element.shape),
          rectRadius: element.borderRadius > 0 ? pxToInches(element.borderRadius) : undefined,
        })
        continue
      }

      slide.addShape(toPptxShapeName(element.shape), {
        x: pxToInches(element.x),
        y: pxToInches(element.y),
        w: pxToInches(element.w),
        h: pxToInches(element.h),
        rotate: element.rotate,
        flipH: element.flipH,
        flipV: element.flipV,
        fill: colorToFill(element.fill, combinedOpacity(element.fillOpacity, element.opacity)),
        line: colorToLine(
          element.stroke,
          element.strokeWidth,
          combinedOpacity(element.strokeOpacity, element.opacity),
        ),
        rectRadius:
          element.shape === 'roundRect' && element.borderRadius > 0
            ? pxToInches(element.borderRadius)
            : undefined,
      })

    }
  }

  return pptx
}

function toPptxLineShape(element: NormalizedLineElement) {
  const deltaX = Math.abs(element.x2 - element.x1)
  const deltaY = Math.abs(element.y2 - element.y1)
  const isElbow = element.lineType === 'elbow' && deltaX > 0.5 && deltaY > 0.5
  // bentConnector2 travels horizontally first. Reversing its endpoints and arrowheads
  // produces the same visible connector as a vertical-first route.
  const elbowDirection = normalizeElbowDirection(
    element.elbowDirection,
    element.x1,
    element.y1,
    element.x2,
    element.y2,
  )
  const reverse = isElbow && elbowDirection === 'vertical-first'
  const points = reverse
    ? { x1: element.x2, y1: element.y2, x2: element.x1, y2: element.y1 }
    : element

  return {
    beginArrow: reverse ? element.endArrow : element.beginArrow,
    endArrow: reverse ? element.beginArrow : element.endArrow,
    geometry: toPptxLineGeometry(points),
    shape: isElbow ? PPTX_ELBOW_CONNECTOR_SHAPE : PPTX_STRAIGHT_CONNECTOR_SHAPE,
  }
}

function toPptxLineGeometry(element: Pick<NormalizedLineElement, 'x1' | 'x2' | 'y1' | 'y2'>) {
  return {
    x: Math.min(element.x1, element.x2),
    y: Math.min(element.y1, element.y2),
    w: Math.abs(element.x2 - element.x1),
    h: Math.abs(element.y2 - element.y1),
    flipH: element.x2 < element.x1,
    flipV: element.y2 < element.y1,
  }
}

function buildImageOptions(element: NormalizedImageElement) {
  const crop = element.crop
  const visibleWidth = crop
    ? Math.max(1 - crop.left - crop.right, PPTX_MIN_VISIBLE_CROP_FRACTION)
    : DEFAULT_OPACITY
  const visibleHeight = crop
    ? Math.max(1 - crop.top - crop.bottom, PPTX_MIN_VISIBLE_CROP_FRACTION)
    : DEFAULT_OPACITY
  const sourceBoxWidth = element.w / visibleWidth
  const sourceBoxHeight = element.h / visibleHeight
  const base = {
    x: pxToInches(element.x),
    y: pxToInches(element.y),
    w: pxToInches(sourceBoxWidth),
    h: pxToInches(sourceBoxHeight),
    altText: element.altText,
    transparency: opacityToTransparency(element.opacity),
    rotate: element.rotate,
    flipH: element.flipH,
    flipV: element.flipV,
    sizing: crop
      ? {
          type: 'crop' as const,
          x: pxToInches(crop.left * sourceBoxWidth),
          y: pxToInches(crop.top * sourceBoxHeight),
          w: pxToInches(element.w),
          h: pxToInches(element.h),
        }
      : undefined,
  }

  if (!element.src.startsWith('data:')) {
    throw new Error('PowerPoint image sources must be embedded data URIs.')
  }

  return {
    ...base,
    data: element.src,
  }
}

function toPptxTextRuns(runs: NormalizedTextRun[]) {
  return runs.map((run, index) => ({
    text: run.text,
    options: {
      bold: run.bold,
      italic: run.italic,
      underline: run.underline ? {} : undefined,
      breakLine: run.breakLine && index < runs.length - 1,
      color: cleanHex(run.color, PPTX_DEFAULT_TEXT_COLOR),
      fontFace: run.fontFace,
      fontSize: run.fontSize,
    },
  }))
}

function combinedOpacity(localOpacity: number | undefined, elementOpacity: number) {
  return (localOpacity ?? DEFAULT_OPACITY) * elementOpacity
}

function colorToFill(color: string, opacity = DEFAULT_OPACITY) {
  if (color === 'transparent') {
    return { color: PPTX_DEFAULT_BACKGROUND_COLOR, transparency: PPTX_FULL_TRANSPARENCY }
  }

  return {
    color: cleanHex(color, PPTX_DEFAULT_BACKGROUND_COLOR),
    transparency: opacityToTransparency(opacity),
  }
}

function colorToLine(color: string, width: number, opacity = DEFAULT_OPACITY) {
  if (color === 'transparent' || width <= 0) {
    return {
      color: PPTX_DEFAULT_BACKGROUND_COLOR,
      transparency: PPTX_FULL_TRANSPARENCY,
      width: 0,
    }
  }

  return {
    color: cleanHex(color, PPTX_DEFAULT_LINE_COLOR),
    transparency: opacityToTransparency(opacity),
    width,
  }
}
