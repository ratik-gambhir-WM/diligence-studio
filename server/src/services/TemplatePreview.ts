import { chromium } from 'playwright'

import type { PowerPointCanvasJson } from '../lib/import/PowerpointImportTypes'

export type GeneratedTemplatePreview = {
  bytes: Buffer
  contentType: 'image/png'
  height: number
  width: number
}

export interface TemplatePreviewGenerator {
  generate(
    request: TemplatePreviewRequest,
    signal?: AbortSignal,
  ): Promise<GeneratedTemplatePreview | undefined>
}

export type TemplatePreviewRequest = {
  templateJson: PowerPointCanvasJson
}

export type HeadlessPreviewOptions = {
  maxBytes: number
  maxDimension?: number
  renderSize: number
  renderUrl: string
  timeoutMs: number
}

export type HeadlessPreviewRunner = (
  templateJson: PowerPointCanvasJson,
  options: HeadlessPreviewOptions,
  signal?: AbortSignal,
) => Promise<Buffer>

export class DisabledTemplatePreviewGenerator implements TemplatePreviewGenerator {
  async generate() {
    return undefined
  }
}

export class HeadlessTemplatePreviewGenerator implements TemplatePreviewGenerator {
  private queue: Promise<void> = Promise.resolve()

  constructor(
    private readonly options: HeadlessPreviewOptions,
    private readonly runner: HeadlessPreviewRunner = runHeadlessPreview,
  ) {}

  generate(request: TemplatePreviewRequest, signal?: AbortSignal) {
    const operation = this.queue.then(() => this.generatePreview(request, signal))
    this.queue = operation.then(() => undefined, () => undefined)
    return operation
  }

  private async generatePreview(request: TemplatePreviewRequest, signal?: AbortSignal) {
    try {
      const bytes = await this.runner(request.templateJson, this.options, signal)
      const dimensions = readPngDimensions(bytes)
      const maxDimension = this.options.maxDimension ?? Math.max(this.options.renderSize * 2, 4096)
      if (
        !dimensions
        || bytes.length > this.options.maxBytes
        || dimensions.width > maxDimension
        || dimensions.height > maxDimension
      ) {
        return undefined
      }

      return { bytes, contentType: 'image/png' as const, ...dimensions }
    } catch {
      return undefined
    }
  }
}

/**
 * Renders previews through Quarry's equivalent read-only SVG preview route.
 *
 * The implementation intentionally shares the headless runner and PNG validation
 * with the Diligence Studio renderer so both providers have identical output,
 * cancellation, resource-bounding, and failure behavior.
 */
export class QuarryTemplatePreviewGenerator extends HeadlessTemplatePreviewGenerator {}

export function readPngDimensions(bytes: Buffer) {
  const signature = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10])
  if (
    bytes.length < 24
    || !bytes.subarray(0, signature.length).equals(signature)
    || bytes.toString('ascii', 12, 16) !== 'IHDR'
  ) {
    return undefined
  }

  const width = bytes.readUInt32BE(16)
  const height = bytes.readUInt32BE(20)
  return width > 0 && height > 0 ? { height, width } : undefined
}

async function runHeadlessPreview(
  templateJson: PowerPointCanvasJson,
  options: HeadlessPreviewOptions,
  signal?: AbortSignal,
) {
  if (signal?.aborted) {
    throw new Error('Template preview rendering was cancelled.')
  }

  const browser = await chromium.launch({ chromiumSandbox: true, headless: true })
  const closeBrowser = () => {
    void browser.close()
  }
  signal?.addEventListener('abort', closeBrowser, { once: true })

  try {
    const renderUrl = new URL(options.renderUrl)
    renderUrl.searchParams.set('maxDimension', String(options.renderSize))
    const allowedOrigin = renderUrl.origin
    const context = await browser.newContext({
      deviceScaleFactor: 1,
      viewport: { height: options.renderSize, width: options.renderSize },
    })
    const page = await context.newPage()
    page.setDefaultTimeout(options.timeoutMs)
    await page.route('**/*', async (route) => {
      const requestUrl = new URL(route.request().url())
      if (
        requestUrl.origin === allowedOrigin
        || requestUrl.protocol === 'blob:'
        || requestUrl.protocol === 'data:'
      ) {
        await route.continue()
      } else {
        await route.abort('blockedbyclient')
      }
    })
    await page.addInitScript((input) => {
      Object.defineProperty(window, '__TTS_MERMAID_TEMPLATE_PREVIEW_INPUT__', {
        configurable: true,
        value: input,
      })
    }, templateJson)
    await page.goto(renderUrl.toString(), {
      timeout: options.timeoutMs,
      waitUntil: 'domcontentloaded',
    })

    const state = page.locator('[data-preview-state]')
    await state.waitFor({ state: 'attached', timeout: options.timeoutMs })
    if (await state.getAttribute('data-preview-state') !== 'ready') {
      throw new Error('The template preview page could not render the slide.')
    }

    await page.evaluate(async () => {
      await document.fonts.ready
      await new Promise<void>((resolve) => {
        requestAnimationFrame(() => requestAnimationFrame(() => resolve()))
      })
    })
    return await page.locator('[data-template-preview-surface]').screenshot({
      animations: 'disabled',
      type: 'png',
    })
  } finally {
    signal?.removeEventListener('abort', closeBrowser)
    await browser.close()
  }
}
