import { createServer } from 'node:http'

import OpenAI from 'openai'

import { writeApiLogToConsole } from './apiLogging'
import { createApp } from './app'
import { seedBuiltinTemplates } from './catalog/seedBuiltinTemplates'
import { loadServerConfig } from './config'
import { OpenAISlideClassifier } from './integrations/OpenAISlideClassifier'
import { OpenAISlideEmbedder } from './integrations/OpenAISlideEmbedder'
import { SqliteTemplateRepository } from './repositories/SqliteTemplateRepository'
import { ExportService } from './services/ExportPowerPointService'
import { ImportService } from './services/ImportTemplateService'
import { SlideClassificationService } from './services/SlideClassificationService'
import { SlideClassificationWorker } from './services/SlideClassificationWorker'
import { LibraryPowerPointConverter } from './services/PowerPointConverter'
import {
  DisabledTemplatePreviewGenerator,
  HeadlessTemplatePreviewGenerator,
  QuickLookTemplatePreviewGenerator,
} from './services/TemplatePreview'

const config = loadServerConfig()
const templates = new SqliteTemplateRepository(config.databasePath)
await seedBuiltinTemplates(templates)
const exportService = new ExportService(templates)
const previewOptions = {
  maxBytes: config.maxPreviewBytes,
  renderSize: config.previewRenderSize,
  timeoutMs: config.previewTimeoutMs,
}
const previewGenerator = config.previewProvider === 'headless'
  ? new HeadlessTemplatePreviewGenerator({
      ...previewOptions,
      renderUrl: config.previewRenderUrl,
    })
  : config.previewProvider === 'quicklook'
    ? new QuickLookTemplatePreviewGenerator(previewOptions)
    : new DisabledTemplatePreviewGenerator()
let classificationWorker: SlideClassificationWorker | undefined
const importService = new ImportService(
  new LibraryPowerPointConverter(),
  templates,
  undefined,
  undefined,
  previewGenerator,
  () => classificationWorker?.notify(),
)

if (
  config.slideClassificationProvider === 'openai'
  && config.openaiApiKey
  && config.openaiSlideClassificationModel
  && config.openaiSlideEmbeddingModel
) {
  const openai = new OpenAI({ apiKey: config.openaiApiKey, maxRetries: 0 })
  const classificationService = new SlideClassificationService({
    classificationModel: config.openaiSlideClassificationModel,
    classifier: new OpenAISlideClassifier({
      client: openai,
      model: config.openaiSlideClassificationModel,
      timeoutMs: config.slideClassificationTimeoutMs,
    }),
    embedder: new OpenAISlideEmbedder({
      client: openai,
      dimensions: config.openaiSlideEmbeddingDimensions,
      maxInputBytes: config.slideEmbeddingMaxTextBytes,
      model: config.openaiSlideEmbeddingModel,
      timeoutMs: config.slideEmbeddingTimeoutMs,
    }),
    embeddingDimensions: config.openaiSlideEmbeddingDimensions,
    embeddingMaxTextBytes: config.slideEmbeddingMaxTextBytes,
    embeddingModel: config.openaiSlideEmbeddingModel,
    inputLimits: {
      maxPreviewBytes: config.slideClassificationMaxImageBytes,
      maxTextChars: config.slideClassificationMaxTextChars,
    },
    repository: templates,
  })
  classificationWorker = new SlideClassificationWorker({
    classificationMaxAttempts: config.slideClassificationMaxAttempts,
    classificationTimeoutMs: config.slideClassificationTimeoutMs,
    concurrency: config.slideClassificationConcurrency,
    embeddingMaxAttempts: config.slideEmbeddingMaxAttempts,
    embeddingTimeoutMs: config.slideEmbeddingTimeoutMs,
    repository: templates,
    service: classificationService,
  })
  classificationWorker.start()
}
const server = createServer(createApp({
  exportService,
  importService,
  logger: writeApiLogToConsole,
  maxExportJsonBytes: config.maxExportJsonBytes,
  maxUploadBytes: config.maxUploadBytes,
  requestTimeoutMs: config.requestTimeoutMs,
}))

server.listen(config.port, config.host, () => {
  console.log(
    `PowerPoint API listening at ${config.host}:${config.port}; template previews: ${config.previewProvider}; slide classification: ${config.slideClassificationProvider}${config.openaiSlideClassificationModel ? ` (${config.openaiSlideClassificationModel}, concurrency ${config.slideClassificationConcurrency})` : ''}.`,
  )
})

let shuttingDown = false
function shutdown() {
  if (shuttingDown) {
    return
  }
  shuttingDown = true

  const forceCloseTimer = setTimeout(() => {
    server.closeAllConnections()
  }, 10_000)
  forceCloseTimer.unref()

  server.close(async (error) => {
    clearTimeout(forceCloseTimer)
    await classificationWorker?.stop()
    templates.close()
    if (error) {
      console.error('The PowerPoint API did not shut down cleanly.')
      process.exitCode = 1
    }
  })
  server.closeIdleConnections()
}

process.on('SIGINT', shutdown)
process.on('SIGTERM', shutdown)
