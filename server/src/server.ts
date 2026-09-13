import { createServer } from 'node:http'

import { writeApiLogToConsole } from './apiLogging'
import { createApp } from './app'
import { seedBuiltinTemplates } from './catalog/seedBuiltinTemplates'
import { loadServerConfig } from './config'
import { SqliteTemplateRepository } from './repositories/SqliteTemplateRepository'
import { ExportService } from './services/ExportPowerPointService'
import { ImportService } from './services/ImportTemplateService'
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
const importService = new ImportService(
  new LibraryPowerPointConverter(),
  templates,
  undefined,
  undefined,
  previewGenerator,
)
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
    `PowerPoint API listening at ${config.host}:${config.port}; template previews: ${config.previewProvider}.`,
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

  server.close((error) => {
    clearTimeout(forceCloseTimer)
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
