import {
  buildThemedPptxBytes,
  normalizePresentationSpec,
} from '../src/lib/export/PowerpointGenerator'
import type { ThrownValue } from '../src/lib/shared/PowerpointTypes'
import { readStdin, writeStdout } from './stdio'

async function main() {
  if (process.argv.length > 2) {
    throw new Error('This command accepts presentation JSON on stdin and returns PPTX bytes on stdout.')
  }

  const raw = await readStdin(50 * 1024 * 1024)
  const { presentation, issues } = normalizePresentationSpec(JSON.parse(raw.toString('utf8')))
  const errors = issues.filter((issue) => issue.level === 'error')
  if (!presentation || errors.length > 0) {
    const formattedIssues = issues
      .map((issue) => `${issue.level.toUpperCase()} ${issue.path}: ${issue.message}`)
      .join('\n')
    throw new Error(`The JSON could not be converted into a PowerPoint deck.\n${formattedIssues}`)
  }

  for (const issue of issues.filter((issue) => issue.level === 'warning')) {
    console.error(`WARNING ${issue.path}: ${issue.message}`)
  }

  writeStdout(Buffer.from(await buildThemedPptxBytes(presentation)))
}

main().catch((error: ThrownValue) => {
  const message = error instanceof Error ? error.message : String(error)
  console.error(message)
  process.exitCode = 1
})
