import { compactPresentation } from '../src/lib/import/PowerpointCompactWriter'
import { normalizePresentationSpec } from '../src/lib/shared/PowerpointNormalizer'
import type { ThrownValue } from '../src/lib/shared/PowerpointTypes'
import { readStdin, writeStdout } from './stdio'

async function main() {
  if (process.argv.length > 2) {
    throw new Error('This command accepts presentation JSON on stdin and returns compact JSON on stdout.')
  }

  const raw = await readStdin(50 * 1024 * 1024)
  const { presentation, issues } = normalizePresentationSpec(JSON.parse(raw.toString('utf8')))
  const errors = issues.filter((issue) => issue.level === 'error')
  if (!presentation || errors.length > 0) {
    const formattedIssues = issues
      .map((issue) => `${issue.level.toUpperCase()} ${issue.path}: ${issue.message}`)
      .join('\n')
    throw new Error(`The JSON could not be condensed.\n${formattedIssues}`)
  }

  writeStdout(Buffer.from(`${JSON.stringify(compactPresentation(presentation), null, 2)}\n`))
  for (const issue of issues.filter((issue) => issue.level === 'warning')) {
    console.error(`WARNING ${issue.path}: ${issue.message}`)
  }
}

main().catch((error: ThrownValue) => {
  const message = error instanceof Error ? error.message : String(error)
  console.error(message)
  process.exitCode = 1
})
