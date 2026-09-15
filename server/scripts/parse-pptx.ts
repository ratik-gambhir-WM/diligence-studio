import { importPowerPointBytes } from '../src/lib/import/PowerpointImporter'
import type { ThrownValue } from '../src/lib/shared/PowerpointTypes'
import { readStdin, writeStdout } from './stdio'

type CliOptions = {
  slide?: number
  sourceName?: string
}

function parseCliOptions(args: string[]): CliOptions {
  let slide: number | undefined
  let sourceName: string | undefined

  for (let index = 0; index < args.length; index += 1) {
    const argument = args[index]
    if (argument === '--slide') {
      slide = parseSlide(args[index + 1])
      index += 1
      continue
    }
    if (argument.startsWith('--slide=')) {
      slide = parseSlide(argument.slice('--slide='.length))
      continue
    }
    if (argument === '--source-name') {
      sourceName = args[index + 1]
      if (!sourceName?.trim()) throw new Error('--source-name must be followed by a name.')
      index += 1
      continue
    }
    if (argument.startsWith('--source-name=')) {
      sourceName = argument.slice('--source-name='.length)
      if (!sourceName.trim()) throw new Error('--source-name must not be blank.')
      continue
    }
    throw new Error(`Unknown option: ${argument}`)
  }

  return { slide, sourceName }
}

function parseSlide(value: string | undefined) {
  const slide = Number(value)
  if (!Number.isInteger(slide) || slide < 1) {
    throw new Error('--slide must be followed by a positive whole slide number.')
  }
  return slide
}

async function main() {
  const options = parseCliOptions(process.argv.slice(2))
  const source = await readStdin(25 * 1024 * 1024)
  const result = await importPowerPointBytes(source, options)
  writeStdout(Buffer.from(`${JSON.stringify(result.jsonSpec, null, 2)}\n`))

  for (const warning of result.warnings) {
    console.error(`WARNING ${warning}`)
  }
}

main().catch((error: ThrownValue) => {
  const message = error instanceof Error ? error.message : String(error)
  console.error(message)
  process.exitCode = 1
})
