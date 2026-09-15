const EMPTY_STDIN_MESSAGE = 'Input bytes are required on stdin.'

export async function readStdin(maxBytes: number): Promise<Buffer> {
  const chunks: Buffer[] = []
  let totalBytes = 0

  for await (const chunk of process.stdin) {
    const bytes = Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk)
    totalBytes += bytes.length
    if (totalBytes > maxBytes) {
      throw new Error(`Stdin exceeds the ${maxBytes}-byte limit.`)
    }
    chunks.push(bytes)
  }

  if (totalBytes === 0) throw new Error(EMPTY_STDIN_MESSAGE)
  return Buffer.concat(chunks, totalBytes)
}

export function writeStdout(bytes: Buffer): void {
  process.stdout.write(bytes)
}
