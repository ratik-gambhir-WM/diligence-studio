// @vitest-environment node

import { readFile, readdir } from 'node:fs/promises'

import { describe, expect, it } from 'vitest'

const SERVER_ROOT = new URL('../', import.meta.url)
const FILESYSTEM_IMPORT = /import\s*\{([^}]*)\}\s*from\s*['"]node:fs(?:\/promises)?['"]/gu

describe('server in-memory I/O boundary', () => {
  it('contains no filesystem-writing API in production or CLI TypeScript', async () => {
    const files = [
      ...await listTypeScriptFiles(new URL('src/', SERVER_ROOT)),
      ...await listTypeScriptFiles(new URL('scripts/', SERVER_ROOT), '.test.ts'),
    ]

    for (const file of files) {
      const source = await readFile(file, 'utf8')
      if (file.pathname.endsWith('/src/catalog/seedBuiltinTemplates.ts')) {
        const imports = [...source.matchAll(FILESYSTEM_IMPORT)]
        expect(imports.map((match) => match[1]?.trim()), file.pathname).toEqual(['readFile'])
      } else {
        expect(source, file.pathname).not.toMatch(/['"]node:fs(?:\/promises)?['"]/u)
      }
      expect(source, file.pathname).not.toMatch(/\.writeFile\s*\(/u)
    }
  })

  it('creates SQLite storage in memory and captures Chromium screenshots as bytes', async () => {
    const repositorySource = await readFile(
      new URL('src/repositories/SqliteTemplateRepository.ts', SERVER_ROOT),
      'utf8',
    )
    expect(repositorySource).toContain("new DatabaseSync(':memory:')")
    expect(repositorySource).toContain('PRAGMA temp_store = MEMORY')

    const previewSource = await readFile(
      new URL('src/services/TemplatePreview.ts', SERVER_ROOT),
      'utf8',
    )
    expect(previewSource).toContain("type: 'png'")
    expect(previewSource).not.toMatch(/screenshot\(\{[^}]*\bpath\s*:/su)
  })
})

async function listTypeScriptFiles(directory: URL, excludedSuffix?: string): Promise<URL[]> {
  const entries = await readdir(directory, { withFileTypes: true })
  const files: URL[] = []

  for (const entry of entries) {
    const url = new URL(entry.name, directory)
    if (entry.isDirectory()) {
      files.push(...await listTypeScriptFiles(new URL(`${entry.name}/`, directory), excludedSuffix))
    } else if (
      entry.name.endsWith('.ts')
      && (!excludedSuffix || !entry.name.endsWith(excludedSuffix))
    ) {
      files.push(url)
    }
  }

  return files
}
