import { importPowerPointBytes } from '../lib/import/PowerpointImporter'
import type { PowerPointCanvasJson } from '../lib/import/PowerpointImportTypes'
import { ApiError } from '../errors'

const ZIP_CENTRAL_DIRECTORY_ENTRY_SIGNATURE = 0x02014b50
const ZIP_END_OF_CENTRAL_DIRECTORY_SIGNATURE = 0x06054b50
const ZIP_CENTRAL_DIRECTORY_FIXED_BYTES = 46
const ZIP_END_OF_CENTRAL_DIRECTORY_FIXED_BYTES = 22
const ZIP_MAX_COMMENT_BYTES = 65_535
const ZIP_MAX_ENTRIES = 10_000
const ZIP_MAX_ENTRY_UNCOMPRESSED_BYTES = 50 * 1024 * 1024
const ZIP_MAX_TOTAL_UNCOMPRESSED_BYTES = 250 * 1024 * 1024
const ZIP_DISK_NUMBER_OFFSET = 4
const ZIP_CENTRAL_DIRECTORY_DISK_OFFSET = 6
const ZIP_ENTRY_COUNT_OFFSET = 10
const ZIP_DIRECTORY_SIZE_OFFSET = 12
const ZIP_DIRECTORY_OFFSET = 16
const ZIP_ENTRY_UNCOMPRESSED_BYTES_OFFSET = 24
const ZIP_ENTRY_NAME_LENGTH_OFFSET = 28
const ZIP_ENTRY_EXTRA_LENGTH_OFFSET = 30
const ZIP_ENTRY_COMMENT_LENGTH_OFFSET = 32

export type PowerPointConversion = {
  templateJson: PowerPointCanvasJson
  warnings: string[]
}

export interface PowerPointConverter {
  convert(source: Buffer): Promise<PowerPointConversion>
}

export class LibraryPowerPointConverter implements PowerPointConverter {
  async convert(source: Buffer): Promise<PowerPointConversion> {
    validatePowerPointPackage(source)
    const result = await runImporter(source)
    return { templateJson: result.jsonSpec, warnings: result.warnings }
  }
}

async function runImporter(source: Buffer) {
  try {
    return await importPowerPointBytes(source)
  } catch (error) {
    throw new ApiError(
      422,
      'invalid_powerpoint',
      'The uploaded file is not a supported PowerPoint OOXML presentation.',
      { cause: error },
    )
  }
}

export function validatePowerPointPackage(source: Buffer) {
  const endOfDirectory = findEndOfCentralDirectory(source)
  if (endOfDirectory === undefined) {
    throw invalidPowerPointError()
  }

  const diskNumber = source.readUInt16LE(endOfDirectory + ZIP_DISK_NUMBER_OFFSET)
  const centralDirectoryDisk = source.readUInt16LE(endOfDirectory + ZIP_CENTRAL_DIRECTORY_DISK_OFFSET)
  const entryCount = source.readUInt16LE(endOfDirectory + ZIP_ENTRY_COUNT_OFFSET)
  const directorySize = source.readUInt32LE(endOfDirectory + ZIP_DIRECTORY_SIZE_OFFSET)
  const directoryOffset = source.readUInt32LE(endOfDirectory + ZIP_DIRECTORY_OFFSET)
  if (
    diskNumber !== 0
    || centralDirectoryDisk !== 0
    || entryCount === 0xffff
    || directorySize === 0xffffffff
    || directoryOffset === 0xffffffff
    || entryCount > ZIP_MAX_ENTRIES
    || directoryOffset + directorySize > endOfDirectory
  ) {
    throw invalidPowerPointError()
  }

  const requiredParts = new Set([
    '[Content_Types].xml',
    'ppt/presentation.xml',
    'ppt/_rels/presentation.xml.rels',
  ])
  let offset = directoryOffset
  let uncompressedBytes = 0

  for (let index = 0; index < entryCount; index += 1) {
    if (
      offset + ZIP_CENTRAL_DIRECTORY_FIXED_BYTES > source.length ||
      source.readUInt32LE(offset) !== ZIP_CENTRAL_DIRECTORY_ENTRY_SIGNATURE
    ) {
      throw invalidPowerPointError()
    }

    const entryBytes = source.readUInt32LE(offset + ZIP_ENTRY_UNCOMPRESSED_BYTES_OFFSET)
    const nameLength = source.readUInt16LE(offset + ZIP_ENTRY_NAME_LENGTH_OFFSET)
    const extraLength = source.readUInt16LE(offset + ZIP_ENTRY_EXTRA_LENGTH_OFFSET)
    const commentLength = source.readUInt16LE(offset + ZIP_ENTRY_COMMENT_LENGTH_OFFSET)
    const nextOffset =
      offset + ZIP_CENTRAL_DIRECTORY_FIXED_BYTES + nameLength + extraLength + commentLength
    if (entryBytes === 0xffffffff || nextOffset > source.length) {
      throw invalidPowerPointError()
    }

    uncompressedBytes += entryBytes
    if (
      entryBytes > ZIP_MAX_ENTRY_UNCOMPRESSED_BYTES ||
      uncompressedBytes > ZIP_MAX_TOTAL_UNCOMPRESSED_BYTES
    ) {
      throw new ApiError(
        413,
        'powerpoint_too_large',
        'The expanded PowerPoint package exceeds the processing limit.',
      )
    }

    const entryName = source.toString(
      'utf8',
      offset + ZIP_CENTRAL_DIRECTORY_FIXED_BYTES,
      offset + ZIP_CENTRAL_DIRECTORY_FIXED_BYTES + nameLength,
    )
    if (
      entryName.startsWith('/')
      || entryName.includes('\0')
      || entryName.includes('\\')
      || entryName.split('/').includes('..')
    ) {
      throw invalidPowerPointError()
    }
    requiredParts.delete(entryName)
    offset = nextOffset
  }

  if (offset !== directoryOffset + directorySize || requiredParts.size > 0) {
    throw invalidPowerPointError()
  }
}

function findEndOfCentralDirectory(source: Buffer) {
  const minimumOffset = Math.max(0, source.length - (ZIP_MAX_COMMENT_BYTES + ZIP_END_OF_CENTRAL_DIRECTORY_FIXED_BYTES))
  for (
    let offset = source.length - ZIP_END_OF_CENTRAL_DIRECTORY_FIXED_BYTES;
    offset >= minimumOffset;
    offset -= 1
  ) {
    if (source.readUInt32LE(offset) === ZIP_END_OF_CENTRAL_DIRECTORY_SIGNATURE) {
      return offset
    }
  }
  return undefined
}

function invalidPowerPointError() {
  return new ApiError(
    422,
    'invalid_powerpoint',
    'The uploaded file is not a supported PowerPoint OOXML presentation.',
  )
}
