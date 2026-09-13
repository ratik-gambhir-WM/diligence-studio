import { TextDecoder } from 'node:util'
import path from 'node:path'

import Busboy from 'busboy'
import type { RequestHandler } from 'express'

import { parseAppId } from '../appIdentity'
import { ApiError } from '../errors'
import type { ExportPowerPointUseCase } from '../services/ExportPowerPointService'

const POWERPOINT_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.presentation'

export function createExportHandlers(service: ExportPowerPointUseCase) {
  const create: RequestHandler = async (request, response) => {
    if (!request.is('application/json')) {
      throw new ApiError(
        415,
        'unsupported_media_type',
        'Content-Type must be application/json.',
      )
    }

    const result = await service.export(parseJsonBody(request.body), parseAppId(request))
    if (response.writableEnded) {
      return
    }
    const bytes = Buffer.from(result.bytes)
    response
      .status(200)
      .set({
        'Cache-Control': 'no-store',
        'Content-Disposition': `attachment; filename="${result.fileName}"`,
        'Content-Length': String(bytes.length),
        'Content-Type': POWERPOINT_CONTENT_TYPE,
        'X-PowerPoint-Warning-Count': String(result.warnings.length),
      })
      .send(bytes)
  }

  const insert = (maxJsonBytes: number, maxTargetBytes: number): RequestHandler =>
    async (request, response) => {
      const parts = await parseInsertMultipart(request, maxJsonBytes, maxTargetBytes)
      const result = await service.insert(
        parseJsonBytes(parts.presentation),
        parts.target,
        parts.targetFileName,
        parts.insertAfterSlide,
        parseAppId(request),
      )
      if (response.writableEnded) {
        return
      }
      sendPowerPoint(response, result)
    }

  return { create, insert }
}

type InsertMultipartParts = {
  insertAfterSlide: number
  presentation: Buffer
  target: Buffer
  targetFileName: string
}

function parseInsertMultipart(
  request: Parameters<RequestHandler>[0],
  maxJsonBytes: number,
  maxTargetBytes: number,
) {
  return new Promise<InsertMultipartParts>((resolve, reject) => {
    let parser: Busboy.Busboy
    try {
      parser = Busboy({
        headers: request.headers,
        limits: {
          fieldNameSize: 64,
          fieldSize: Math.max(maxJsonBytes, 64),
          fields: 2,
          fileSize: maxTargetBytes,
          files: 1,
          parts: 4,
        },
      })
    } catch {
      reject(new ApiError(415, 'unsupported_media_type', 'Content-Type must be multipart/form-data.'))
      return
    }

    let presentation: Buffer | undefined
    let insertAfterSlide: number | undefined
    let target: Buffer | undefined
    let targetFileName = 'presentation.pptx'
    let failed = false
    const fail = (error: ApiError) => {
      if (!failed) {
        failed = true
        reject(error)
      }
    }

    parser.on('field', (name, value, info) => {
      if (info.valueTruncated) {
        fail(new ApiError(413, 'payload_too_large', 'A multipart field exceeds the configured size limit.'))
        return
      }
      if (name === 'presentation') {
        const bytes = Buffer.from(value, 'utf8')
        if (bytes.length > maxJsonBytes) {
          fail(new ApiError(413, 'payload_too_large', 'The presentation JSON exceeds the configured size limit.'))
          return
        }
        presentation = bytes
      } else if (name === 'insertAfterSlide') {
        if (!/^\d+$/u.test(value)) {
          fail(new ApiError(400, 'invalid_insert_position', 'insertAfterSlide must be a non-negative whole number.'))
          return
        }
        insertAfterSlide = Number(value)
      } else {
        fail(new ApiError(400, 'unexpected_multipart_field', 'The multipart request contains an unexpected field.'))
      }
    })
    parser.on('file', (name, stream, info) => {
      if (name !== 'target') {
        stream.resume()
        fail(new ApiError(400, 'unexpected_multipart_field', 'The multipart request contains an unexpected file.'))
        return
      }
      const normalizedName = path.basename(info.filename || 'presentation.pptx')
      if (
        !normalizedName.toLowerCase().endsWith('.pptx')
        || info.mimeType !== POWERPOINT_CONTENT_TYPE
      ) {
        stream.resume()
        fail(new ApiError(415, 'unsupported_target_media_type', 'The target must be a .pptx PowerPoint file.'))
        return
      }
      targetFileName = normalizedName
      const chunks: Buffer[] = []
      stream.on('data', (chunk: Buffer) => chunks.push(chunk))
      stream.on('limit', () => fail(new ApiError(413, 'payload_too_large', 'The target PowerPoint exceeds the configured size limit.')))
      stream.on('end', () => {
        target = Buffer.concat(chunks)
      })
    })
    parser.on('partsLimit', () => fail(new ApiError(413, 'payload_too_large', 'The multipart request contains too many parts.')))
    parser.on('fieldsLimit', () => fail(new ApiError(413, 'payload_too_large', 'The multipart request contains too many fields.')))
    parser.on('filesLimit', () => fail(new ApiError(413, 'payload_too_large', 'The multipart request contains too many files.')))
    parser.on('error', () => fail(new ApiError(400, 'invalid_multipart', 'The multipart request could not be parsed.')))
    parser.on('close', () => {
      if (failed) {
        return
      }
      if (!presentation || !target || target.length === 0 || insertAfterSlide === undefined) {
        fail(new ApiError(400, 'missing_multipart_field', 'presentation, target, and insertAfterSlide are required.'))
        return
      }
      resolve({ insertAfterSlide, presentation, target, targetFileName })
    })
    request.pipe(parser)
  })
}

function parseJsonBytes(body: Buffer) {
  return parseJsonBody(body)
}

function sendPowerPoint(
  response: Parameters<RequestHandler>[1],
  result: Awaited<ReturnType<ExportPowerPointUseCase['export']>>,
) {
  const bytes = Buffer.from(result.bytes)
  response.status(200).set({
    'Cache-Control': 'no-store',
    'Content-Disposition': `attachment; filename="${result.fileName}"`,
    'Content-Length': String(bytes.length),
    'Content-Type': POWERPOINT_CONTENT_TYPE,
    'X-PowerPoint-Warning-Count': String(result.warnings.length),
  }).send(bytes)
}

function parseJsonBody(body: unknown) {
  if (!Buffer.isBuffer(body) || body.length === 0) {
    throw invalidJsonError()
  }

  try {
    const json = new TextDecoder('utf-8', { fatal: true }).decode(body)
    return JSON.parse(json) as unknown
  } catch {
    throw invalidJsonError()
  }
}

function invalidJsonError() {
  return new ApiError(400, 'invalid_json', 'The request body must contain valid JSON.')
}
