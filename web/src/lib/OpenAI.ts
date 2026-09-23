import OpenAIClient from 'openai'
import type {
  ResponseCreateParamsNonStreaming,
  ResponseInput,
  ResponseInputContent,
  ResponseInputFile,
  ResponseInputImage,
  ResponseInputText,
} from 'openai/resources/responses/responses'

import slideTextOnlyInstructions from '../prompts/SlideTextOnlyPrompt.md?raw'
import { SLIDE_PROMPT_OUTPUT_FORMAT } from '../types/SlidePromptOutput'
import type { SlidePromptOutput } from '../types/SlidePromptOutput'
import { getExtension } from '../utils/files'
import {
  logAttachmentDebug,
  summarizeAttachments,
} from './attachmentDiagnostics'
import type { JsonValue } from './canvas-model/CanvasTypes'

type CreateOpenAIResponseParams = {
  attachments?: File[]
  model?: string
  operation?: 'model-selection' | 'diagram-generation' | 'template-text-update' | 'unspecified'
  prompt: string
  systemInstructions?: string
  text?: ResponseCreateParamsNonStreaming['text']
}

type GenerateSlidePromptOutputParams = {
  attachments?: File[]
  prompt?: string
  templateJson: JsonValue
}

const DEFAULT_MODEL = import.meta.env.VITE_OPENAI_MODEL || 'gpt-5.2'

const MIME_BY_EXTENSION: Record<string, string> = {
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  jpeg: 'image/jpeg',
  jpg: 'image/jpeg',
  markdown: 'text/markdown',
  md: 'text/markdown',
  pdf: 'application/pdf',
  png: 'image/png',
  ppt: 'application/vnd.ms-powerpoint',
  pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  rtf: 'application/rtf',
  txt: 'text/plain',
}

let client: OpenAIClient | null = null

function getApiKey() {
  const apiKey =
    import.meta.env.VITE_OPENAI_API_KEY?.trim() ||
    import.meta.env.VITE_OPENAI_SECRET_KEY?.trim()

  if (!apiKey) {
    throw new Error(
      'Missing `VITE_OPENAI_API_KEY`. Add it to your Vite environment before generating a diagram.',
    )
  }

  return apiKey
}

function getClient() {
  client ??= new OpenAIClient({
    apiKey: getApiKey(),
    dangerouslyAllowBrowser: true,
  })

  return client
}

function arrayBufferToBase64(buffer: ArrayBuffer) {
  const bytes = new Uint8Array(buffer)
  const chunkSize = 0x8000
  let binary = ''

  for (let index = 0; index < bytes.length; index += chunkSize) {
    const chunk = bytes.subarray(index, index + chunkSize)
    binary += String.fromCharCode(...chunk)
  }

  return btoa(binary)
}

async function fileToBase64(file: File) {
  return arrayBufferToBase64(await file.arrayBuffer())
}

function isImageFile(file: File) {
  if (file.type.startsWith('image/')) {
    return true
  }

  const extension = getExtension(file.name)
  return extension === 'png' || extension === 'jpg' || extension === 'jpeg'
}

function getMimeType(file: File) {
  if (file.type.trim()) {
    return file.type
  }

  const extension = getExtension(file.name)
  return MIME_BY_EXTENSION[extension] ?? 'application/octet-stream'
}

async function buildImageAttachmentContent(file: File): Promise<ResponseInputImage> {
  const base64 = await fileToBase64(file)
  const mimeType = getMimeType(file)
  const dataUrl = `data:${mimeType};base64,${base64}`

  return {
    type: 'input_image',
    image_url: dataUrl,
    detail: 'auto',
  }
}

async function buildFileAttachmentContent(file: File): Promise<ResponseInputFile> {
  const base64 = await fileToBase64(file)
  const mimeType = getMimeType(file)
  const dataUrl = `data:${mimeType};base64,${base64}`

  return {
    type: 'input_file',
    filename: file.name,
    file_data: dataUrl,
  }
}

export async function buildOpenAIUserContent(
  prompt: string,
  attachments: File[] = [],
): Promise<ResponseInputContent[]> {
  const imageAttachments = attachments.filter(isImageFile)
  const fileAttachments = attachments.filter((file) => !isImageFile(file))
  const textContent: ResponseInputText = {
    type: 'input_text',
    text: prompt.trim(),
  }

  return [
    textContent,
    ...(await Promise.all(fileAttachments.map(buildFileAttachmentContent))),
    ...(await Promise.all(imageAttachments.map(buildImageAttachmentContent))),
  ]
}

export async function createOpenAIResponse({
  attachments = [],
  model = DEFAULT_MODEL,
  operation = 'unspecified',
  prompt,
  systemInstructions,
  text,
}: CreateOpenAIResponseParams) {
  const input: ResponseInput = []

  logAttachmentDebug('openai-request-started', {
    operation,
    attachments: summarizeAttachments(attachments),
  })

  if (systemInstructions?.trim()) {
    input.push({
      role: 'system',
      content: systemInstructions.trim(),
    })
  }

  let userContent: ResponseInputContent[]

  try {
    userContent = await buildOpenAIUserContent(prompt, attachments)
  } catch (error) {
    logAttachmentDebug('openai-attachment-build-failed', {
      attachmentCount: attachments.length,
      errorType: error instanceof Error ? error.name : 'unknown',
    })
    throw error
  }

  const attachmentContent = userContent.filter((content) => content.type !== 'input_text')
  const serializedAttachmentPayloadChars = attachmentContent.reduce((total, content) => {
    if (content.type === 'input_file') {
      return total + (typeof content.file_data === 'string' ? content.file_data.length : 0)
    }

    if (content.type === 'input_image') {
      return total + (typeof content.image_url === 'string' ? content.image_url.length : 0)
    }

    return total
  }, 0)
  const emptySerializedAttachmentCount = attachmentContent.filter((content) => {
    if (content.type === 'input_file') {
      return typeof content.file_data !== 'string' || content.file_data.length === 0
    }

    return content.type === 'input_image'
      && (typeof content.image_url !== 'string' || content.image_url.length === 0)
  }).length

  logAttachmentDebug('openai-request-attachments-built', {
    expectedAttachmentCount: attachments.length,
    contentCount: userContent.length,
    attachmentContentCount: attachmentContent.length,
    contentTypes: userContent.map((content) => content.type),
    serializedAttachmentPayloadChars,
    emptySerializedAttachmentCount,
  })

  input.push({
    role: 'user',
    content: userContent,
  })

  const response = await getClient().responses.create({
    model,
    input,
    ...(text ? { text } : {}),
  })

  logAttachmentDebug('openai-response-received', {
    operation,
    outputTextLength: response.output_text?.length ?? 0,
  })

  return response
}

function buildSlideTextOnlyPrompt(templateJson: JsonValue, prompt?: string) {
  const userContext = prompt?.trim()

  return [
    userContext
      ? `Use this user request as additional guidance:\n${userContext}`
      : 'Use the attached technical information to update the slide text values.',
    'Return the full resulting JSON object after editing only allowed text values.',
    'Base PowerPoint architecture diagram JSON:',
    JSON.stringify(templateJson, null, 2),
  ].join('\n\n')
}

export async function generateSlidePromptOutput({
  attachments = [],
  prompt,
  templateJson,
}: GenerateSlidePromptOutputParams): Promise<SlidePromptOutput> {
  const response = await createOpenAIResponse({
    attachments,
    operation: 'template-text-update',
    prompt: buildSlideTextOnlyPrompt(templateJson, prompt),
    systemInstructions: slideTextOnlyInstructions,
    text: {
      format: {
        type: 'json_schema',
        ...SLIDE_PROMPT_OUTPUT_FORMAT,
      },
    },
  })

  if (response.output_text) {
    return JSON.parse(response.output_text) as SlidePromptOutput
  }

  throw new Error('OpenAI did not return a structured slide JSON payload.')
}
