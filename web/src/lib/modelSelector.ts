import modelSelectorInstructions from '../prompts/ModelSelectorPrompt.md?raw'
import type { PickerTemplateSummary } from './api/templateApi'
import { createOpenAIResponse } from './OpenAI'
import {
  MODEL_SELECTOR_OUTPUT_FORMAT,
  type ModelSelectorOutput,
} from '../types/ModelSelectorOutput'

type SelectArchitectureDiagramParams = {
  candidates: PickerTemplateSummary[]
  uploadedFiles: File[]
}

type DiagramCandidateForPrompt = {
  description: string
  id: string
  imageAttachmentLabel: string
  name: string
  relatedAlt: string
}

const CANDIDATE_IMAGE_WIDTH = 960
const CANDIDATE_IMAGE_HEADER_HEIGHT = 140
const CANDIDATE_IMAGE_PADDING = 28

function getCandidateAttachmentLabel(template: PickerTemplateSummary) {
  return `architecture-candidate__${template.templateId}.png`
}

function loadImageFromBlob(blob: Blob) {
  return new Promise<HTMLImageElement>((resolve, reject) => {
    const image = new Image()
    const objectUrl = URL.createObjectURL(blob)

    image.onload = () => {
      URL.revokeObjectURL(objectUrl)
      resolve(image)
    }
    image.onerror = () => {
      URL.revokeObjectURL(objectUrl)
      reject(new Error('Failed to load an architecture diagram candidate image.'))
    }
    image.src = objectUrl
  })
}

function canvasToPngBlob(canvas: HTMLCanvasElement) {
  return new Promise<Blob>((resolve, reject) => {
    canvas.toBlob((blob) => {
      if (blob) {
        resolve(blob)
      } else {
        reject(new Error('Failed to render an architecture diagram candidate image.'))
      }
    }, 'image/png')
  })
}

function drawTextLine({
  baseline,
  color,
  ctx,
  font,
  text,
  x,
}: {
  baseline: number
  color: string
  ctx: CanvasRenderingContext2D
  font: string
  text: string
  x: number
}) {
  ctx.fillStyle = color
  ctx.font = font
  ctx.fillText(text, x, baseline)
}

async function buildLabeledCandidateImage(template: PickerTemplateSummary) {
  if (!template.previewUrl) return undefined
  const response = await fetch(template.previewUrl)

  if (!response.ok) {
    throw new Error(`Failed to load architecture diagram candidate image: ${template.title}`)
  }

  const sourceBlob = await response.blob()
  const image = await loadImageFromBlob(sourceBlob)
  const imageScale = Math.min(1, CANDIDATE_IMAGE_WIDTH / image.naturalWidth)
  const imageWidth = Math.round(image.naturalWidth * imageScale)
  const imageHeight = Math.round(image.naturalHeight * imageScale)
  const canvas = document.createElement('canvas')
  canvas.width = imageWidth
  canvas.height = CANDIDATE_IMAGE_HEADER_HEIGHT + imageHeight

  const ctx = canvas.getContext('2d')

  if (!ctx) {
    throw new Error('Unable to render architecture diagram candidate labels.')
  }

  ctx.fillStyle = '#ffffff'
  ctx.fillRect(0, 0, canvas.width, canvas.height)
  ctx.fillStyle = '#171717'
  ctx.fillRect(0, 0, canvas.width, CANDIDATE_IMAGE_HEADER_HEIGHT)

  drawTextLine({
    baseline: 46,
    color: '#ffffff',
    ctx,
    font: '700 26px Arial, sans-serif',
    text: `ID: ${template.templateId}`,
    x: CANDIDATE_IMAGE_PADDING,
  })
  drawTextLine({
    baseline: 84,
    color: '#ffffff',
    ctx,
    font: '700 24px Arial, sans-serif',
    text: template.title,
    x: CANDIDATE_IMAGE_PADDING,
  })
  drawTextLine({
    baseline: 116,
    color: '#d8d8d8',
    ctx,
    font: '18px Arial, sans-serif',
    text: template.description.slice(0, 120),
    x: CANDIDATE_IMAGE_PADDING,
  })

  ctx.drawImage(image, 0, CANDIDATE_IMAGE_HEADER_HEIGHT, imageWidth, imageHeight)

  const labeledBlob = await canvasToPngBlob(canvas)
  const filename = getCandidateAttachmentLabel(template)

  return new File([labeledBlob], filename, { type: 'image/png' })
}

async function buildCandidateImageAttachments(candidates: PickerTemplateSummary[]) {
  const attachments = await Promise.all(candidates.map(buildLabeledCandidateImage))
  return attachments.filter((attachment): attachment is File => attachment !== undefined)
}

function buildCandidateListForPrompt(candidates: PickerTemplateSummary[]): DiagramCandidateForPrompt[] {
  return candidates.map((template) => ({
    id: template.templateId,
    name: template.title,
    description: template.description,
    relatedAlt: `${template.title} template preview`,
    imageAttachmentLabel: getCandidateAttachmentLabel(template),
  }))
}

function buildSelectorPrompt(uploadedFiles: File[], candidates: PickerTemplateSummary[]) {
  return [
    'Select the best architecture diagram candidate for the uploaded diligence source material.',
    'The uploaded diligence source files are:',
    JSON.stringify(
      uploadedFiles.map((file) => ({
        name: file.name,
        type: file.type || 'unknown',
        sizeBytes: file.size,
      })),
      null,
      2,
    ),
    'The architecture diagram candidate previews are attached as labeled images. Each candidate preview image has the candidate ID and name printed in a black header.',
    'Candidate catalog:',
    JSON.stringify(buildCandidateListForPrompt(candidates), null, 2),
    'Choose exactly one candidate ID from the candidate catalog.',
  ].join('\n\n')
}

function validateSelectedDiagram(output: ModelSelectorOutput, candidates: PickerTemplateSummary[]) {
  const selectedTemplate = candidates.find((template) => template.templateId === output.selectedDiagramId)

  if (!selectedTemplate) {
    throw new Error(`Model selector returned an unknown diagram id: ${output.selectedDiagramId}`)
  }

  return {
    ...output,
    selectedDiagramName: selectedTemplate.title,
  }
}

export async function selectArchitectureDiagramModel({
  candidates,
  uploadedFiles,
}: SelectArchitectureDiagramParams): Promise<ModelSelectorOutput> {
  if (candidates.length === 0) {
    throw new Error('No diagram templates are available for selection.')
  }
  const candidateImageAttachments = await buildCandidateImageAttachments(candidates)
  const response = await createOpenAIResponse({
    attachments: [...uploadedFiles, ...candidateImageAttachments],
    operation: 'model-selection',
    prompt: buildSelectorPrompt(uploadedFiles, candidates),
    systemInstructions: modelSelectorInstructions,
    text: {
      format: {
        type: 'json_schema',
        ...MODEL_SELECTOR_OUTPUT_FORMAT,
      },
    },
  })

  if (!response.output_text) {
    throw new Error('OpenAI did not return a model selector payload.')
  }

  const parsedOutput = JSON.parse(response.output_text) as ModelSelectorOutput

  return validateSelectedDiagram(parsedOutput, candidates)
}

export function getSelectedArchitectureTemplate(
  selection: ModelSelectorOutput | null,
  candidates: PickerTemplateSummary[],
) {
  return candidates.find((template) => template.templateId === selection?.selectedDiagramId)
}
