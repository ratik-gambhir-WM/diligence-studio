import slideDiagramGenerationInstructions from '../prompts/SlideDiagramGenerationPrompt.md?raw'
import {
  GENERATED_SLIDE_PROMPT_OUTPUT_FORMAT,
  type GeneratedSlidePromptOutput,
} from '../types/SlidePromptOutput'
import type { CanvasTemplate } from './canvas-model/templates'
import { createOpenAIResponse } from './OpenAI'
import type { JsonValue } from './canvas-model/CanvasTypes'

type GenerateArchitectureDiagramParams = {
  candidates: CanvasTemplate[]
  uploadedFiles: File[]
}

type DiagramExampleForPrompt = {
  description: string
  id: string
  imageAttachmentLabel: string
  jsonSpec: JsonValue
  name: string
}

function getExampleAttachmentLabel(template: CanvasTemplate) {
  return `architecture-example__${template.id}.png`
}

async function buildExampleImageAttachment(template: CanvasTemplate) {
  const response = await fetch(template.image)

  if (!response.ok) {
    throw new Error(`Failed to load architecture diagram example image: ${template.name}`)
  }

  const blob = await response.blob()

  return new File([blob], getExampleAttachmentLabel(template), {
    type: blob.type || 'image/png',
  })
}

function buildExampleListForPrompt(candidates: CanvasTemplate[]): DiagramExampleForPrompt[] {
  return candidates.map((template) => ({
    id: template.id,
    name: template.name,
    description: template.description,
    imageAttachmentLabel: getExampleAttachmentLabel(template),
    jsonSpec: template.jsonSpec,
  }))
}

function buildGenerationPrompt(uploadedFiles: File[], candidates: CanvasTemplate[]) {
  return [
    'Create a new architecture diagram JSON object from the uploaded source material.',
    'This should be a platform-level M&A due diligence architecture and data-flow view.',
    'Prioritize applications and modules that directly develop, deliver, or support the product platform. Deprioritize back-office applications unless they are critical integrations or materially affect platform operations.',
    'Focus your generation effort on: selecting the right node text, grouping the whole platform, showing important data flow relationships through layout, and choosing clean x/y/w/h dimensions for every box.',
    'Do not redesign typography, colors, theme values, or styling. Preserve those values from the examples and only choose which existing example style pattern to reuse for each element.',
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
    'The attached architecture diagram example images correspond to the example JSON objects below. Match each image using its imageAttachmentLabel.',
    'Use the examples to learn the JSON schema, allowed element structures, existing style values, spacing patterns, and rendered output conventions.',
    'Use the example JSON primarily as a library of reusable style patterns. Keep fill, stroke, fontFace, fontSize, textColor, bold, align, valign, padding, and other theme/style values consistent with the examples.',
    'Your main layout task is to manipulate x, y, w, and h so the whole platform is readable and the important applications are shown as peer boxes where appropriate.',
    'Architecture diagram JSON examples:',
    JSON.stringify(buildExampleListForPrompt(candidates), null, 2),
    'Generate one complete new architecture diagram JSON for the uploaded source material. Return the JSON object directly.',
  ].join('\n\n')
}

export async function generateArchitectureDiagramFromExamples({
  candidates,
  uploadedFiles,
}: GenerateArchitectureDiagramParams): Promise<GeneratedSlidePromptOutput> {
  if (candidates.length === 0) {
    throw new Error('No diagram templates are available as generation examples.')
  }
  const exampleImageAttachments = await Promise.all(candidates.map(buildExampleImageAttachment))
  const response = await createOpenAIResponse({
    attachments: [...uploadedFiles, ...exampleImageAttachments],
    operation: 'diagram-generation',
    prompt: buildGenerationPrompt(uploadedFiles, candidates),
    systemInstructions: slideDiagramGenerationInstructions,
    text: {
      format: {
        type: 'json_schema',
        ...GENERATED_SLIDE_PROMPT_OUTPUT_FORMAT,
      },
    },
  })

  if (response.output_text) {
    return constrainSlideBounds(
      validateGeneratedDiagram(JSON.parse(response.output_text) as GeneratedSlidePromptOutput),
    )
  }

  throw new Error('OpenAI did not return a generated slide JSON payload.')
}

function validateGeneratedDiagram(output: GeneratedSlidePromptOutput): GeneratedSlidePromptOutput {
  const slides = output.presentation?.slides
  const firstSlide = slides?.[0]

  if (!firstSlide) {
    throw new Error('OpenAI returned JSON without a slide. Try Create again or use the template flow.')
  }

  if (!Array.isArray(firstSlide.elements) || firstSlide.elements.length === 0) {
    throw new Error('OpenAI returned an empty architecture diagram. Try Create again or use the template flow.')
  }

  if (
    slides.some((slide) =>
      slide.elements.some(
        (element) => (element as { type?: string }).type === 'line',
      ),
    )
  ) {
    throw new Error('OpenAI returned a connector line, which the Create flow does not accept. Try Create again.')
  }

  return output
}

function constrainSlideBounds(output: GeneratedSlidePromptOutput): GeneratedSlidePromptOutput {
  return {
    ...output,
    presentation: {
      ...output.presentation,
      slides: output.presentation.slides.map((slide) => {
        const width = Math.max(slide.width, 1)
        const height = Math.max(slide.height, 1)

        return {
          ...slide,
          elements: slide.elements.map((element) => {
            const x = clamp(element.x, 0, width)
            const y = clamp(element.y, 0, height)
            const w = clamp(element.w, 1, Math.max(width - x, 1))
            const h = clamp(element.h, 1, Math.max(height - y, 1))

            return {
              ...element,
              h,
              w,
              x,
              y,
            }
          }),
        }
      }),
    },
  }
}

function clamp(value: number, min: number, max: number) {
  if (!Number.isFinite(value)) {
    return min
  }

  return Math.min(Math.max(value, min), max)
}
