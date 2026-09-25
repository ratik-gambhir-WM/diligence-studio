import { z } from 'zod'

import {
  FindSlidesForFindingInputSchema,
  SlideQueryInputSchema,
  type SlideRetrievalService,
} from './SlideRetrievalService'
import type { ImportService } from './ImportTemplateService'

export const GetSlideInputSchema = z.object({
  template_id: z.string().trim().min(1).max(200),
}).strict()

/**
 * Internal, app-bound functions for a future Agents SDK runtime. App identity is
 * captured by trusted server context and is never part of model-visible input.
 */
export function createSlideAgentTools(
  appId: string,
  retrieval: SlideRetrievalService,
  templates: ImportService,
) {
  return {
    get_slide: (untrustedInput: unknown) => {
      const { template_id } = GetSlideInputSchema.parse(untrustedInput)
      const slide = templates.find(template_id, appId)
      if (!slide) return { error: { code: 'slide_not_available' } }
      return slide
    },
    find_slides_for_finding: (untrustedInput: unknown, signal?: AbortSignal) => {
      const input = FindSlidesForFindingInputSchema.parse(untrustedInput)
      return retrieval.findForFinding(appId, input, signal)
    },
    query_slides: (untrustedInput: unknown, signal?: AbortSignal) => {
      const input = SlideQueryInputSchema.parse(untrustedInput)
      return retrieval.query(appId, input, signal)
    },
  }
}
