import { z } from 'zod'

export const SLIDE_CLASSIFICATION_SCHEMA_VERSION = 2
export const SLIDE_CLASSIFICATION_PROMPT_VERSION = '2026-09-14.2'

export const SLIDE_DOMAIN_TAXONOMY = [
  'cybersecurity',
  'software-architecture',
  'data-ai',
  'infrastructure',
  'cloud',
  'product-engineering',
  'it-operations',
  'business-applications',
  'governance-risk-compliance',
  'other',
] as const

export const COMMUNICATION_INTENT_TAXONOMY = [
  'finding',
  'evidence',
  'risk',
  'recommendation',
  'current-state',
  'target-state',
  'comparison',
  'process',
  'timeline',
  'metric',
  'explanation',
] as const

export const SLIDE_LAYOUT_TAXONOMY = [
  'title-only',
  'title-body',
  'single-column',
  'two-column',
  'three-column',
  'grid',
  'flow',
  'timeline',
  'chart-led',
  'table-led',
  'diagram-led',
  'mixed',
] as const

export const CONTENT_SLOT_CAPACITY_TAXONOMY = [
  'short-text',
  'paragraph',
  'three-to-five-bullets',
  'number',
  'table',
  'diagram-label',
  'image',
  'chart',
  'flexible',
] as const

const requiredText = (maximum: number) => z.string().trim().min(1).max(maximum)
const textList = (maximumItems: number, maximumItemLength: number) => z
  .array(requiredText(maximumItemLength))
  .max(maximumItems)

const DomainSchema = z.object({
  id: z.enum(SLIDE_DOMAIN_TAXONOMY),
  relevance: z.enum(['primary', 'secondary']),
  topics: textList(16, 120),
}).strict()

const ContentSlotSchema = z.object({
  element_id: requiredText(200),
  role: requiredText(100),
  capacity: z.enum(CONTENT_SLOT_CAPACITY_TAXONOMY),
}).strict()

export const SlideRetrievalMetadataSchema = z.object({
  subject: z.object({
    summary: requiredText(1_500),
    domains: z.array(DomainSchema).min(1).max(8),
    other_topics: textList(16, 120),
    technologies: textList(20, 120),
    entities: textList(20, 160),
    claims: textList(16, 300),
    synonyms: textList(24, 120),
  }).strict(),
  communication: z.object({
    intents: z.array(z.enum(COMMUNICATION_INTENT_TAXONOMY)).min(1).max(8),
    information_types: textList(16, 120),
    audience: textList(12, 120),
  }).strict(),
  template_fit: z.object({
    archetype: requiredText(120),
    content_slots: z.array(ContentSlotSchema).max(40),
  }).strict(),
  visual: z.object({
    layout_type: z.enum(SLIDE_LAYOUT_TAXONOMY),
    visual_elements: textList(20, 120),
    content_density: z.enum(['low', 'medium', 'high']),
    structural_features: textList(20, 160),
  }).strict(),
  retrieval_keywords: textList(30, 120),
}).strict().superRefine((value, context) => {
  const primaryCount = value.subject.domains.filter((domain) => domain.relevance === 'primary').length
  if (primaryCount !== 1) {
    context.addIssue({
      code: 'custom',
      message: 'Exactly one subject domain must be primary.',
      path: ['subject', 'domains'],
    })
  }
  const domainIds = value.subject.domains.map((domain) => domain.id)
  if (new Set(domainIds).size !== domainIds.length) {
    context.addIssue({ code: 'custom', message: 'Subject domains must be unique.', path: ['subject', 'domains'] })
  }
  const slotIds = value.template_fit.content_slots.map((slot) => slot.element_id)
  if (new Set(slotIds).size !== slotIds.length) {
    context.addIssue({
      code: 'custom',
      message: 'Content slots must reference unique element IDs.',
      path: ['template_fit', 'content_slots'],
    })
  }
})

export type SlideRetrievalMetadata = z.infer<typeof SlideRetrievalMetadataSchema>
export type SlideDomainId = SlideRetrievalMetadata['subject']['domains'][number]['id']
export type CommunicationIntent = SlideRetrievalMetadata['communication']['intents'][number]

/** Parse untrusted v2 classifier output and apply the canonical normalization policy. */
export function normalizeSlideRetrievalMetadata(value: unknown): SlideRetrievalMetadata {
  const parsed = SlideRetrievalMetadataSchema.parse(value)
  return SlideRetrievalMetadataSchema.parse({
    subject: {
      ...parsed.subject,
      summary: normalizeWhitespace(parsed.subject.summary),
      domains: parsed.subject.domains.map((domain) => ({
        ...domain,
        topics: deduplicate(domain.topics.map(normalizeTaxonomyValue)),
      })),
      other_topics: deduplicate(parsed.subject.other_topics.map(normalizeTaxonomyValue)),
      technologies: deduplicate(parsed.subject.technologies.map(normalizeWhitespace)),
      entities: deduplicate(parsed.subject.entities.map(normalizeWhitespace)),
      claims: deduplicate(parsed.subject.claims.map(normalizeWhitespace)),
      synonyms: deduplicate(parsed.subject.synonyms.map(normalizeWhitespace)),
    },
    communication: {
      ...parsed.communication,
      intents: deduplicate(parsed.communication.intents),
      information_types: deduplicate(parsed.communication.information_types.map(normalizeTaxonomyValue)),
      audience: deduplicate(parsed.communication.audience.map(normalizeWhitespace)),
    },
    template_fit: {
      archetype: normalizeTaxonomyValue(parsed.template_fit.archetype),
      content_slots: parsed.template_fit.content_slots.map((slot) => ({
        ...slot,
        element_id: normalizeWhitespace(slot.element_id),
        role: normalizeTaxonomyValue(slot.role),
      })),
    },
    visual: {
      ...parsed.visual,
      visual_elements: deduplicate(parsed.visual.visual_elements.map(normalizeTaxonomyValue)),
      structural_features: deduplicate(parsed.visual.structural_features.map(normalizeWhitespace)),
    },
    retrieval_keywords: deduplicate(parsed.retrieval_keywords.map(normalizeWhitespace)),
  })
}

/** Ensure model-proposed injection slots actually exist in the stored slide. */
export function assertContentSlotsExist(
  metadata: SlideRetrievalMetadata,
  elementIds: ReadonlySet<string>,
) {
  for (const slot of metadata.template_fit.content_slots) {
    if (!elementIds.has(slot.element_id)) {
      throw new Error('Classification content slots must reference stored slide elements.')
    }
  }
}

/** Convert an evolving classifier taxonomy label to stable lower-case kebab case. */
export function normalizeTaxonomyValue(value: string): string {
  const normalized = normalizeWhitespace(value)
    .toLocaleLowerCase('en-US')
    .replace(/[^a-z0-9]+/gu, '-')
    .replace(/^-+|-+$/gu, '')
  return normalized || 'other'
}

function normalizeWhitespace(value: string) {
  return value.trim().replace(/\s+/gu, ' ')
}

function deduplicate<Value extends string>(values: readonly Value[]): Value[] {
  const seen = new Set<string>()
  return values.filter((value) => {
    const key = value.toLocaleLowerCase('en-US')
    if (seen.has(key)) return false
    seen.add(key)
    return true
  })
}
