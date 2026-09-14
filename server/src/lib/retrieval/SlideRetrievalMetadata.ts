import { z } from 'zod'

export const SLIDE_CLASSIFICATION_SCHEMA_VERSION = 1
export const SLIDE_CLASSIFICATION_PROMPT_VERSION = '2026-09-13.1'

export const SLIDE_TYPE_TAXONOMY = [
  'title',
  'agenda',
  'executive-summary',
  'architecture',
  'process',
  'timeline',
  'comparison',
  'kpi-dashboard',
  'financial',
  'market-overview',
  'recommendation',
  'risk',
  'organization',
  'table',
  'other',
] as const

export const SLIDE_LAYOUT_TAXONOMY = [
  'title-only',
  'title-body',
  'single-column',
  'two-column',
  'grid',
  'flow',
  'timeline',
  'chart-led',
  'table-led',
  'diagram-led',
  'mixed',
] as const

const requiredText = (maximum: number) => z.string().trim().min(1).max(maximum)
const textList = (maximumItems: number, maximumItemLength: number) => z
  .array(requiredText(maximumItemLength))
  .max(maximumItems)

export const SlideRetrievalMetadataSchema = z.object({
  slide_type: requiredText(80),
  slide_purpose: requiredText(500),
  description: requiredText(1_500),
  topics: textList(20, 120),
  business_domains: textList(12, 120),
  technologies: textList(20, 120),
  entities: textList(20, 160),
  use_cases: textList(16, 240),
  audience: textList(12, 120),
  layout_type: requiredText(80),
  visual_elements: textList(20, 120),
  content_density: z.enum(['low', 'medium', 'high']),
  information_types: textList(16, 120),
  structural_features: textList(20, 160),
  retrieval_keywords: textList(30, 120),
  has_timeline: z.boolean(),
  has_table: z.boolean(),
  has_chart: z.boolean(),
  has_process_flow: z.boolean(),
  has_kpis: z.boolean(),
  has_recommendations: z.boolean(),
}).strict()

export type SlideRetrievalMetadata = z.infer<typeof SlideRetrievalMetadataSchema>

const ARRAY_FIELDS = [
  'topics',
  'business_domains',
  'technologies',
  'entities',
  'use_cases',
  'audience',
  'visual_elements',
  'information_types',
  'structural_features',
  'retrieval_keywords',
] as const satisfies readonly (keyof SlideRetrievalMetadata)[]

/** Parse untrusted classifier output and apply the one canonical normalization policy. */
export function normalizeSlideRetrievalMetadata(value: unknown): SlideRetrievalMetadata {
  const parsed = SlideRetrievalMetadataSchema.parse(value)
  const normalized: SlideRetrievalMetadata = {
    ...parsed,
    slide_type: normalizeTaxonomyValue(parsed.slide_type),
    slide_purpose: normalizeWhitespace(parsed.slide_purpose),
    description: normalizeWhitespace(parsed.description),
    layout_type: normalizeTaxonomyValue(parsed.layout_type),
  }

  for (const field of ARRAY_FIELDS) {
    normalized[field] = deduplicate(parsed[field].map(normalizeWhitespace))
  }
  return normalized
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

function deduplicate(values: readonly string[]) {
  const seen = new Set<string>()
  return values.filter((value) => {
    const key = value.toLocaleLowerCase('en-US')
    if (seen.has(key)) return false
    seen.add(key)
    return true
  })
}
