import { z } from 'zod'

import { API_V1_PATH } from '../apiPaths'
import { buildAppScopedPath } from '../appIdentity'
import { SlideProviderError, type SlideEmbedder } from '../integrations/SlideProvider'
import { cosineSimilarity, validateEmbedding } from '../lib/retrieval/SlideVector'
import type { SlideRetrievalMetadata } from '../lib/retrieval/SlideRetrievalMetadata'
import type {
  RetrievalClassificationRecord,
  TemplateKind,
  TemplateRepository,
} from '../repositories/TemplateRepository'

const FACET_NAMES = [
  'slide_type',
  'business_domains',
  'technologies',
  'audience',
  'layout_type',
  'content_density',
  'information_types',
  'structural_features',
  'has_timeline',
  'has_table',
  'has_chart',
  'has_process_flow',
  'has_kpis',
  'has_recommendations',
] as const

const SELECT_NAMES = [
  'all',
  'identity',
  'content',
  'use_cases',
  'visual',
  'information',
  'structure',
  'keywords',
  'capabilities',
] as const

export const SlideQueryInputSchema = z.object({
  mode: z.enum(['text', 'semantic', 'hybrid', 'filter', 'similar', 'by_id', 'facets']),
  query: z.string().max(500).nullable(),
  template_ids: z.array(z.string().trim().min(1).max(200)).max(20),
  similar_to_template_id: z.string().trim().min(1).max(200).nullable(),
  facet_name: z.enum(FACET_NAMES).nullable(),
  filters: z.object({
    kinds: z.array(z.enum(['diagram', 'commentary'])).max(2),
    slide_types: z.array(z.string().trim().min(1).max(80)).max(10),
    business_domains: z.array(z.string().trim().min(1).max(120)).max(10),
    technologies: z.array(z.string().trim().min(1).max(120)).max(10),
    content_density: z.array(z.enum(['low', 'medium', 'high'])).max(3),
    has_timeline: z.boolean().nullable(),
    has_table: z.boolean().nullable(),
    has_chart: z.boolean().nullable(),
    has_process_flow: z.boolean().nullable(),
    has_kpis: z.boolean().nullable(),
    has_recommendations: z.boolean().nullable(),
  }).strict(),
  select: z.array(z.enum(SELECT_NAMES)).min(1).max(9),
  limit: z.number().int().min(1).max(20),
}).strict().superRefine((value, context) => {
  const textMode = value.mode === 'text' || value.mode === 'semantic' || value.mode === 'hybrid'
  if (textMode !== (value.query !== null && value.query.trim().length > 0)) {
    context.addIssue({ code: 'custom', message: 'query is required only for text-bearing modes.', path: ['query'] })
  }
  if ((value.mode === 'by_id') !== (value.template_ids.length > 0)) {
    context.addIssue({ code: 'custom', message: 'template_ids is required only for by_id.', path: ['template_ids'] })
  }
  if ((value.mode === 'similar') !== (value.similar_to_template_id !== null)) {
    context.addIssue({ code: 'custom', message: 'similar_to_template_id is required only for similar.', path: ['similar_to_template_id'] })
  }
  if ((value.mode === 'facets') !== (value.facet_name !== null)) {
    context.addIssue({ code: 'custom', message: 'facet_name is required only for facets.', path: ['facet_name'] })
  }
  if (value.select.includes('all') && value.select.length !== 1) {
    context.addIssue({ code: 'custom', message: 'all cannot be combined with another projection.', path: ['select'] })
  }
})

export type SlideQueryInput = z.infer<typeof SlideQueryInputSchema>
type SlideFacetName = NonNullable<SlideQueryInput['facet_name']>

export type SlideMetadataSections = {
  identity?: Pick<SlideRetrievalMetadata, 'slide_type' | 'slide_purpose' | 'description'>
  content?: Pick<SlideRetrievalMetadata, 'topics' | 'business_domains' | 'technologies' | 'entities'>
  use_cases?: Pick<SlideRetrievalMetadata, 'use_cases' | 'audience'>
  visual?: Pick<SlideRetrievalMetadata, 'layout_type' | 'visual_elements' | 'content_density'>
  information?: Pick<SlideRetrievalMetadata, 'information_types'>
  structure?: Pick<SlideRetrievalMetadata, 'structural_features'>
  keywords?: Pick<SlideRetrievalMetadata, 'retrieval_keywords'>
  capabilities?: Pick<SlideRetrievalMetadata,
    | 'has_timeline' | 'has_table' | 'has_chart' | 'has_process_flow'
    | 'has_kpis' | 'has_recommendations'>
}

export type SlideQueryResult = {
  kind: TemplateKind
  previewUrl: string | null
  sections: SlideMetadataSections
  templateId: string
  title: string
}

export type SlideRetrievalResponse = {
  mode_used: Exclude<SlideQueryInput['mode'], 'facets'>
  results: SlideQueryResult[]
  warnings: string[]
} | {
  facet: SlideFacetName
  mode_used: 'facets'
  values: Array<{ count: number; value: string | boolean }>
}

export class SlideRetrievalError extends Error {
  constructor(readonly code: 'invalid_query' | 'semantic_unavailable' | 'slide_not_available') {
    super('The slide retrieval request could not be completed.')
    this.name = 'SlideRetrievalError'
  }
}

export type SlideRetrievalServiceOptions = {
  embedder?: SlideEmbedder
  embeddingDimensions?: number
  embeddingModel?: string
  repository: TemplateRepository
}

export class SlideRetrievalService {
  constructor(private readonly options: SlideRetrievalServiceOptions) {}

  /** Validate and execute one app-scoped retrieval mode without exposing SQL or vectors. */
  async query(appId: string, untrustedInput: unknown, signal?: AbortSignal): Promise<SlideRetrievalResponse> {
    const parsed = SlideQueryInputSchema.safeParse(untrustedInput)
    if (!parsed.success) throw new SlideRetrievalError('invalid_query')
    const input = parsed.data
    const appRecords = this.options.repository.listRetrievalClassifications(appId)
    const records = applyFilters(appRecords, input.filters)

    if (input.mode === 'facets') {
      if (input.facet_name === null) throw new SlideRetrievalError('invalid_query')
      return buildFacetResponse(records, input.facet_name)
    }

    const warnings: string[] = []
    let ranked: RetrievalClassificationRecord[]
    switch (input.mode) {
      case 'by_id': {
        const byId = new Map(records.map((record) => [record.templateId, record]))
        ranked = input.template_ids.flatMap((templateId) => {
          const record = byId.get(templateId)
          return record ? [record] : []
        })
        break
      }
      case 'filter':
        ranked = records
        break
      case 'text':
        ranked = this.#lexical(appId, input.query, records)
        break
      case 'semantic':
        ranked = await this.#semantic(input.query, records, signal)
        break
      case 'hybrid': {
        const lexical = this.#lexical(appId, input.query, records)
        try {
          const semantic = await this.#semantic(input.query, records, signal)
          if (semantic.length === 0 && records.length > 0) {
            warnings.push('semantic_unavailable')
            ranked = lexical
          } else {
            ranked = reciprocalRankFusion(lexical, semantic)
          }
        } catch (error) {
          if (!(error instanceof SlideRetrievalError) || error.code !== 'semantic_unavailable') throw error
          warnings.push('semantic_unavailable')
          ranked = lexical
        }
        break
      }
      case 'similar': {
        const source = appRecords.find((record) => record.templateId === input.similar_to_template_id)
        if (!source?.classification.vector || source.classification.embeddingStatus !== 'ready') {
          throw new SlideRetrievalError('slide_not_available')
        }
        ranked = rankByVector(source.classification.vector, records.filter((record) => (
          record.templateId !== source.templateId && compatibleVector(record, source.classification.embeddingModel, source.classification.embeddingDimensions)
        )))
        break
      }
    }

    return {
      mode_used: input.mode,
      results: ranked.slice(0, input.limit).map((record) => projectRecord(record, input.select, appId)),
      warnings,
    }
  }

  #lexical(appId: string, query: string | null, records: readonly RetrievalClassificationRecord[]) {
    if (query === null) throw new SlideRetrievalError('invalid_query')
    const expression = buildSafeFtsQuery(query)
    const available = new Map(records.map((record) => [record.templateId, record]))
    return this.options.repository.searchSlideClassifications(appId, expression, 200).flatMap((id) => {
      const record = available.get(id)
      return record ? [record] : []
    })
  }

  async #semantic(
    query: string | null,
    records: readonly RetrievalClassificationRecord[],
    signal?: AbortSignal,
  ) {
    if (query === null || !this.options.embedder || !this.options.embeddingModel) {
      throw new SlideRetrievalError('semantic_unavailable')
    }
    const compatible = records.filter((record) => compatibleVector(
      record,
      this.options.embeddingModel ?? null,
      this.options.embeddingDimensions ?? record.classification.embeddingDimensions,
    ))
    if (compatible.length === 0) return []
    let vector: readonly number[]
    try {
      vector = await this.options.embedder.embed(query.trim(), signal)
      validateEmbedding(vector, this.options.embeddingDimensions)
    } catch (error) {
      if (error instanceof SlideProviderError || error instanceof Error) {
        throw new SlideRetrievalError('semantic_unavailable')
      }
      throw error
    }
    return rankByVector(vector, compatible.filter((record) => (
      record.classification.embeddingDimensions === vector.length
    )))
  }
}

function buildSafeFtsQuery(query: string) {
  if (/["*():^{}\[\]]/u.test(query)) throw new SlideRetrievalError('invalid_query')
  const tokens = query.toLocaleLowerCase('en-US').match(/[\p{L}\p{N}]+/gu) ?? []
  if (tokens.length === 0 || tokens.length > 50) throw new SlideRetrievalError('invalid_query')
  return [...new Set(tokens)].map((token) => `"${token}"*`).join(' AND ')
}

function compatibleVector(
  record: RetrievalClassificationRecord,
  model: string | null,
  dimensions: number | null,
) {
  const classification = record.classification
  return classification.embeddingStatus === 'ready'
    && classification.vector !== null
    && classification.embeddingModel === model
    && classification.embeddingDimensions === dimensions
}

function rankByVector(vector: readonly number[], records: readonly RetrievalClassificationRecord[]) {
  return records.map((record) => ({
    record,
    score: cosineSimilarity(vector, record.classification.vector ?? []),
  })).sort((left, right) => (
    right.score - left.score
    || right.record.createdAt.localeCompare(left.record.createdAt)
    || left.record.templateId.localeCompare(right.record.templateId)
  )).map(({ record }) => record)
}

/** Fuse lexical and semantic ordinal ranks without combining incomparable raw scores. */
export function reciprocalRankFusion(
  lexical: readonly RetrievalClassificationRecord[],
  semantic: readonly RetrievalClassificationRecord[],
  rankConstant = 60,
) {
  const records = new Map<string, { record: RetrievalClassificationRecord; score: number }>()
  for (const [index, record] of lexical.entries()) {
    records.set(record.templateId, { record, score: 1 / (rankConstant + index + 1) })
  }
  for (const [index, record] of semantic.entries()) {
    const current = records.get(record.templateId)
    if (current) current.score += 1 / (rankConstant + index + 1)
    else records.set(record.templateId, { record, score: 1 / (rankConstant + index + 1) })
  }
  return [...records.values()].sort((left, right) => (
    right.score - left.score
    || right.record.createdAt.localeCompare(left.record.createdAt)
    || left.record.templateId.localeCompare(right.record.templateId)
  )).map(({ record }) => record)
}

function applyFilters(
  records: readonly RetrievalClassificationRecord[],
  filters: SlideQueryInput['filters'],
) {
  const includes = (values: readonly string[], requested: readonly string[]) => requested.length === 0
    || requested.some((value) => values.some((candidate) => equal(candidate, value)))
  return records.filter(({ classification, kind }) => {
    const metadata = classification.metadata
    if (!metadata) return false
    return (filters.kinds.length === 0 || filters.kinds.includes(kind))
      && includes([metadata.slide_type], filters.slide_types)
      && includes(metadata.business_domains, filters.business_domains)
      && includes(metadata.technologies, filters.technologies)
      && (filters.content_density.length === 0 || filters.content_density.includes(metadata.content_density))
      && matchesBoolean(metadata.has_timeline, filters.has_timeline)
      && matchesBoolean(metadata.has_table, filters.has_table)
      && matchesBoolean(metadata.has_chart, filters.has_chart)
      && matchesBoolean(metadata.has_process_flow, filters.has_process_flow)
      && matchesBoolean(metadata.has_kpis, filters.has_kpis)
      && matchesBoolean(metadata.has_recommendations, filters.has_recommendations)
  })
}

function matchesBoolean(value: boolean, requested: boolean | null) {
  return requested === null || value === requested
}

function equal(left: string, right: string) {
  return left.toLocaleLowerCase('en-US') === right.toLocaleLowerCase('en-US')
}

function projectRecord(
  record: RetrievalClassificationRecord,
  select: SlideQueryInput['select'],
  appId: string,
): SlideQueryResult {
  const metadata = record.classification.metadata
  if (!metadata) throw new Error('Retrieval metadata is missing.')
  const requested = new Set(select.includes('all') ? SELECT_NAMES.filter((name) => name !== 'all') : select)
  const sections: SlideMetadataSections = {}
  if (requested.has('identity')) sections.identity = pick(metadata, ['slide_type', 'slide_purpose', 'description'])
  if (requested.has('content')) sections.content = pick(metadata, ['topics', 'business_domains', 'technologies', 'entities'])
  if (requested.has('use_cases')) sections.use_cases = pick(metadata, ['use_cases', 'audience'])
  if (requested.has('visual')) sections.visual = pick(metadata, ['layout_type', 'visual_elements', 'content_density'])
  if (requested.has('information')) sections.information = pick(metadata, ['information_types'])
  if (requested.has('structure')) sections.structure = pick(metadata, ['structural_features'])
  if (requested.has('keywords')) sections.keywords = pick(metadata, ['retrieval_keywords'])
  if (requested.has('capabilities')) sections.capabilities = pick(metadata, [
    'has_timeline', 'has_table', 'has_chart', 'has_process_flow', 'has_kpis', 'has_recommendations',
  ])
  return {
    kind: record.kind,
    previewUrl: record.previewAvailable
      ? buildAppScopedPath(`${API_V1_PATH}/templates/${record.templateId}/preview`, appId)
      : null,
    sections,
    templateId: record.templateId,
    title: record.title,
  }
}

function pick<ObjectType extends object, Key extends keyof ObjectType>(
  object: ObjectType,
  keys: readonly Key[],
): Pick<ObjectType, Key> {
  return Object.fromEntries(keys.map((key) => [key, object[key]])) as Pick<ObjectType, Key>
}

function buildFacetResponse(
  records: readonly RetrievalClassificationRecord[],
  facet: SlideFacetName,
): Extract<SlideRetrievalResponse, { mode_used: 'facets' }> {
  const counts = new Map<string | boolean, number>()
  for (const { classification } of records) {
    const metadata = classification.metadata
    if (!metadata) continue
    const raw = metadata[facet]
    const values = Array.isArray(raw) ? raw : [raw]
    for (const value of values) counts.set(value, (counts.get(value) ?? 0) + 1)
  }
  return {
    facet,
    mode_used: 'facets',
    values: [...counts.entries()]
      .map(([value, count]) => ({ count, value }))
      .sort((left, right) => right.count - left.count || String(left.value).localeCompare(String(right.value))),
  }
}
