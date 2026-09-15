import { z } from 'zod'

import { API_V1_PATH } from '../apiPaths'
import { buildAppScopedPath } from '../appIdentity'
import { SlideProviderError, type SlideEmbedder } from '../integrations/SlideProvider'
import {
  COMMUNICATION_INTENT_TAXONOMY,
  SLIDE_DOMAIN_TAXONOMY,
  normalizeTaxonomyValue,
  type CommunicationIntent,
  type SlideDomainId,
  type SlideRetrievalMetadata,
} from '../lib/retrieval/SlideRetrievalMetadata'
import { cosineSimilarity, validateEmbedding } from '../lib/retrieval/SlideVector'
import type {
  RetrievalClassificationRecord,
  TemplateKind,
  TemplateRepository,
} from '../repositories/TemplateRepository'

const FACET_NAMES = [
  'domains',
  'topics',
  'intents',
  'archetype',
  'layout_type',
  'content_density',
  'slot_roles',
] as const

const SELECT_NAMES = ['all', 'subject', 'communication', 'template_fit', 'visual', 'keywords'] as const

export const SlideQueryInputSchema = z.object({
  mode: z.enum(['text', 'semantic', 'hybrid', 'filter', 'similar', 'by_id', 'facets']),
  query: z.string().max(2_000).nullable(),
  semantic_target: z.enum(['subject', 'capability', 'both']),
  template_ids: z.array(z.string().trim().min(1).max(200)).max(20),
  similar_to_template_id: z.string().trim().min(1).max(200).nullable(),
  facet_name: z.enum(FACET_NAMES).nullable(),
  filters: z.object({
    kinds: z.array(z.enum(['diagram', 'commentary'])).max(2),
    domains: z.array(z.enum(SLIDE_DOMAIN_TAXONOMY)).max(10),
    topics: z.array(z.string().trim().min(1).max(120)).max(20),
    intents: z.array(z.enum(COMMUNICATION_INTENT_TAXONOMY)).max(10),
    archetypes: z.array(z.string().trim().min(1).max(120)).max(10),
    slot_roles: z.array(z.string().trim().min(1).max(100)).max(20),
    content_density: z.array(z.enum(['low', 'medium', 'high'])).max(3),
  }).strict(),
  select: z.array(z.enum(SELECT_NAMES)).min(1).max(6),
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

export const FindSlidesForFindingInputSchema = z.object({
  markdown: z.string().trim().min(1).max(20_000),
  limit: z.number().int().min(1).max(20).default(5),
}).strict()

export type SlideQueryInput = z.infer<typeof SlideQueryInputSchema>
export type FindSlidesForFindingInput = z.infer<typeof FindSlidesForFindingInputSchema>
type SlideFacetName = NonNullable<SlideQueryInput['facet_name']>

export type SlideMetadataSections = {
  subject?: SlideRetrievalMetadata['subject']
  communication?: SlideRetrievalMetadata['communication']
  template_fit?: SlideRetrievalMetadata['template_fit']
  visual?: SlideRetrievalMetadata['visual']
  keywords?: Pick<SlideRetrievalMetadata, 'retrieval_keywords'>
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
  values: Array<{ count: number; value: string }>
}

export type FindingSlideMatch = {
  kind: TemplateKind
  matchedDomains: SlideDomainId[]
  matchedIntents: CommunicationIntent[]
  matchedTopics: string[]
  matchReasons: string[]
  previewUrl: string | null
  templateId: string
  title: string
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
        ranked = await this.#semantic(input.query, records, input.semantic_target, signal)
        break
      case 'hybrid': {
        const lexical = this.#lexical(appId, input.query, records)
        try {
          const semantic = await this.#semantic(input.query, records, input.semantic_target, signal)
          ranked = semantic.length === 0 && records.length > 0
            ? lexical
            : reciprocalRankFusion(lexical, semantic)
          if (semantic.length === 0 && records.length > 0) warnings.push('semantic_unavailable')
        } catch (error) {
          if (!(error instanceof SlideRetrievalError) || error.code !== 'semantic_unavailable') throw error
          warnings.push('semantic_unavailable')
          ranked = lexical
        }
        break
      }
      case 'similar': {
        const source = appRecords.find((record) => record.templateId === input.similar_to_template_id)
        if (!source || !hasReadyVectors(source)) throw new SlideRetrievalError('slide_not_available')
        const candidates = records.filter((record) => (
          record.templateId !== source.templateId && compatibleVectors(
            record,
            source.classification.embeddingModel,
            source.classification.embeddingDimensions,
          )
        ))
        ranked = reciprocalRankFusion(
          rankByVector(source.classification.subjectVector, candidates, 'subject'),
          rankByVector(source.classification.capabilityVector, candidates, 'capability'),
        )
        break
      }
    }

    return {
      mode_used: input.mode,
      results: ranked.slice(0, input.limit).map((record) => projectRecord(record, input.select, appId)),
      warnings,
    }
  }

  /** Retrieve slides for a complete finding using subject and layout fitness independently. */
  async findForFinding(
    appId: string,
    untrustedInput: unknown,
    signal?: AbortSignal,
  ): Promise<{ results: FindingSlideMatch[]; warnings: string[] }> {
    const parsed = FindSlidesForFindingInputSchema.safeParse(untrustedInput)
    if (!parsed.success) throw new SlideRetrievalError('invalid_query')
    const records = this.options.repository.listRetrievalClassifications(appId)
    if (!this.options.embedder || !this.options.embeddingModel) {
      throw new SlideRetrievalError('semantic_unavailable')
    }
    const analysis = analyzeFinding(parsed.data.markdown, records)
    const compatible = records.filter((record) => compatibleVectors(
      record,
      this.options.embeddingModel ?? null,
      this.options.embeddingDimensions ?? record.classification.embeddingDimensions,
    ))
    if (compatible.length === 0) return { results: [], warnings: [] }

    const capabilityQuery = [
      `communication intents: ${analysis.intents.join(' | ') || 'finding | evidence | recommendation'}`,
      `required content slots: ${analysis.requiredSlotRoles.join(' | ') || 'headline | finding | evidence'}`,
    ].join('\n')
    const [subjectVector, capabilityVector] = await Promise.all([
      this.#embed(parsed.data.markdown, signal),
      this.#embed(capabilityQuery, signal),
    ])
    const lexicalTerms = [...analysis.domains, ...analysis.topics, ...analysis.intents]
    const lexical = lexicalTerms.length > 0
      ? this.#lexical(appId, lexicalTerms.join(' '), compatible, 'OR')
      : []
    const fused = reciprocalRankScores([
      lexical,
      rankByVector(subjectVector, compatible, 'subject'),
      rankByVector(capabilityVector, compatible, 'capability'),
    ])
    const ranked = fused.map(({ record, score }) => ({
      combinedScore: score + exactFindingBoost(record, analysis) * 0.005,
      record,
    })).sort((left, right) => (
      right.combinedScore - left.combinedScore
      || right.record.createdAt.localeCompare(left.record.createdAt)
      || left.record.templateId.localeCompare(right.record.templateId)
    ))

    return {
      results: ranked.slice(0, parsed.data.limit).map(({ record }) => explainFindingMatch(record, analysis, appId)),
      warnings: [],
    }
  }

  #lexical(
    appId: string,
    query: string | null,
    records: readonly RetrievalClassificationRecord[],
    operator: 'AND' | 'OR' = 'AND',
  ) {
    if (query === null) throw new SlideRetrievalError('invalid_query')
    const expression = buildSafeFtsQuery(query, operator)
    const available = new Map(records.map((record) => [record.templateId, record]))
    return this.options.repository.searchSlideClassifications(appId, expression, 200).flatMap((id) => {
      const record = available.get(id)
      return record ? [record] : []
    })
  }

  async #semantic(
    query: string | null,
    records: readonly RetrievalClassificationRecord[],
    target: SlideQueryInput['semantic_target'],
    signal?: AbortSignal,
  ) {
    if (query === null) throw new SlideRetrievalError('invalid_query')
    if (!this.options.embedder || !this.options.embeddingModel) {
      throw new SlideRetrievalError('semantic_unavailable')
    }
    const compatible = records.filter((record) => compatibleVectors(
      record,
      this.options.embeddingModel ?? null,
      this.options.embeddingDimensions ?? record.classification.embeddingDimensions,
    ))
    if (compatible.length === 0) return []
    const vector = await this.#embed(query.trim(), signal)
    if (target === 'subject') return rankByVector(vector, compatible, 'subject')
    if (target === 'capability') return rankByVector(vector, compatible, 'capability')
    return reciprocalRankFusion(
      rankByVector(vector, compatible, 'subject'),
      rankByVector(vector, compatible, 'capability'),
    )
  }

  async #embed(document: string, signal?: AbortSignal) {
    if (!this.options.embedder || !this.options.embeddingModel) {
      throw new SlideRetrievalError('semantic_unavailable')
    }
    try {
      const vector = await this.options.embedder.embed(document, signal)
      validateEmbedding(vector, this.options.embeddingDimensions)
      return vector
    } catch (error) {
      if (error instanceof SlideProviderError || error instanceof Error) {
        throw new SlideRetrievalError('semantic_unavailable')
      }
      throw error
    }
  }
}

function buildSafeFtsQuery(query: string, operator: 'AND' | 'OR') {
  if (/["*():^{}\[\]]/u.test(query)) throw new SlideRetrievalError('invalid_query')
  const tokens = query.toLocaleLowerCase('en-US').match(/[\p{L}\p{N}]+/gu) ?? []
  if (tokens.length === 0 || tokens.length > 200) throw new SlideRetrievalError('invalid_query')
  return [...new Set(tokens)].map((token) => `"${token}"*`).join(` ${operator} `)
}

function hasReadyVectors(record: RetrievalClassificationRecord): record is RetrievalClassificationRecord & {
  classification: RetrievalClassificationRecord['classification'] & {
    capabilityVector: readonly number[]
    subjectVector: readonly number[]
  }
} {
  return record.classification.embeddingStatus === 'ready'
    && record.classification.subjectVector !== null
    && record.classification.capabilityVector !== null
}

function compatibleVectors(
  record: RetrievalClassificationRecord,
  model: string | null,
  dimensions: number | null,
) {
  return hasReadyVectors(record)
    && record.classification.embeddingModel === model
    && record.classification.embeddingDimensions === dimensions
}

function rankByVector(
  vector: readonly number[],
  records: readonly RetrievalClassificationRecord[],
  role: 'subject' | 'capability',
) {
  return records.map((record) => ({
    record,
    score: cosineSimilarity(
      vector,
      role === 'subject'
        ? record.classification.subjectVector ?? []
        : record.classification.capabilityVector ?? [],
    ),
  })).sort((left, right) => (
    right.score - left.score
    || right.record.createdAt.localeCompare(left.record.createdAt)
    || left.record.templateId.localeCompare(right.record.templateId)
  )).map(({ record }) => record)
}

/** Fuse ordinal ranks without combining incomparable lexical and vector scores. */
export function reciprocalRankFusion(
  first: readonly RetrievalClassificationRecord[],
  second: readonly RetrievalClassificationRecord[],
  rankConstant = 60,
) {
  return reciprocalRankScores([first, second], rankConstant).map(({ record }) => record)
}

function reciprocalRankScores(
  rankings: ReadonlyArray<readonly RetrievalClassificationRecord[]>,
  rankConstant = 60,
) {
  const records = new Map<string, { record: RetrievalClassificationRecord; score: number }>()
  for (const ranking of rankings) {
    for (const [index, record] of ranking.entries()) {
      const current = records.get(record.templateId)
      if (current) current.score += 1 / (rankConstant + index + 1)
      else records.set(record.templateId, { record, score: 1 / (rankConstant + index + 1) })
    }
  }
  return [...records.values()].sort((left, right) => (
    right.score - left.score
    || right.record.createdAt.localeCompare(left.record.createdAt)
    || left.record.templateId.localeCompare(right.record.templateId)
  ))
}

function applyFilters(records: readonly RetrievalClassificationRecord[], filters: SlideQueryInput['filters']) {
  const includes = (values: readonly string[], requested: readonly string[]) => requested.length === 0
    || requested.some((value) => values.some((candidate) => equal(candidate, value)))
  return records.filter(({ classification, kind }) => {
    const metadata = classification.metadata
    if (!metadata) return false
    const domains = metadata.subject.domains.map((domain) => domain.id)
    const topics = [...metadata.subject.domains.flatMap((domain) => domain.topics), ...metadata.subject.other_topics]
    const slotRoles = metadata.template_fit.content_slots.map((slot) => slot.role)
    return (filters.kinds.length === 0 || filters.kinds.includes(kind))
      && includes(domains, filters.domains)
      && includes(topics, filters.topics)
      && includes(metadata.communication.intents, filters.intents)
      && includes([metadata.template_fit.archetype], filters.archetypes)
      && includes(slotRoles, filters.slot_roles)
      && (filters.content_density.length === 0 || filters.content_density.includes(metadata.visual.content_density))
  })
}

function equal(left: string, right: string) {
  return normalizeTaxonomyValue(left) === normalizeTaxonomyValue(right)
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
  if (requested.has('subject')) sections.subject = metadata.subject
  if (requested.has('communication')) sections.communication = metadata.communication
  if (requested.has('template_fit')) sections.template_fit = metadata.template_fit
  if (requested.has('visual')) sections.visual = metadata.visual
  if (requested.has('keywords')) sections.keywords = { retrieval_keywords: metadata.retrieval_keywords }
  return {
    kind: record.kind,
    previewUrl: previewUrl(record, appId),
    sections,
    templateId: record.templateId,
    title: record.title,
  }
}

function buildFacetResponse(
  records: readonly RetrievalClassificationRecord[],
  facet: SlideFacetName,
): Extract<SlideRetrievalResponse, { mode_used: 'facets' }> {
  const counts = new Map<string, number>()
  for (const { classification } of records) {
    const metadata = classification.metadata
    if (!metadata) continue
    const values = facetValues(metadata, facet)
    for (const value of new Set(values)) counts.set(value, (counts.get(value) ?? 0) + 1)
  }
  return {
    facet,
    mode_used: 'facets',
    values: [...counts.entries()]
      .map(([value, count]) => ({ count, value }))
      .sort((left, right) => right.count - left.count || left.value.localeCompare(right.value)),
  }
}

function facetValues(metadata: SlideRetrievalMetadata, facet: SlideFacetName): string[] {
  switch (facet) {
    case 'domains': return metadata.subject.domains.map((domain) => domain.id)
    case 'topics': return [...metadata.subject.domains.flatMap((domain) => domain.topics), ...metadata.subject.other_topics]
    case 'intents': return metadata.communication.intents
    case 'archetype': return [metadata.template_fit.archetype]
    case 'layout_type': return [metadata.visual.layout_type]
    case 'content_density': return [metadata.visual.content_density]
    case 'slot_roles': return metadata.template_fit.content_slots.map((slot) => slot.role)
  }
}

type FindingAnalysis = {
  domains: SlideDomainId[]
  intents: CommunicationIntent[]
  requiredSlotRoles: string[]
  topics: string[]
}

const DOMAIN_TERMS: Readonly<Record<SlideDomainId, readonly string[]>> = {
  cybersecurity: ['cyber', 'cybersecurity', 'security', 'vulnerability', 'iam', 'penetration test'],
  'software-architecture': ['software architecture', 'application architecture', 'extensibility', 'integration'],
  'data-ai': ['data', 'analytics', 'artificial intelligence', 'machine learning', 'ai'],
  infrastructure: ['infrastructure', 'network', 'server', 'compute', 'storage', 'datacenter'],
  cloud: ['cloud', 'aws', 'azure', 'gcp'],
  'product-engineering': ['product engineering', 'sdlc', 'developer', 'code quality'],
  'it-operations': ['it operations', 'service management', 'incident', 'monitoring'],
  'business-applications': ['erp', 'crm', 'business application'],
  'governance-risk-compliance': ['governance', 'compliance', 'regulatory', 'risk management'],
  other: [],
}

const INTENT_SLOT_ROLES: Readonly<Partial<Record<CommunicationIntent, readonly string[]>>> = {
  finding: ['finding'],
  evidence: ['evidence'],
  recommendation: ['recommendation'],
  comparison: ['comparison-left', 'comparison-right'],
  metric: ['metric'],
  process: ['process-step'],
  timeline: ['timeline-event'],
}

function analyzeFinding(markdown: string, records: readonly RetrievalClassificationRecord[]): FindingAnalysis {
  const searchable = normalizeSearchText(markdown)
  const domains = SLIDE_DOMAIN_TAXONOMY.filter((domain) => (
    domain !== 'other' && DOMAIN_TERMS[domain].some((term) => containsPhrase(searchable, term))
  ))
  const intents = COMMUNICATION_INTENT_TAXONOMY.filter((intent) => (
    containsPhrase(searchable, intent)
  ))
  const resolvedIntents: CommunicationIntent[] = intents.length > 0
    ? intents
    : ['finding', 'evidence', 'recommendation']
  const allTopics = new Set(records.flatMap(({ classification }) => {
    const metadata = classification.metadata
    return metadata
      ? [...metadata.subject.domains.flatMap((domain) => domain.topics), ...metadata.subject.other_topics]
      : []
  }))
  const topics = [...allTopics].filter((topic) => containsPhrase(searchable, topic))
  const requiredSlotRoles = [
    'headline',
    ...resolvedIntents.flatMap((intent) => INTENT_SLOT_ROLES[intent] ?? []),
  ]
  return {
    domains,
    intents: resolvedIntents,
    requiredSlotRoles: [...new Set(requiredSlotRoles)],
    topics,
  }
}

function exactFindingBoost(record: RetrievalClassificationRecord, analysis: FindingAnalysis) {
  const metadata = record.classification.metadata
  if (!metadata) return 0
  const domainScore = metadata.subject.domains.reduce((score, domain) => (
    analysis.domains.includes(domain.id) ? score + (domain.relevance === 'primary' ? 6 : 3) : score
  ), 0)
  const topics = [...metadata.subject.domains.flatMap((domain) => domain.topics), ...metadata.subject.other_topics]
  const topicScore = analysis.topics.filter((topic) => topics.some((candidate) => equal(candidate, topic))).length * 4
  const intentScore = analysis.intents.filter((intent) => metadata.communication.intents.includes(intent)).length * 2
  const slotRoles = metadata.template_fit.content_slots.map((slot) => slot.role)
  const slotScore = analysis.requiredSlotRoles.filter((role) => slotRoles.some((slot) => equal(slot, role))).length * 2
  return domainScore + topicScore + intentScore + slotScore
}

function explainFindingMatch(
  record: RetrievalClassificationRecord,
  analysis: FindingAnalysis,
  appId: string,
): FindingSlideMatch {
  const metadata = record.classification.metadata
  if (!metadata) throw new Error('Retrieval metadata is missing.')
  const matchedDomains = metadata.subject.domains
    .filter((domain) => analysis.domains.includes(domain.id))
    .map((domain) => domain.id)
  const slideTopics = [...metadata.subject.domains.flatMap((domain) => domain.topics), ...metadata.subject.other_topics]
  const matchedTopics = analysis.topics.filter((topic) => slideTopics.some((candidate) => equal(candidate, topic)))
  const matchedIntents = analysis.intents.filter((intent) => metadata.communication.intents.includes(intent))
  const slotRoles = metadata.template_fit.content_slots.map((slot) => slot.role)
  const matchedSlots = analysis.requiredSlotRoles.filter((role) => slotRoles.some((slot) => equal(slot, role)))
  const matchReasons = [
    ...matchedDomains.map((domain) => `${domain} subject match`),
    ...matchedTopics.map((topic) => `${topic} topic match`),
    ...(matchedIntents.length > 0 ? [`Supports ${matchedIntents.join(', ')} communication`] : []),
    ...(matchedSlots.length > 0 ? [`Provides ${matchedSlots.join(', ')} content slots`] : []),
  ]
  if (matchReasons.length === 0) matchReasons.push('Semantic subject and template-capability similarity')
  return {
    kind: record.kind,
    matchedDomains,
    matchedIntents,
    matchedTopics,
    matchReasons,
    previewUrl: previewUrl(record, appId),
    templateId: record.templateId,
    title: record.title,
  }
}

function previewUrl(record: RetrievalClassificationRecord, appId: string) {
  return record.previewAvailable
    ? buildAppScopedPath(`${API_V1_PATH}/templates/${record.templateId}/preview`, appId)
    : null
}

function normalizeSearchText(value: string) {
  return value.toLocaleLowerCase('en-US').replace(/[^a-z0-9]+/gu, ' ').replace(/\s+/gu, ' ').trim()
}

function containsPhrase(normalizedText: string, phrase: string) {
  const normalizedPhrase = normalizeSearchText(phrase)
  return normalizedPhrase.length > 0 && ` ${normalizedText} `.includes(` ${normalizedPhrase} `)
}
