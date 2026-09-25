import type { PowerPointCanvasJson } from '../lib/import/PowerpointImportTypes'
import type { SlideRetrievalMetadata } from '../lib/retrieval/SlideRetrievalMetadata'

export type StoredApp = {
  appId: string
  createdAt: string
  displayName: string
  lastSeenAt: string
  metadata: Record<string, string>
}

export type AppRegistration = {
  displayName?: string
  metadata?: Record<string, string>
}

export type StoredTemplate = {
  templateId: string
  templateJson: PowerPointCanvasJson
}

export type TemplateKind = 'commentary' | 'diagram'
export type TemplateSource = 'builtin' | 'import'

export type StoredTemplateMetadata = {
  checksum: string | null
  createdAt: string
  description: string
  kind: TemplateKind
  source: TemplateSource
  templateId: string
}

export type StoredTemplatePreview = {
  bytes: Buffer
  contentType: 'image/png'
  height: number
  templateId: string
  width: number
}

export type StoredTemplatePreviewPage = {
  previews: StoredTemplatePreview[]
  total: number
}

export type StoredTemplateSummary = StoredTemplate & {
  metadata: StoredTemplateMetadata
  previewAvailable: boolean
}

export type TemplateAsset = {
  assetId: string
  bytes: Buffer
  contentType: string
  templateId: string
}

export type StoredTemplateWithAssets = StoredTemplate & {
  assets: TemplateAsset[]
}

export type TemplateInsert = {
  assets: readonly TemplateAsset[]
  metadata?: StoredTemplateMetadata
  preview?: StoredTemplatePreview
  template: StoredTemplate
}

export type PendingSlideClassification = {
  inputFingerprint: string
  promptVersion: string
  schemaVersion: number
}

export type SlideClassificationStatus = 'pending' | 'processing' | 'ready' | 'failed'
export type SlideEmbeddingStatus = 'not_ready' | 'pending' | 'processing' | 'ready' | 'failed'

export type StoredSlideClassification = {
  attemptCount: number
  classifiedAt: string | null
  classifiedFingerprint: string | null
  embeddingAttemptCount: number
  embeddingDimensions: number | null
  capabilityEmbeddingDocument: string | null
  capabilityEmbeddingFingerprint: string | null
  capabilityVector: readonly number[] | null
  embeddingLastErrorCode: string | null
  embeddingModel: string | null
  embeddingNextAttemptAt: string | null
  embeddingStatus: SlideEmbeddingStatus
  inputFingerprint: string
  lastErrorCode: string | null
  metadata: SlideRetrievalMetadata | null
  model: string | null
  nextAttemptAt: string | null
  promptVersion: string
  schemaVersion: number
  status: SlideClassificationStatus
  templateId: string
  subjectEmbeddingDocument: string | null
  subjectEmbeddingFingerprint: string | null
  subjectVector: readonly number[] | null
}

export type SlideClassificationJob = {
  attemptCount: number
  inputFingerprint: string
  kind: TemplateKind
  preview?: StoredTemplatePreview
  template: StoredTemplate
  title: string
}

export type SlideEmbeddingJob = {
  attemptCount: number
  capabilityDocument: string
  capabilityFingerprint: string
  subjectDocument: string
  subjectFingerprint: string
  templateId: string
}

export type RetrievalClassificationRecord = {
  classification: StoredSlideClassification
  createdAt: string
  kind: TemplateKind
  previewAvailable: boolean
  templateId: string
  title: string
}

export type CompleteSlideClassification = {
  capabilityEmbeddingDocument: string
  capabilityEmbeddingFingerprint: string
  classifiedAt: string
  classifiedFingerprint: string
  metadata: SlideRetrievalMetadata
  model: string
  subjectEmbeddingDocument: string
  subjectEmbeddingFingerprint: string
  templateId: string
}

export type CompleteSlideProcessing = CompleteSlideClassification & {
  capabilityVector: readonly number[]
  embeddingDimensions: number
  embeddingModel: string
  subjectVector: readonly number[]
}

export interface TemplateRepository {
  /** Lease the next eligible classification row and load its app-neutral stored slide input. */
  claimNextSlideClassification(now: string, leaseExpiresAt: string): SlideClassificationJob | undefined
  /** Lease the next eligible embedding row without re-running classification. */
  claimNextSlideEmbedding(now: string, leaseExpiresAt: string): SlideEmbeddingJob | undefined
  /** Publish validated metadata, replace FTS content, and schedule embedding atomically. */
  completeSlideClassification(result: CompleteSlideClassification): void
  /** Store one validated vector when the claimed embedding fingerprint is still current. */
  completeSlideEmbedding(
    templateId: string,
    subjectFingerprint: string,
    capabilityFingerprint: string,
    subjectVector: readonly number[],
    capabilityVector: readonly number[],
    model: string,
    dimensions: number,
  ): void
  /** Atomically publish classification metadata, FTS content, and its validated embedding. */
  completeSlideProcessing(result: CompleteSlideProcessing): void
  delete(templateId: string, appId?: string): boolean
  ensureApp(appId: string, registration?: AppRegistration): StoredApp
  findApp(appId: string): StoredApp | undefined
  findAsset(templateId: string, assetId: string, appId?: string): TemplateAsset | undefined
  findById(templateId: string, appId?: string): StoredTemplate | undefined
  findByIdWithAssets(templateId: string, appId?: string): StoredTemplateWithAssets | undefined
  findPreview(templateId: string, appId?: string): StoredTemplatePreview | undefined
  findSlideClassification(templateId: string): StoredSlideClassification | undefined
  hasTemplate(templateId: string, appId?: string): boolean
  insert(
    template: StoredTemplate,
    assets: readonly TemplateAsset[],
    metadata?: StoredTemplateMetadata,
    preview?: StoredTemplatePreview,
    appId?: string,
  ): void
  insertMany(records: readonly TemplateInsert[], appId?: string): void
  /** Insert a template and its initial pending classification row in one transaction. */
  insertWithPendingClassification(
    record: TemplateInsert,
    pending: PendingSlideClassification,
    appId?: string,
  ): void
  /** Insert a template with a classification pipeline already owned by its import request. */
  insertWithProcessingClassification(
    record: TemplateInsert,
    pending: PendingSlideClassification,
    appId?: string,
  ): void
  list(kind?: TemplateKind, appId?: string): StoredTemplateSummary[]
  listPreviews(limit: number, offset: number, appId?: string): StoredTemplatePreviewPage
  /** Return only metadata reachable through the supplied app-template access scope. */
  listRetrievalClassifications(appId: string): RetrievalClassificationRecord[]
  /** Record a sanitized classification failure as retryable pending work or terminal failure. */
  recordSlideClassificationFailure(
    templateId: string,
    errorCode: string,
    retryAt: string | null,
  ): void
  /** Record a sanitized embedding failure independently from usable classification metadata. */
  recordSlideEmbeddingFailure(
    templateId: string,
    errorCode: string,
    retryAt: string | null,
  ): void
  /** Schedule classification refresh only when its complete input fingerprint changed. */
  refreshPendingSlideClassification(
    templateId: string,
    pending: PendingSlideClassification,
  ): boolean
  /** Schedule re-embedding without invalidating still-usable lexical metadata. */
  refreshPendingSlideEmbedding(
    templateId: string,
    subjectDocument: string,
    subjectFingerprint: string,
    capabilityDocument: string,
    capabilityFingerprint: string,
  ): boolean
  /** Execute a pre-sanitized FTS5 expression within an app's authorized template set. */
  searchSlideClassifications(appId: string, ftsQuery: string, limit: number): string[]
  upsertBuiltin(
    template: StoredTemplate,
    assets: readonly TemplateAsset[],
    metadata: StoredTemplateMetadata,
    preview: StoredTemplatePreview,
  ): void
  update(template: StoredTemplate, assets: readonly TemplateAsset[], appId?: string): void
}
