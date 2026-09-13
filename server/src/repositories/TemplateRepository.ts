import type { PowerPointCanvasJson } from '../lib/import/PowerpointImportTypes'

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

export interface TemplateRepository {
  delete(templateId: string, appId?: string): boolean
  ensureApp(appId: string, registration?: AppRegistration): StoredApp
  findApp(appId: string): StoredApp | undefined
  findAsset(templateId: string, assetId: string, appId?: string): TemplateAsset | undefined
  findById(templateId: string, appId?: string): StoredTemplate | undefined
  findByIdWithAssets(templateId: string, appId?: string): StoredTemplateWithAssets | undefined
  findPreview(templateId: string, appId?: string): StoredTemplatePreview | undefined
  insert(
    template: StoredTemplate,
    assets: readonly TemplateAsset[],
    metadata?: StoredTemplateMetadata,
    preview?: StoredTemplatePreview,
    appId?: string,
  ): void
  insertMany(records: readonly TemplateInsert[], appId?: string): void
  list(kind?: TemplateKind, appId?: string): StoredTemplateSummary[]
  listPreviews(limit: number, offset: number, appId?: string): StoredTemplatePreviewPage
  upsertBuiltin(
    template: StoredTemplate,
    assets: readonly TemplateAsset[],
    metadata: StoredTemplateMetadata,
    preview: StoredTemplatePreview,
  ): void
  update(template: StoredTemplate, assets: readonly TemplateAsset[], appId?: string): void
}
