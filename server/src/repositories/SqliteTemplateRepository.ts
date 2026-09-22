import { DatabaseSync } from 'node:sqlite'

import { DEFAULT_APP_ID } from '../appIdentity'
import type { PowerPointCanvasJson } from '../lib/import/PowerpointImportTypes'
import { deserializeEmbedding, serializeEmbedding } from '../lib/retrieval/SlideVector'
import { normalizeSlideRetrievalMetadata } from '../lib/retrieval/SlideRetrievalMetadata'
import type {
  AppRegistration,
  CompleteSlideClassification,
  CompleteSlideProcessing,
  PendingSlideClassification,
  RetrievalClassificationRecord,
  SlideClassificationJob,
  SlideClassificationStatus,
  SlideEmbeddingJob,
  SlideEmbeddingStatus,
  StoredApp,
  StoredSlideClassification,
  StoredTemplate,
  StoredTemplateMetadata,
  StoredTemplatePreview,
  StoredTemplatePreviewPage,
  StoredTemplateSummary,
  StoredTemplateWithAssets,
  TemplateAsset,
  TemplateInsert,
  TemplateKind,
  TemplateRepository,
} from './TemplateRepository'

type TemplateRow = {
  template_id: string
  template_json: string
}

type TemplateAssetRow = {
  asset_data: Uint8Array
  asset_id: string
  content_type: string
  template_id: string
}

type TemplateWithAssetRow = TemplateRow & {
  asset_data: Uint8Array | null
  asset_id: string | null
  content_type: string | null
  asset_template_id: string | null
}

type TemplatePreviewRow = {
  content_type: 'image/png'
  height: number
  preview_data: Uint8Array
  template_id: string
  width: number
}

type TemplateSummaryRow = TemplateRow & {
  checksum: string | null
  created_at: string
  description: string
  kind: TemplateKind
  preview_available: number
  source: 'builtin' | 'import'
}

type AppRow = {
  app_id: string
  created_at: string
  display_name: string
  last_seen_at: string
  metadata_json: string
}

type ClassificationRow = {
  attempt_count: number
  classified_at: string | null
  classified_fingerprint: string | null
  embedding_attempt_count: number
  embedding_dimensions: number | null
  embedding_last_error_code: string | null
  embedding_model: string | null
  embedding_next_attempt_at: string | null
  embedding_status: SlideEmbeddingStatus
  capability_embedding: Uint8Array | null
  capability_embedding_document: string | null
  capability_embedding_fingerprint: string | null
  input_fingerprint: string
  last_error_code: string | null
  metadata_json: string | null
  model: string | null
  next_attempt_at: string | null
  prompt_version: string
  schema_version: number
  subject_embedding: Uint8Array | null
  subject_embedding_document: string | null
  subject_embedding_fingerprint: string | null
  status: SlideClassificationStatus
  template_id: string
}

type RetrievalRow = ClassificationRow & {
  created_at: string
  kind: TemplateKind
  preview_available: number
  template_json: string
}

const SCHEMA_VERSION = 5

export class SqliteTemplateRepository implements TemplateRepository {
  readonly #database: DatabaseSync

  constructor(initialize?: (database: DatabaseSync) => void) {
    this.#database = new DatabaseSync(':memory:')
    initialize?.(this.#database)
    this.#database.exec('PRAGMA foreign_keys = ON')
    this.#database.exec('PRAGMA journal_mode = MEMORY')
    this.#database.exec('PRAGMA temp_store = MEMORY')
    this.#migrate()
  }

  #migrate() {
    const version = this.#database.prepare('PRAGMA user_version').get() as { user_version: number }
    if (version.user_version > SCHEMA_VERSION) {
      throw new Error('The template database schema is newer than this server supports.')
    }

    this.#database.exec('BEGIN IMMEDIATE')
    try {
      this.#database.exec(`
      CREATE TABLE IF NOT EXISTS templates (
        template_id TEXT PRIMARY KEY,
        template_json TEXT NOT NULL CHECK (json_valid(template_json))
      ) STRICT;

      CREATE TABLE IF NOT EXISTS template_assets (
        asset_id TEXT PRIMARY KEY,
        template_id TEXT NOT NULL,
        content_type TEXT NOT NULL CHECK (content_type LIKE 'image/%'),
        asset_data BLOB NOT NULL CHECK (length(asset_data) > 0),
        FOREIGN KEY (template_id) REFERENCES templates(template_id) ON DELETE CASCADE
      ) STRICT;

      CREATE INDEX IF NOT EXISTS template_assets_template_id_idx
        ON template_assets (template_id);
      `)

      if (version.user_version < 2) {
        this.#database.exec(`
          CREATE TABLE IF NOT EXISTS template_metadata (
            template_id TEXT PRIMARY KEY,
            kind TEXT NOT NULL CHECK (kind IN ('diagram', 'commentary')),
            description TEXT NOT NULL,
            source TEXT NOT NULL CHECK (source IN ('builtin', 'import')),
            checksum TEXT,
            created_at TEXT NOT NULL,
            FOREIGN KEY (template_id) REFERENCES templates(template_id) ON DELETE CASCADE
          ) STRICT;

          CREATE TABLE IF NOT EXISTS template_previews (
            template_id TEXT PRIMARY KEY,
            content_type TEXT NOT NULL CHECK (content_type = 'image/png'),
            preview_data BLOB NOT NULL CHECK (length(preview_data) > 0),
            width INTEGER NOT NULL CHECK (width > 0),
            height INTEGER NOT NULL CHECK (height > 0),
            FOREIGN KEY (template_id) REFERENCES templates(template_id) ON DELETE CASCADE
          ) STRICT;

          INSERT OR IGNORE INTO template_metadata (
            template_id, kind, description, source, checksum, created_at
          )
          SELECT template_id, 'diagram', '', 'import', NULL, CURRENT_TIMESTAMP
          FROM templates;
        `)
      }

      if (version.user_version < 3) {
        this.#database.exec(`
          CREATE TABLE IF NOT EXISTS apps (
            app_id TEXT PRIMARY KEY,
            display_name TEXT NOT NULL,
            metadata_json TEXT NOT NULL CHECK (json_valid(metadata_json)),
            created_at TEXT NOT NULL,
            last_seen_at TEXT NOT NULL
          ) STRICT;

          CREATE TABLE IF NOT EXISTS app_templates (
            app_id TEXT NOT NULL,
            template_id TEXT NOT NULL,
            created_at TEXT NOT NULL,
            PRIMARY KEY (app_id, template_id),
            FOREIGN KEY (app_id) REFERENCES apps(app_id) ON DELETE CASCADE,
            FOREIGN KEY (template_id) REFERENCES templates(template_id) ON DELETE CASCADE
          ) STRICT;

          CREATE INDEX IF NOT EXISTS app_templates_template_id_idx
            ON app_templates (template_id);

          INSERT OR IGNORE INTO apps (
            app_id, display_name, metadata_json, created_at, last_seen_at
          ) VALUES (
            '${DEFAULT_APP_ID}', '${DEFAULT_APP_ID}', json('{}'), CURRENT_TIMESTAMP, CURRENT_TIMESTAMP
          );

          INSERT OR IGNORE INTO app_templates (app_id, template_id, created_at)
          SELECT '${DEFAULT_APP_ID}', template_id, CURRENT_TIMESTAMP
          FROM templates;
        `)
      }

      if (version.user_version < 5) {
        this.#database.exec(`
          DROP TABLE IF EXISTS slide_classifications_fts;
          DROP TABLE IF EXISTS slide_classifications;

          CREATE TABLE slide_classifications (
            template_id TEXT PRIMARY KEY,
            status TEXT NOT NULL CHECK (status IN ('pending', 'processing', 'ready', 'failed')),
            metadata_json TEXT CHECK (metadata_json IS NULL OR json_valid(metadata_json)),
            archetype TEXT,
            layout_type TEXT,
            content_density TEXT CHECK (
              content_density IS NULL OR content_density IN ('low', 'medium', 'high')
            ),
            embedding_status TEXT NOT NULL DEFAULT 'not_ready' CHECK (
              embedding_status IN ('not_ready', 'pending', 'processing', 'ready', 'failed')
            ),
            subject_embedding_document TEXT,
            subject_embedding BLOB,
            subject_embedding_fingerprint TEXT,
            capability_embedding_document TEXT,
            capability_embedding BLOB,
            capability_embedding_fingerprint TEXT,
            embedding_model TEXT,
            embedding_dimensions INTEGER CHECK (embedding_dimensions IS NULL OR embedding_dimensions > 0),
            embedding_attempt_count INTEGER NOT NULL DEFAULT 0,
            embedding_next_attempt_at TEXT,
            embedding_lease_expires_at TEXT,
            embedding_last_error_code TEXT,
            input_fingerprint TEXT NOT NULL,
            classified_fingerprint TEXT,
            schema_version INTEGER NOT NULL CHECK (schema_version = 2),
            prompt_version TEXT NOT NULL,
            model TEXT,
            attempt_count INTEGER NOT NULL DEFAULT 0,
            next_attempt_at TEXT,
            lease_expires_at TEXT,
            classified_at TEXT,
            last_error_code TEXT,
            FOREIGN KEY (template_id) REFERENCES templates(template_id) ON DELETE CASCADE
          ) STRICT;

          CREATE INDEX slide_classifications_classification_jobs_idx
            ON slide_classifications (status, next_attempt_at, lease_expires_at);
          CREATE INDEX slide_classifications_embedding_jobs_idx
            ON slide_classifications (embedding_status, embedding_next_attempt_at, embedding_lease_expires_at);
          CREATE INDEX slide_classifications_schema_idx
            ON slide_classifications (schema_version, status, embedding_status);

          CREATE VIRTUAL TABLE slide_classifications_fts USING fts5(
            template_id UNINDEXED,
            title,
            subject_summary,
            domains,
            topics,
            technologies,
            entities,
            claims,
            synonyms,
            intents,
            information_types,
            audience,
            archetype,
            content_slots,
            visual_structure,
            retrieval_keywords,
            tokenize = 'unicode61 remove_diacritics 2'
          );
        `)
      }

      this.#database.exec(`PRAGMA user_version = ${SCHEMA_VERSION}`)
      this.#database.exec('COMMIT')
    } catch (error) {
      this.#database.exec('ROLLBACK')
      if (error instanceof Error && /no such module:\s*fts5/iu.test(error.message)) {
        throw new Error('The SQLite runtime must provide the FTS5 extension.', { cause: error })
      }
      throw error
    }
  }

  insert(
    template: StoredTemplate,
    assets: readonly TemplateAsset[],
    metadata: StoredTemplateMetadata = defaultMetadata(template.templateId),
    preview?: StoredTemplatePreview,
    appId: string = DEFAULT_APP_ID,
  ) {
    this.insertMany([{ assets, metadata, preview, template }], appId)
  }

  insertMany(records: readonly TemplateInsert[], appId: string = DEFAULT_APP_ID) {
    this.#database.exec('BEGIN IMMEDIATE')
    try {
      for (const record of records) {
        this.#writeTemplateAndAssets(
          'insert',
          record.template,
          record.assets,
          record.metadata ?? defaultMetadata(record.template.templateId),
          record.preview,
          appId,
        )
      }
      this.#database.exec('COMMIT')
    } catch (error) {
      this.#database.exec('ROLLBACK')
      throw error
    }
  }

  insertWithPendingClassification(
    record: TemplateInsert,
    pending: PendingSlideClassification,
    appId: string = DEFAULT_APP_ID,
  ) {
    this.#insertWithClassification(record, pending, 'pending', 0, appId)
  }

  insertWithProcessingClassification(
    record: TemplateInsert,
    pending: PendingSlideClassification,
    appId: string = DEFAULT_APP_ID,
  ) {
    this.#insertWithClassification(record, pending, 'processing', 1, appId)
  }

  #insertWithClassification(
    record: TemplateInsert,
    pending: PendingSlideClassification,
    status: 'pending' | 'processing',
    attemptCount: number,
    appId: string,
  ) {
    this.#database.exec('BEGIN IMMEDIATE')
    try {
      this.#writeTemplateAndAssets(
        'insert',
        record.template,
        record.assets,
        record.metadata ?? defaultMetadata(record.template.templateId),
        record.preview,
        appId,
      )
      this.#database.prepare(`
        INSERT INTO slide_classifications (
          template_id, status, embedding_status, input_fingerprint, schema_version,
          prompt_version, attempt_count
        ) VALUES (?, ?, 'not_ready', ?, ?, ?, ?)
      `).run(
        record.template.templateId,
        status,
        pending.inputFingerprint,
        pending.schemaVersion,
        pending.promptVersion,
        attemptCount,
      )
      this.#database.exec('COMMIT')
    } catch (error) {
      this.#database.exec('ROLLBACK')
      throw error
    }
  }

  update(
    template: StoredTemplate,
    assets: readonly TemplateAsset[],
    appId: string = DEFAULT_APP_ID,
  ) {
    this.#database.exec('BEGIN IMMEDIATE')
    try {
      this.#writeTemplateAndAssets('update', template, assets, undefined, undefined, appId)
      this.#database.exec('COMMIT')
    } catch (error) {
      this.#database.exec('ROLLBACK')
      throw error
    }
  }

  #writeTemplateAndAssets(
    operation: 'insert' | 'update',
    template: StoredTemplate,
    assets: readonly TemplateAsset[],
    metadata?: StoredTemplateMetadata,
    preview?: StoredTemplatePreview,
    appId: string = DEFAULT_APP_ID,
  ) {
    const serializedTemplate = JSON.stringify(template.templateJson)
    if (operation === 'insert') {
      this.#database
        .prepare('INSERT INTO templates (template_id, template_json) VALUES (?, json(?))')
        .run(template.templateId, serializedTemplate)
    } else {
      const result = this.#database.prepare(`
        UPDATE templates
        SET template_json = json(?)
        WHERE template_id = ?
          AND EXISTS (
            SELECT 1 FROM app_templates
            WHERE app_id = ? AND template_id = templates.template_id
          )
      `).run(serializedTemplate, template.templateId, appId)
      if (result.changes === 0) {
        throw new Error('The template is not available to this app.')
      }
    }

    if (metadata) {
      assertTemplateOwnership(template.templateId, metadata.templateId)
      this.#database.prepare(`
          INSERT INTO template_metadata (
            template_id, kind, description, source, checksum, created_at
          ) VALUES (?, ?, ?, ?, ?, ?)
          ON CONFLICT(template_id) DO UPDATE SET
            kind = excluded.kind,
            description = excluded.description,
            source = excluded.source,
            checksum = excluded.checksum,
            created_at = excluded.created_at
      `).run(
        metadata.templateId,
        metadata.kind,
        metadata.description,
        metadata.source,
        metadata.checksum,
        metadata.createdAt,
      )
    }

    const insertAsset = this.#database.prepare(`
        INSERT INTO template_assets (asset_id, template_id, content_type, asset_data)
        VALUES (?, ?, ?, ?)
      `)
    for (const asset of assets) {
      if (asset.templateId !== template.templateId) {
        throw new Error('A template asset cannot be stored under a different template.')
      }
      insertAsset.run(asset.assetId, asset.templateId, asset.contentType, asset.bytes)
    }

    if (preview) {
      assertTemplateOwnership(template.templateId, preview.templateId)
      this.#database.prepare(`
          INSERT INTO template_previews (
            template_id, content_type, preview_data, width, height
          ) VALUES (?, ?, ?, ?, ?)
          ON CONFLICT(template_id) DO UPDATE SET
            content_type = excluded.content_type,
            preview_data = excluded.preview_data,
            width = excluded.width,
            height = excluded.height
      `).run(
        preview.templateId,
        preview.contentType,
        preview.bytes,
        preview.width,
        preview.height,
      )
    }

    if (operation === 'insert') {
      this.#database.prepare(`
        INSERT INTO app_templates (app_id, template_id, created_at)
        VALUES (?, ?, CURRENT_TIMESTAMP)
      `).run(appId, template.templateId)
    }
  }

  upsertBuiltin(
    template: StoredTemplate,
    assets: readonly TemplateAsset[],
    metadata: StoredTemplateMetadata,
    preview: StoredTemplatePreview,
  ) {
    assertTemplateOwnership(template.templateId, metadata.templateId)
    assertTemplateOwnership(template.templateId, preview.templateId)
    const existing = this.#database.prepare(`
      SELECT source, checksum FROM template_metadata WHERE template_id = ?
    `).get(template.templateId) as { checksum: string | null; source: string } | undefined

    if (existing?.source === 'import' || existing?.checksum === metadata.checksum) {
      return
    }

    this.#database.exec('BEGIN IMMEDIATE')
    try {
      this.#database.prepare(`
        INSERT INTO templates (template_id, template_json) VALUES (?, json(?))
        ON CONFLICT(template_id) DO UPDATE SET template_json = excluded.template_json
      `).run(template.templateId, JSON.stringify(template.templateJson))
      this.#database.prepare('DELETE FROM template_assets WHERE template_id = ?').run(template.templateId)
      this.#database.prepare('DELETE FROM template_previews WHERE template_id = ?').run(template.templateId)

      const insertAsset = this.#database.prepare(`
        INSERT INTO template_assets (asset_id, template_id, content_type, asset_data)
        VALUES (?, ?, ?, ?)
      `)
      for (const asset of assets) {
        assertTemplateOwnership(template.templateId, asset.templateId)
        insertAsset.run(asset.assetId, asset.templateId, asset.contentType, asset.bytes)
      }

      this.#database.prepare(`
        INSERT INTO template_metadata (
          template_id, kind, description, source, checksum, created_at
        ) VALUES (?, ?, ?, ?, ?, ?)
        ON CONFLICT(template_id) DO UPDATE SET
          kind = excluded.kind,
          description = excluded.description,
          source = excluded.source,
          checksum = excluded.checksum,
          created_at = excluded.created_at
      `).run(
        metadata.templateId,
        metadata.kind,
        metadata.description,
        metadata.source,
        metadata.checksum,
        metadata.createdAt,
      )
      this.#database.prepare(`
        INSERT INTO template_previews (
          template_id, content_type, preview_data, width, height
        ) VALUES (?, ?, ?, ?, ?)
      `).run(
        preview.templateId,
        preview.contentType,
        preview.bytes,
        preview.width,
        preview.height,
      )
      this.#database.prepare(`
        INSERT OR IGNORE INTO app_templates (app_id, template_id, created_at)
        SELECT app_id, ?, CURRENT_TIMESTAMP FROM apps
      `).run(template.templateId)
      this.#database.exec('COMMIT')
    } catch (error) {
      this.#database.exec('ROLLBACK')
      throw error
    }
  }

  findAsset(templateId: string, assetId: string, appId: string = DEFAULT_APP_ID) {
    const row = this.#database
      .prepare(`
        SELECT asset_id, template_id, content_type, asset_data
        FROM template_assets
        JOIN app_templates USING (template_id)
        WHERE app_templates.app_id = ?
          AND template_assets.template_id = ?
          AND asset_id = ?
      `)
      .get(appId, templateId, assetId) as TemplateAssetRow | undefined

    if (!row) {
      return undefined
    }

    return {
      assetId: row.asset_id,
      bytes: Buffer.from(row.asset_data),
      contentType: row.content_type,
      templateId: row.template_id,
    }
  }

  findById(templateId: string, appId: string = DEFAULT_APP_ID) {
    const row = this.#database
      .prepare(`
        SELECT templates.template_id, templates.template_json
        FROM templates
        JOIN app_templates USING (template_id)
        WHERE app_templates.app_id = ? AND templates.template_id = ?
      `)
      .get(appId, templateId) as TemplateRow | undefined

    if (!row) {
      return undefined
    }

    return {
      templateId: row.template_id,
      templateJson: JSON.parse(row.template_json) as PowerPointCanvasJson,
    }
  }

  hasTemplate(templateId: string, appId: string = DEFAULT_APP_ID) {
    const row = this.#database.prepare(`
      SELECT 1
      FROM app_templates
      WHERE app_id = ? AND template_id = ?
    `).get(appId, templateId) as { 1: number } | undefined
    return row !== undefined
  }

  list(kind?: TemplateKind, appId: string = DEFAULT_APP_ID): StoredTemplateSummary[] {
    const rows = this.#database
      .prepare(`
        SELECT
          templates.template_id,
          templates.template_json,
          template_metadata.kind,
          template_metadata.description,
          template_metadata.source,
          template_metadata.checksum,
          template_metadata.created_at,
          CASE WHEN template_previews.template_id IS NULL THEN 0 ELSE 1 END AS preview_available
        FROM templates
        JOIN app_templates ON app_templates.template_id = templates.template_id
        JOIN template_metadata ON template_metadata.template_id = templates.template_id
        LEFT JOIN template_previews ON template_previews.template_id = templates.template_id
        WHERE app_templates.app_id = ?
          AND (? IS NULL OR template_metadata.kind = ?)
        ORDER BY template_metadata.created_at DESC, templates.rowid DESC
      `)
      .all(appId, kind ?? null, kind ?? null) as TemplateSummaryRow[]

    return rows.map((row) => ({
      templateId: row.template_id,
      templateJson: JSON.parse(row.template_json) as PowerPointCanvasJson,
      metadata: {
        checksum: row.checksum,
        createdAt: row.created_at,
        description: row.description,
        kind: row.kind,
        source: row.source,
        templateId: row.template_id,
      },
      previewAvailable: row.preview_available === 1,
    }))
  }

  listPreviews(
    limit: number,
    offset: number,
    appId: string = DEFAULT_APP_ID,
  ): StoredTemplatePreviewPage {
    const { total } = this.#database
      .prepare(`
        SELECT COUNT(*) AS total
        FROM template_previews
        JOIN app_templates USING (template_id)
        WHERE app_templates.app_id = ?
      `)
      .get(appId) as { total: number }
    const rows = this.#database.prepare(`
      SELECT
        template_previews.template_id,
        template_previews.content_type,
        template_previews.preview_data,
        template_previews.width,
        template_previews.height
      FROM template_previews
      JOIN app_templates ON app_templates.template_id = template_previews.template_id
      JOIN templates ON templates.template_id = template_previews.template_id
      JOIN template_metadata ON template_metadata.template_id = template_previews.template_id
      WHERE app_templates.app_id = ?
      ORDER BY template_metadata.created_at DESC, templates.rowid DESC
      LIMIT ? OFFSET ?
    `).all(appId, limit, offset) as TemplatePreviewRow[]

    return {
      previews: rows.map((row) => ({
        bytes: Buffer.from(row.preview_data),
        contentType: row.content_type,
        height: row.height,
        templateId: row.template_id,
        width: row.width,
      })),
      total,
    }
  }

  findPreview(templateId: string, appId: string = DEFAULT_APP_ID) {
    const row = this.#database.prepare(`
      SELECT template_id, content_type, preview_data, width, height
      FROM template_previews
      JOIN app_templates USING (template_id)
      WHERE app_templates.app_id = ? AND template_previews.template_id = ?
    `).get(appId, templateId) as TemplatePreviewRow | undefined

    return row
      ? {
          bytes: Buffer.from(row.preview_data),
          contentType: row.content_type,
          height: row.height,
          templateId: row.template_id,
          width: row.width,
        }
      : undefined
  }

  findSlideClassification(templateId: string): StoredSlideClassification | undefined {
    const row = this.#database.prepare(`
      SELECT * FROM slide_classifications WHERE template_id = ?
    `).get(templateId) as ClassificationRow | undefined
    return row ? mapClassificationRow(row) : undefined
  }

  claimNextSlideClassification(now: string, leaseExpiresAt: string): SlideClassificationJob | undefined {
    this.#database.exec('BEGIN IMMEDIATE')
    try {
      this.#database.prepare(`
        UPDATE slide_classifications
        SET status = 'pending', lease_expires_at = NULL
        WHERE status = 'processing' AND lease_expires_at <= ?
      `).run(now)
      const candidate = this.#database.prepare(`
        SELECT template_id
        FROM slide_classifications
        WHERE status = 'pending' AND (next_attempt_at IS NULL OR next_attempt_at <= ?)
        ORDER BY COALESCE(next_attempt_at, ''), template_id
        LIMIT 1
      `).get(now) as { template_id: string } | undefined
      if (!candidate) {
        this.#database.exec('COMMIT')
        return undefined
      }
      this.#database.prepare(`
        UPDATE slide_classifications
        SET status = 'processing', attempt_count = attempt_count + 1,
            lease_expires_at = ?, next_attempt_at = NULL, last_error_code = NULL
        WHERE template_id = ? AND status = 'pending'
      `).run(leaseExpiresAt, candidate.template_id)
      const row = this.#database.prepare(`
        SELECT
          templates.template_id,
          templates.template_json,
          template_metadata.kind,
          template_previews.content_type,
          template_previews.preview_data,
          template_previews.width,
          template_previews.height,
          slide_classifications.attempt_count,
          slide_classifications.input_fingerprint
        FROM templates
        JOIN template_metadata USING (template_id)
        JOIN slide_classifications USING (template_id)
        LEFT JOIN template_previews USING (template_id)
        WHERE templates.template_id = ?
      `).get(candidate.template_id) as {
        attempt_count: number
        content_type: 'image/png' | null
        height: number | null
        input_fingerprint: string
        kind: TemplateKind
        preview_data: Uint8Array | null
        template_id: string
        template_json: string
        width: number | null
      }
      this.#database.exec('COMMIT')
      const templateJson = JSON.parse(row.template_json) as PowerPointCanvasJson
      const job: SlideClassificationJob = {
        attemptCount: row.attempt_count,
        inputFingerprint: row.input_fingerprint,
        kind: row.kind,
        template: { templateId: row.template_id, templateJson },
        title: templateJson.presentation.title,
      }
      if (
        row.content_type === 'image/png'
        && row.preview_data !== null
        && row.width !== null
        && row.height !== null
      ) {
        job.preview = {
          bytes: Buffer.from(row.preview_data),
          contentType: row.content_type,
          height: row.height,
          templateId: row.template_id,
          width: row.width,
        }
      }
      return job
    } catch (error) {
      this.#database.exec('ROLLBACK')
      throw error
    }
  }

  completeSlideClassification(result: CompleteSlideClassification) {
    const metadata = normalizeSlideRetrievalMetadata(result.metadata)
    this.#database.exec('BEGIN IMMEDIATE')
    try {
      const update = this.#database.prepare(`
        UPDATE slide_classifications
        SET status = 'ready', metadata_json = json(?), archetype = ?, layout_type = ?,
            content_density = ?,
            classified_fingerprint = ?, model = ?, classified_at = ?,
            lease_expires_at = NULL, next_attempt_at = NULL, last_error_code = NULL,
            embedding_status = 'pending', embedding_model = NULL, embedding_dimensions = NULL,
            subject_embedding_document = ?, subject_embedding = NULL,
            subject_embedding_fingerprint = ?, capability_embedding_document = ?,
            capability_embedding = NULL, capability_embedding_fingerprint = ?,
            embedding_next_attempt_at = NULL, embedding_lease_expires_at = NULL,
            embedding_last_error_code = NULL
        WHERE template_id = ? AND status = 'processing' AND input_fingerprint = ?
      `).run(
        JSON.stringify(metadata),
        metadata.template_fit.archetype,
        metadata.visual.layout_type,
        metadata.visual.content_density,
        result.classifiedFingerprint,
        result.model,
        result.classifiedAt,
        result.subjectEmbeddingDocument,
        result.subjectEmbeddingFingerprint,
        result.capabilityEmbeddingDocument,
        result.capabilityEmbeddingFingerprint,
        result.templateId,
        result.classifiedFingerprint,
      )
      if (update.changes !== 1) throw new Error('The classification job lease is no longer active.')
      this.#replaceFtsRow(result.templateId, metadata)
      this.#database.exec('COMMIT')
    } catch (error) {
      this.#database.exec('ROLLBACK')
      throw error
    }
  }

  recordSlideClassificationFailure(templateId: string, errorCode: string, retryAt: string | null) {
    this.#database.prepare(`
      UPDATE slide_classifications
      SET status = ?, next_attempt_at = ?, lease_expires_at = NULL, last_error_code = ?
      WHERE template_id = ? AND status = 'processing'
    `).run(retryAt === null ? 'failed' : 'pending', retryAt, errorCode, templateId)
  }

  claimNextSlideEmbedding(now: string, leaseExpiresAt: string): SlideEmbeddingJob | undefined {
    this.#database.exec('BEGIN IMMEDIATE')
    try {
      this.#database.prepare(`
        UPDATE slide_classifications
        SET embedding_status = 'pending', embedding_lease_expires_at = NULL
        WHERE embedding_status = 'processing' AND embedding_lease_expires_at <= ?
      `).run(now)
      const row = this.#database.prepare(`
        SELECT template_id, subject_embedding_document, subject_embedding_fingerprint,
          capability_embedding_document, capability_embedding_fingerprint
        FROM slide_classifications
        WHERE embedding_status = 'pending'
          AND subject_embedding_document IS NOT NULL
          AND subject_embedding_fingerprint IS NOT NULL
          AND capability_embedding_document IS NOT NULL
          AND capability_embedding_fingerprint IS NOT NULL
          AND (embedding_next_attempt_at IS NULL OR embedding_next_attempt_at <= ?)
        ORDER BY COALESCE(embedding_next_attempt_at, ''), template_id
        LIMIT 1
      `).get(now) as {
        capability_embedding_document: string
        capability_embedding_fingerprint: string
        subject_embedding_document: string
        subject_embedding_fingerprint: string
        template_id: string
      } | undefined
      if (!row) {
        this.#database.exec('COMMIT')
        return undefined
      }
      this.#database.prepare(`
        UPDATE slide_classifications
        SET embedding_status = 'processing', embedding_attempt_count = embedding_attempt_count + 1,
            embedding_lease_expires_at = ?, embedding_next_attempt_at = NULL,
            embedding_last_error_code = NULL
        WHERE template_id = ? AND embedding_status = 'pending'
      `).run(leaseExpiresAt, row.template_id)
      const { embedding_attempt_count } = this.#database.prepare(`
        SELECT embedding_attempt_count FROM slide_classifications WHERE template_id = ?
      `).get(row.template_id) as { embedding_attempt_count: number }
      this.#database.exec('COMMIT')
      return {
        attemptCount: embedding_attempt_count,
        capabilityDocument: row.capability_embedding_document,
        capabilityFingerprint: row.capability_embedding_fingerprint,
        subjectDocument: row.subject_embedding_document,
        subjectFingerprint: row.subject_embedding_fingerprint,
        templateId: row.template_id,
      }
    } catch (error) {
      this.#database.exec('ROLLBACK')
      throw error
    }
  }

  completeSlideEmbedding(
    templateId: string,
    subjectFingerprint: string,
    capabilityFingerprint: string,
    subjectVector: readonly number[],
    capabilityVector: readonly number[],
    model: string,
    dimensions: number,
  ) {
    const subjectBytes = serializeEmbedding(subjectVector)
    const capabilityBytes = serializeEmbedding(capabilityVector)
    if (subjectVector.length !== dimensions || capabilityVector.length !== dimensions) {
      throw new Error('Embedding dimensions do not match the vectors.')
    }
    const result = this.#database.prepare(`
      UPDATE slide_classifications
      SET embedding_status = 'ready', subject_embedding = ?, capability_embedding = ?, embedding_model = ?,
          embedding_dimensions = ?, embedding_lease_expires_at = NULL,
          embedding_next_attempt_at = NULL, embedding_last_error_code = NULL
      WHERE template_id = ? AND embedding_status = 'processing'
        AND subject_embedding_fingerprint = ? AND capability_embedding_fingerprint = ?
    `).run(
      subjectBytes,
      capabilityBytes,
      model,
      dimensions,
      templateId,
      subjectFingerprint,
      capabilityFingerprint,
    )
    if (result.changes !== 1) throw new Error('The embedding job lease is no longer active.')
  }

  completeSlideProcessing(result: CompleteSlideProcessing) {
    const metadata = normalizeSlideRetrievalMetadata(result.metadata)
    const subjectBytes = serializeEmbedding(result.subjectVector)
    const capabilityBytes = serializeEmbedding(result.capabilityVector)
    if (
      result.subjectVector.length !== result.embeddingDimensions
      || result.capabilityVector.length !== result.embeddingDimensions
    ) {
      throw new Error('Embedding dimensions do not match the vectors.')
    }

    this.#database.exec('BEGIN IMMEDIATE')
    try {
      const update = this.#database.prepare(`
        UPDATE slide_classifications
        SET status = 'ready', metadata_json = json(?), archetype = ?, layout_type = ?,
            content_density = ?,
            classified_fingerprint = ?, model = ?, classified_at = ?,
            lease_expires_at = NULL, next_attempt_at = NULL, last_error_code = NULL,
            embedding_status = 'ready', subject_embedding_document = ?, subject_embedding = ?,
            subject_embedding_fingerprint = ?, capability_embedding_document = ?,
            capability_embedding = ?, capability_embedding_fingerprint = ?,
            embedding_model = ?, embedding_dimensions = ?,
            embedding_attempt_count = 1, embedding_next_attempt_at = NULL,
            embedding_lease_expires_at = NULL, embedding_last_error_code = NULL
        WHERE template_id = ? AND status = 'processing' AND input_fingerprint = ?
      `).run(
        JSON.stringify(metadata),
        metadata.template_fit.archetype,
        metadata.visual.layout_type,
        metadata.visual.content_density,
        result.classifiedFingerprint,
        result.model,
        result.classifiedAt,
        result.subjectEmbeddingDocument,
        subjectBytes,
        result.subjectEmbeddingFingerprint,
        result.capabilityEmbeddingDocument,
        capabilityBytes,
        result.capabilityEmbeddingFingerprint,
        result.embeddingModel,
        result.embeddingDimensions,
        result.templateId,
        result.classifiedFingerprint,
      )
      if (update.changes !== 1) {
        throw new Error('The synchronous slide processing state is no longer active.')
      }
      this.#replaceFtsRow(result.templateId, metadata)
      this.#database.exec('COMMIT')
    } catch (error) {
      this.#database.exec('ROLLBACK')
      throw error
    }
  }

  recordSlideEmbeddingFailure(templateId: string, errorCode: string, retryAt: string | null) {
    this.#database.prepare(`
      UPDATE slide_classifications
      SET embedding_status = ?, embedding_next_attempt_at = ?,
          embedding_lease_expires_at = NULL, embedding_last_error_code = ?
      WHERE template_id = ? AND embedding_status = 'processing'
    `).run(retryAt === null ? 'failed' : 'pending', retryAt, errorCode, templateId)
  }

  refreshPendingSlideClassification(templateId: string, pending: PendingSlideClassification) {
    const result = this.#database.prepare(`
      UPDATE slide_classifications
      SET status = 'pending', input_fingerprint = ?, schema_version = ?, prompt_version = ?,
          attempt_count = 0, next_attempt_at = NULL, lease_expires_at = NULL, last_error_code = NULL
      WHERE template_id = ? AND input_fingerprint <> ?
    `).run(
      pending.inputFingerprint,
      pending.schemaVersion,
      pending.promptVersion,
      templateId,
      pending.inputFingerprint,
    )
    return result.changes === 1
  }

  refreshPendingSlideEmbedding(
    templateId: string,
    subjectDocument: string,
    subjectFingerprint: string,
    capabilityDocument: string,
    capabilityFingerprint: string,
  ) {
    const result = this.#database.prepare(`
      UPDATE slide_classifications
      SET embedding_status = 'pending', subject_embedding_document = ?,
          subject_embedding_fingerprint = ?, capability_embedding_document = ?,
          capability_embedding_fingerprint = ?, subject_embedding = NULL,
          capability_embedding = NULL,
          embedding_attempt_count = 0, embedding_next_attempt_at = NULL,
          embedding_lease_expires_at = NULL, embedding_last_error_code = NULL
      WHERE template_id = ? AND metadata_json IS NOT NULL
        AND (
          subject_embedding_fingerprint IS NULL OR subject_embedding_fingerprint <> ?
          OR capability_embedding_fingerprint IS NULL OR capability_embedding_fingerprint <> ?
        )
    `).run(
      subjectDocument,
      subjectFingerprint,
      capabilityDocument,
      capabilityFingerprint,
      templateId,
      subjectFingerprint,
      capabilityFingerprint,
    )
    return result.changes === 1
  }

  listRetrievalClassifications(appId: string): RetrievalClassificationRecord[] {
    const rows = this.#database.prepare(`
      SELECT
        slide_classifications.*,
        templates.template_json,
        template_metadata.kind,
        template_metadata.created_at,
        CASE WHEN template_previews.template_id IS NULL THEN 0 ELSE 1 END AS preview_available
      FROM slide_classifications
      JOIN templates USING (template_id)
      JOIN template_metadata USING (template_id)
      JOIN app_templates USING (template_id)
      LEFT JOIN template_previews USING (template_id)
      WHERE app_templates.app_id = ?
        AND slide_classifications.metadata_json IS NOT NULL
        AND slide_classifications.schema_version = 2
      ORDER BY template_metadata.created_at DESC, slide_classifications.template_id
    `).all(appId) as RetrievalRow[]
    return rows.map((row) => {
      const template = JSON.parse(row.template_json) as PowerPointCanvasJson
      return {
        classification: mapClassificationRow(row),
        createdAt: row.created_at,
        kind: row.kind,
        previewAvailable: row.preview_available === 1,
        templateId: row.template_id,
        title: template.presentation.title,
      }
    })
  }

  searchSlideClassifications(appId: string, ftsQuery: string, limit: number): string[] {
    const rows = this.#database.prepare(`
      SELECT slide_classifications_fts.template_id
      FROM slide_classifications_fts
      JOIN app_templates ON app_templates.template_id = slide_classifications_fts.template_id
      JOIN slide_classifications ON slide_classifications.template_id = slide_classifications_fts.template_id
      WHERE app_templates.app_id = ?
        AND slide_classifications.metadata_json IS NOT NULL
        AND slide_classifications.schema_version = 2
        AND slide_classifications_fts MATCH ?
      ORDER BY bm25(slide_classifications_fts, 0.0, 2.0, 5.0, 5.0, 5.0, 2.0, 2.0, 3.0, 3.0, 4.0, 3.0, 1.0, 4.0, 4.0, 2.0, 3.0),
        slide_classifications_fts.template_id
      LIMIT ?
    `).all(appId, ftsQuery, limit) as Array<{ template_id: string }>
    return rows.map((row) => row.template_id)
  }

  #replaceFtsRow(templateId: string, metadata: ReturnType<typeof normalizeSlideRetrievalMetadata>) {
    const templateRow = this.#database.prepare(`
      SELECT template_json FROM templates WHERE template_id = ?
    `).get(templateId) as { template_json: string } | undefined
    if (!templateRow) throw new Error('The classified template no longer exists.')
    const template = JSON.parse(templateRow.template_json) as PowerPointCanvasJson
    this.#database.prepare('DELETE FROM slide_classifications_fts WHERE template_id = ?').run(templateId)
    this.#database.prepare(`
      INSERT INTO slide_classifications_fts (
        template_id, title, subject_summary, domains, topics, technologies, entities,
        claims, synonyms, intents, information_types, audience, archetype, content_slots,
        visual_structure, retrieval_keywords
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `).run(
      templateId,
      template.presentation.title,
      metadata.subject.summary,
      metadata.subject.domains.map((domain) => domain.id).join(' '),
      [
        ...metadata.subject.domains.flatMap((domain) => domain.topics),
        ...metadata.subject.other_topics,
      ].join(' '),
      metadata.subject.technologies.join(' '),
      metadata.subject.entities.join(' '),
      metadata.subject.claims.join(' '),
      metadata.subject.synonyms.join(' '),
      metadata.communication.intents.join(' '),
      metadata.communication.information_types.join(' '),
      metadata.communication.audience.join(' '),
      metadata.template_fit.archetype,
      metadata.template_fit.content_slots.map((slot) => `${slot.role} ${slot.capacity}`).join(' '),
      [metadata.visual.layout_type, ...metadata.visual.visual_elements, ...metadata.visual.structural_features].join(' '),
      metadata.retrieval_keywords.join(' '),
    )
  }

  delete(templateId: string, appId: string = DEFAULT_APP_ID) {
    this.#database.exec('BEGIN IMMEDIATE')
    try {
      const result = this.#database.prepare(`
        DELETE FROM app_templates WHERE app_id = ? AND template_id = ?
      `).run(appId, templateId)
      this.#database.prepare(`
        DELETE FROM slide_classifications_fts
        WHERE template_id = ?
          AND NOT EXISTS (SELECT 1 FROM app_templates WHERE template_id = ?)
      `).run(templateId, templateId)
      this.#database.prepare(`
        DELETE FROM templates
        WHERE template_id = ?
          AND NOT EXISTS (
            SELECT 1 FROM app_templates WHERE template_id = templates.template_id
          )
      `).run(templateId)
      this.#database.exec('COMMIT')
      return result.changes > 0
    } catch (error) {
      this.#database.exec('ROLLBACK')
      throw error
    }
  }

  findByIdWithAssets(
    templateId: string,
    appId: string = DEFAULT_APP_ID,
  ): StoredTemplateWithAssets | undefined {
    const rows = this.#database
      .prepare(`
        SELECT
          templates.template_id,
          templates.template_json,
          template_assets.asset_id,
          template_assets.template_id AS asset_template_id,
          template_assets.content_type,
          template_assets.asset_data
        FROM templates
        JOIN app_templates ON app_templates.template_id = templates.template_id
        LEFT JOIN template_assets
          ON template_assets.template_id = templates.template_id
        WHERE app_templates.app_id = ? AND templates.template_id = ?
        ORDER BY template_assets.asset_id
      `)
      .all(appId, templateId) as TemplateWithAssetRow[]

    const templateRow = rows[0]
    if (!templateRow) {
      return undefined
    }

    const assets: TemplateAsset[] = []
    for (const row of rows) {
      if (
        row.asset_id === null
        || row.asset_template_id === null
        || row.content_type === null
        || row.asset_data === null
      ) {
        continue
      }
      assets.push({
        assetId: row.asset_id,
        bytes: Buffer.from(row.asset_data),
        contentType: row.content_type,
        templateId: row.asset_template_id,
      })
    }

    return {
      assets,
      templateId: templateRow.template_id,
      templateJson: JSON.parse(templateRow.template_json) as PowerPointCanvasJson,
    }
  }

  close() {
    this.#database.close()
  }

  ensureApp(appId: string, registration: AppRegistration = {}): StoredApp {
    const displayName = registration.displayName ?? appId
    const metadataJson = JSON.stringify(registration.metadata ?? {})
    const isNewApp = this.findApp(appId) === undefined
    const updateMetadata = registration.displayName !== undefined || registration.metadata !== undefined
    this.#database.exec('BEGIN IMMEDIATE')
    try {
      const conflictClause = updateMetadata
        ? `DO UPDATE SET
            display_name = excluded.display_name,
            metadata_json = excluded.metadata_json,
            last_seen_at = CURRENT_TIMESTAMP`
        : 'DO UPDATE SET last_seen_at = CURRENT_TIMESTAMP'
      this.#database.prepare(`
        INSERT INTO apps (app_id, display_name, metadata_json, created_at, last_seen_at)
        VALUES (?, ?, json(?), CURRENT_TIMESTAMP, CURRENT_TIMESTAMP)
        ON CONFLICT(app_id) ${conflictClause}
      `).run(appId, displayName, metadataJson)
      if (isNewApp) {
        this.#database.prepare(`
          INSERT OR IGNORE INTO app_templates (app_id, template_id, created_at)
          SELECT ?, template_metadata.template_id, CURRENT_TIMESTAMP
          FROM template_metadata
          WHERE template_metadata.source = 'builtin'
        `).run(appId)
      }
      this.#database.exec('COMMIT')
    } catch (error) {
      this.#database.exec('ROLLBACK')
      throw error
    }

    const app = this.findApp(appId)
    if (!app) {
      throw new Error('The app could not be registered.')
    }
    return app
  }

  findApp(appId: string): StoredApp | undefined {
    const row = this.#database.prepare(`
      SELECT app_id, display_name, metadata_json, created_at, last_seen_at
      FROM apps
      WHERE app_id = ?
    `).get(appId) as AppRow | undefined

    return row
      ? {
          appId: row.app_id,
          createdAt: row.created_at,
          displayName: row.display_name,
          lastSeenAt: row.last_seen_at,
          metadata: parseAppMetadata(row.metadata_json),
        }
      : undefined
  }
}

function mapClassificationRow(row: ClassificationRow): StoredSlideClassification {
  const metadata = row.metadata_json === null
    ? null
    : normalizeSlideRetrievalMetadata(JSON.parse(row.metadata_json) as unknown)
  return {
    attemptCount: row.attempt_count,
    classifiedAt: row.classified_at,
    classifiedFingerprint: row.classified_fingerprint,
    embeddingAttemptCount: row.embedding_attempt_count,
    embeddingDimensions: row.embedding_dimensions,
    capabilityEmbeddingDocument: row.capability_embedding_document,
    capabilityEmbeddingFingerprint: row.capability_embedding_fingerprint,
    capabilityVector: row.capability_embedding === null || row.embedding_dimensions === null
      ? null
      : deserializeEmbedding(row.capability_embedding, row.embedding_dimensions),
    embeddingLastErrorCode: row.embedding_last_error_code,
    embeddingModel: row.embedding_model,
    embeddingNextAttemptAt: row.embedding_next_attempt_at,
    embeddingStatus: row.embedding_status,
    inputFingerprint: row.input_fingerprint,
    lastErrorCode: row.last_error_code,
    metadata,
    model: row.model,
    nextAttemptAt: row.next_attempt_at,
    promptVersion: row.prompt_version,
    schemaVersion: row.schema_version,
    status: row.status,
    templateId: row.template_id,
    subjectEmbeddingDocument: row.subject_embedding_document,
    subjectEmbeddingFingerprint: row.subject_embedding_fingerprint,
    subjectVector: row.subject_embedding === null || row.embedding_dimensions === null
      ? null
      : deserializeEmbedding(row.subject_embedding, row.embedding_dimensions),
  }
}

function defaultMetadata(templateId: string): StoredTemplateMetadata {
  return {
    checksum: null,
    createdAt: new Date().toISOString(),
    description: 'Imported PowerPoint template',
    kind: 'diagram',
    source: 'import',
    templateId,
  }
}

function assertTemplateOwnership(templateId: string, ownedTemplateId: string) {
  if (templateId !== ownedTemplateId) {
    throw new Error('Template data cannot be stored under a different template.')
  }
}

function parseAppMetadata(value: string): Record<string, string> {
  const parsed = JSON.parse(value) as unknown
  if (
    typeof parsed !== 'object'
    || parsed === null
    || Array.isArray(parsed)
    || !Object.values(parsed).every((entry) => typeof entry === 'string')
  ) {
    throw new Error('Stored app metadata is invalid.')
  }
  return parsed as Record<string, string>
}
