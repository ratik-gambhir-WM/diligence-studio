# PowerPoint import and export APIs

This package is the Node-only sibling of the React app in `../web`. It keeps its PowerPoint
implementation under `src/lib`; nothing in this package enters the React/Vite module graph.
Uploaded `.pptx` files are OOXML ZIP packages. Their
compact canvas JSON is stored in an in-memory SQLite database. Imported image bytes are stored
separately from the JSON as in-memory BLOB assets. Template metadata and first-slide PNG previews
are stored in dedicated in-memory tables for the life of the server process.
HTTP imports parse the bounded request `Buffer` directly. The headless preview provider renders
normalized canvas JSON and never writes the uploaded presentation to disk. HTTP export and deck
insertion likewise return byte arrays assembled in memory. Developer CLIs use stdin/stdout and do
not accept input or output filesystem paths. The checked-in built-in catalog is read-only startup
input; the running server creates no application data or conversion artifacts on disk.
Apps are registered in `apps`, and `app_templates` controls which templates each app can access.
The checked-in diagram and commentary catalog is seeded transactionally and idempotently at startup.

The server uses Node's built-in SQLite module and therefore requires Node.js 22.5 or newer.

## Run

```sh
npx playwright install chromium
npm run server:dev
```

Configuration is read once at startup:

- `PORT` defaults to `43127`.
- `HOST` defaults to `0.0.0.0` and must be an IP address.
- `MAX_PPTX_UPLOAD_BYTES` defaults to `26214400` (25 MiB).
- `MAX_EXPORT_JSON_BYTES` defaults to `52428800` (50 MiB).
- `TEMPLATE_PREVIEW_PROVIDER` defaults to `headless`. It accepts `headless` or `disabled`.
- `TEMPLATE_PREVIEW_RENDER_SIZE` defaults to `1600` pixels.
- `TEMPLATE_PREVIEW_RENDER_URL` defaults to
  `http://localhost:5173/_internal/template-preview`. It must point to the Vite app (or the deployed
  web app) while imports are running.
- `TEMPLATE_PREVIEW_TIMEOUT_MS` defaults to `15000`.
- `MAX_TEMPLATE_PREVIEW_BYTES` defaults to `10485760` (10 MiB).
- `REQUEST_TIMEOUT_MS` defaults to `90000` so a synchronous v2 import can finish preview,
  classification, embedding, and in-memory storage.
- `SLIDE_CLASSIFICATION_PROVIDER` is `disabled` by default and may be set to `openai`.
- When classification is enabled, `OPENAI_API_KEY`, `OPENAI_SLIDE_CLASSIFICATION_MODEL`, and
  `OPENAI_SLIDE_EMBEDDING_MODEL` are required. Use `text-embedding-3-small` for the initial pilot;
  `OPENAI_SLIDE_EMBEDDING_DIMENSIONS` is optional.
- `SLIDE_CLASSIFICATION_TIMEOUT_MS` and `SLIDE_EMBEDDING_TIMEOUT_MS` default to `45000` and
  `20000`.
- `SLIDE_CLASSIFICATION_MAX_TEXT_CHARS`, `SLIDE_CLASSIFICATION_MAX_IMAGE_BYTES`, and
  `SLIDE_EMBEDDING_MAX_TEXT_BYTES` default to `12000`, `5242880`, and `32000`.

The API key is read only by the server and must never be exposed through a `VITE_*` value. Provider
requests and routine logs omit slide content, previews, raw provider responses, and credentials.

Requests time out after 90 seconds by default. Compressed request bodies are rejected. Template,
import, and export requests accept an `X-App-Id` header (or `appId` query parameter) and fall back to
`DiligenceStudio_WestMonroe` for legacy callers. App IDs provide basic data segregation, not
authentication.

The server writes one structured JSON log when each request completes or its connection is
aborted. Logs include the request ID, method, path without its query string, status, duration,
success/failure outcome, and a sanitized error code/name when available. Use the `X-Request-Id`
response header to correlate a client-visible result with its server log. Request bodies, headers,
query strings, app IDs, filenames, error messages, and stack traces are not logged.

The default preview provider launches Chromium headlessly, injects normalized slide JSON before
the preview page loads, and screenshots only the read-only SVG slide surface. No browser window is
shown. The renderer blocks requests to origins other than the configured web-app origin. Run the
web app and API together during local imports; production must expose the internal preview route
at `TEMPLATE_PREVIEW_RENDER_URL`.

## API

### Import v2 for synchronous retrieval indexing

`POST /api/v2/import?kind=diagram|commentary` accepts the same bounded raw single-slide PowerPoint
body as v1. It stores the template first, then waits for classification and two embeddings. It
returns `201` only after v2 classification metadata, FTS content, and the subject and template-
capability vectors are committed together:

```json
{
  "previewAvailable": true,
  "templateId": "template-123",
  "templateJson": { "presentation": {} },
  "warnings": [],
  "retrieval": { "status": "ready" }
}
```

The endpoint returns `503` before conversion when the provider is disabled. If provider processing
fails after the template is saved, the request fails and the stored classification record is marked
failed without publishing partial metadata or a vector. Existing v1 import, batch import, built-in
seeding, and existing templates do not create classification work. There are no public
classification, status, retry, backfill, or retrieval routes. Retrieval is available only through
the app-bound in-process `SlideRetrievalService`/Agent adapter; semantic and hybrid text search are
the only query modes that may call the embedding provider. The adapter also exposes
`find_slides_for_finding({ markdown, limit })`, which independently ranks subject relevance and
template capability, applies exact domain/topic/intent/content-slot boosts, and returns match
reasons without exposing vectors or cross-app templates.

Import a deck by sending its binary `.pptx` body:

```sh
curl --request POST 'http://localhost:43127/api/v1/import?kind=diagram' \
  --header 'X-App-Id: DiligenceStudio_WestMonroe' \
  --header 'Content-Type: application/vnd.openxmlformats-officedocument.presentationml.presentation' \
  --data-binary @deck.pptx
```

The `201` response body is canvas JSON with a single top-level `presentation` property. It
is the same JSON contract accepted by `POST /api/v1/export`. Image `src` values are complete
`data:image/...;base64,...` URIs. The generated ID is returned in both the `Location` and
`X-Template-Id` headers. `X-Template-Preview-Status` is `ready` or `unavailable`, and
`X-PowerPoint-Warning-Count` reports non-fatal import warnings. Imported templates must contain
exactly one slide. The validated `kind` query is `diagram` or `commentary` and defaults to `diagram`.

Import a multi-slide deck as one independently stored template JSON per slide with
`POST /api/v1/batchImport?kind=diagram|commentary`. The request body and content type are the same as
`POST /api/v1/import`:

```sh
curl --request POST 'http://localhost:43127/api/v1/batchImport?kind=diagram' \
  --header 'X-App-Id: DiligenceStudio_WestMonroe' \
  --header 'Content-Type: application/vnd.openxmlformats-officedocument.presentationml.presentation' \
  --data-binary @deck.pptx
```

The `201` response contains `templates`, with one `templateId`, `templateJson`, and
`previewAvailable` entry for every source slide, plus an aggregate `warnings` array. Every
`templateJson` contains exactly one slide and is stored with its own metadata, image assets,
and preview. Its presentation title is the source slide's derived name. `X-Imported-Template-Count`
reports the number stored. The complete batch is committed atomically, so a database failure does
not leave a partially imported deck. The headless preview provider renders every slide from its
normalized JSON without writing presentation bytes to disk.

`GET /api/v1/templates/:templateId` returns that same canvas JSON body.
`GET /api/v1/import/:templateId` is a
backward-compatible alias. Both routes join the stored template to all of its assets and hydrate
each image `src` from its BLOB. Image bytes also remain directly available from
`/api/v1/import/:templateId/assets/:assetId`.

`GET /api/v1/templates?kind=diagram` lists lightweight metadata in newest-first order. Each item includes
its ID, kind, title, description, slide and element counts, and a nullable `previewUrl`.
`GET /api/v1/templates/previews?page=1` returns the actual PNG preview data for up to 10 templates in
newest-first order. Each item contains a browser-ready base64 `dataUrl`, the template ID, the
single-preview URL, content type, width, and height. The `pagination` object reports the current
page, fixed page size, total item and page counts, and whether adjacent pages exist. Templates
without a preview are excluded. The `page` query defaults to `1` and must be a positive integer.
`GET /api/v1/templates/:templateId/preview` returns the stored PNG without exposing local paths.
`DELETE /api/v1/templates/:templateId`
removes a template and its associated image assets, returning `204` when deleted and `404` when
the template does not exist.

Send the same canvas JSON structure to create a PowerPoint file:

```sh
curl --request POST http://localhost:43127/api/v1/export \
  --header 'X-App-Id: DiligenceStudio_WestMonroe' \
  --header 'Content-Type: application/json' \
  --data-binary @slide.json \
  --output generated-slide.pptx
```

The `200` response is a PowerPoint OOXML binary with an attachment filename derived from the
presentation title. A single-slide JSON input produces a one-slide deck; a multi-slide input
preserves all normalized slides. For server safety, image elements must contain embedded base64
`data:image/...` sources rather than filesystem paths or remote URLs.

Insert generated slides into an uploaded target deck with bounded multipart fields:

```sh
curl --request POST http://localhost:43127/api/v1/export/insert \
  --header 'X-App-Id: DiligenceStudio_WestMonroe' \
  --form 'presentation=<slide.json' \
  --form 'insertAfterSlide=1' \
  --form 'target=@target.pptx;type=application/vnd.openxmlformats-officedocument.presentationml.presentation' \
  --output target-with-slide.pptx
```

Errors use this shape:

```json
{
  "error": {
    "code": "invalid_powerpoint",
    "message": "The uploaded file is not a supported PowerPoint OOXML presentation.",
    "requestId": "..."
  }
}
```

SQLite runs only as an in-memory database and resets when the server process exits. It uses strict,
versioned schema migrations. `templates` stores `template_id TEXT PRIMARY KEY` and a
JSON-validated `template_json TEXT`. `template_assets` stores each image as an `asset_data BLOB`
under an `asset_id TEXT PRIMARY KEY`, with `template_id` as a foreign key back to `templates`.
`template_metadata` owns picker metadata and built-in checksums; `template_previews` owns PNG bytes
and dimensions. `apps` stores the unique app ID, display name, JSON metadata, and first/last-seen
timestamps; `app_templates` stores app-to-template access. Template, metadata, asset, preview, and
app-access inserts are committed in one transaction.
Schema version 5 stores separate subject and template-capability documents, fingerprints, and
vectors in `slide_classifications`, and rebuilds the FTS5 index around the v2 metadata sections.
Migration deletes pre-v2 classification rows and their search index entries while preserving the
underlying templates and assets; those templates require explicit reclassification. Vectors are
validated little-endian Float32 BLOBs. The pilot scores app-scoped candidate vectors in process; measure query
latency and evaluate a supported vector extension or dedicated store before using this design for a
large catalog. Re-evaluate the in-process scan before an app exceeds 5,000 compatible vectors or
when measured retrieval p95 exceeds 200 ms, whichever comes first.
SQLite journals and temporary tables remain in memory. Existing templates that still contain
embedded base64 images or nonpositive canvas
dimensions are migrated transactionally when they are first retrieved.
