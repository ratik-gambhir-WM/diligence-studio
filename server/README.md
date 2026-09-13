# PowerPoint import and export APIs

This package is the Node-only sibling of the React app in `../web`. It keeps its PowerPoint
implementation under `src/lib`; nothing in this package enters the React/Vite module graph.
Uploaded `.pptx` files are OOXML ZIP packages. Their
compact canvas JSON is stored in SQLite. Imported image bytes are stored separately from the JSON
as BLOB assets. Template metadata and first-slide PNG previews are stored in dedicated tables.
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
- `SQLITE_DB_PATH` defaults to `server/data/templates.sqlite` from the repository root.
- `MAX_PPTX_UPLOAD_BYTES` defaults to `26214400` (25 MiB).
- `MAX_EXPORT_JSON_BYTES` defaults to `52428800` (50 MiB).
- `TEMPLATE_PREVIEW_PROVIDER` defaults to `headless`. It accepts `headless`, `quicklook`, or
  `disabled`.
- `TEMPLATE_PREVIEW_RENDER_SIZE` defaults to `1600` pixels.
- `TEMPLATE_PREVIEW_RENDER_URL` defaults to
  `http://localhost:5173/_internal/template-preview`. It must point to the Vite app (or the deployed
  web app) while imports are running.
- `TEMPLATE_PREVIEW_TIMEOUT_MS` defaults to `15000`.
- `MAX_TEMPLATE_PREVIEW_BYTES` defaults to `10485760` (10 MiB).

Requests time out after 30 seconds. Compressed request bodies are rejected. Template, import, and
export requests accept an `X-App-Id` header (or `appId` query parameter) and fall back to
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
at `TEMPLATE_PREVIEW_RENDER_URL`. `quicklook` remains available as an explicit compatibility
fallback on macOS, but it renders the uploaded PowerPoint rather than the normalized canvas model.

## API

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
`templateJson` contains exactly one slide and is persisted with its own metadata, image assets,
and preview. Its presentation title is the source slide's derived name. `X-Imported-Template-Count`
reports the number stored. The complete batch is committed atomically, so a database failure does
not leave a partially imported deck. The default headless preview provider renders every slide;
the Quick Look compatibility provider can preview only the first slide and reports later previews
as unavailable rather than storing an incorrect image.

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

SQLite uses strict, versioned schema migrations. `templates` stores `template_id TEXT PRIMARY KEY` and a
JSON-validated `template_json TEXT`. `template_assets` stores each image as an `asset_data BLOB`
under an `asset_id TEXT PRIMARY KEY`, with `template_id` as a foreign key back to `templates`.
`template_metadata` owns picker metadata and built-in checksums; `template_previews` owns PNG bytes
and dimensions. `apps` stores the unique app ID, display name, JSON metadata, and first/last-seen
timestamps; `app_templates` stores app-to-template access. Template, metadata, asset, preview, and
app-access inserts are committed in one transaction.
File-backed databases use SQLite WAL
mode. Existing templates that still contain embedded base64 images or nonpositive canvas
dimensions are migrated transactionally when they are first retrieved.
