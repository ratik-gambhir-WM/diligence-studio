# diligence-studio API

The diligence-studio server exposes a versioned HTTP API for importing PowerPoint templates,
browsing stored templates and previews, exporting canvas JSON to PowerPoint, and inserting
generated slides into an existing deck.

The existing API is served beneath `/api/v1`; asynchronous classified import is the single v2
endpoint. The local development server listens on
`http://127.0.0.1:43127` by default. The Vite development server proxies `/api/v1` to that
address.

There is no authentication middleware in the Express application. Production deployments must
put the API behind the application's authentication and network boundary before exposing it to
untrusted callers.

## Request conventions

- JSON requests use `Content-Type: application/json`.
- PowerPoint uploads use
  `application/vnd.openxmlformats-officedocument.presentationml.presentation`.
- Compressed request bodies are rejected. Send `Content-Encoding: identity` or omit the header.
- Every response includes an `X-Request-Id` header. JSON errors repeat that value as
  `error.requestId`.
- Unless noted otherwise, successful JSON responses use `Content-Type: application/json`.
- Template kinds are `diagram` and `commentary`.
- App-scoped routes accept `X-App-Id` or an `appId` query parameter. IDs use 1-128 letters,
  numbers, underscores, or hyphens. The legacy fallback is `DiligenceStudio_WestMonroe`.
- A valid app ID is registered on first use. It scopes template lists, retrieval, previews, assets,
  deletion, imports, and stored-asset resolution during export. This is namespacing, not caller
  authentication; a trusted gateway must prevent callers from claiming another app's ID.

The default limits are a 25 MiB PowerPoint upload, a 50 MiB JSON export body, a 10 MiB stored
preview, and a 30-second request timeout. These can be changed with the server configuration
documented in [`server/README.md`](../../server/README.md).

## Endpoints

| Method | Path | Purpose |
| --- | --- | --- |
| `POST` | `/api/v2/import?kind=...` | Import one slide and atomically create pending retrieval work. |
| `POST` | `/api/v1/import?kind=...` | Import a single-slide PowerPoint as one template. |
| `POST` | `/api/v1/batchImport?kind=...` | Import every slide as a separate template. |
| `GET` | `/api/v1/templates?kind=...` | List template metadata. |
| `GET` | `/api/v1/templates/previews?page=...` | List stored PNG previews in pages of 10. |
| `GET` | `/api/v1/templates/:templateId` | Retrieve a template's canvas JSON. |
| `GET` | `/api/v1/import/:templateId` | Backward-compatible alias for template retrieval. |
| `GET` | `/api/v1/templates/:templateId/preview` | Retrieve a template's stored PNG preview. |
| `GET` | `/api/v1/import/:templateId/assets/:assetId` | Retrieve one stored template image. |
| `DELETE` | `/api/v1/templates/:templateId` | Delete a template and its associated assets and preview. |
| `POST` | `/api/v1/export` | Convert canvas JSON to a PowerPoint file. |
| `POST` | `/api/v1/export/insert` | Insert generated slides into an existing PowerPoint file. |

### Import one slide

`POST /api/v1/import?kind=diagram`

Send the `.pptx` bytes as the request body. `kind` is optional and defaults to `diagram`. The
single-slide endpoint rejects decks containing anything other than exactly one slide.

```sh
curl --request POST 'http://localhost:43127/api/v1/import?kind=diagram' \
  --header 'X-App-Id: DiligenceStudio_WestMonroe' \
  --header 'Content-Type: application/vnd.openxmlformats-officedocument.presentationml.presentation' \
  --data-binary @deck.pptx
```

The response is `201 Created` with a canvas JSON document. Imported image sources are hydrated as
embedded `data:image/...;base64,...` URIs in the response.

Response headers:

- `Location: /api/v1/templates/:templateId?appId=:appId`
- `X-Template-Id: :templateId`
- `X-Template-Preview-Status: ready|unavailable`
- `X-PowerPoint-Warning-Count: <number>`
- `Link: </api/v1/templates/:templateId/preview?appId=:appId>; rel="preview"` when a preview is available

### Import every slide

`POST /api/v1/batchImport?kind=diagram`

The request format is the same as the single-slide import. Each source slide is persisted as an
independent one-slide template. The operation is atomic: if persistence fails, no templates from
the batch remain stored.

The response is `201 Created`:

```json
{
  "templates": [
    {
      "templateId": "8f0d...",
      "previewAvailable": true,
      "templateJson": {
        "presentation": {
          "title": "Architecture",
          "preserveElementOrder": true,
          "showBranding": false,
          "slides": []
        }
      }
    }
  ],
  "warnings": []
}
```

`X-Imported-Template-Count` and `X-PowerPoint-Warning-Count` report the aggregate result.

### Import one slide for retrieval indexing

`POST /api/v2/import?kind=diagram` uses the same content type, app scope, upload limit, compressed
body rejection, and exactly-one-slide rule as v1. Its `201 Created` body is an envelope:

```json
{
  "previewAvailable": true,
  "templateId": "8f0d...",
  "templateJson": { "presentation": {} },
  "warnings": [],
  "retrieval": { "status": "pending" }
}
```

`retrieval.status` confirms durable work was committed; it does not mean classification has begun.
The endpoint returns before classification and embedding. No public v2 read, query, status, retry,
classification, embedding, batch-import, or backfill route exists. Existing v1 imports never create
this work.

### List templates

`GET /api/v1/templates?kind=diagram`

`kind` is optional. Without it, both kinds are returned. Results are newest first.

```json
{
  "templates": [
    {
      "templateId": "8f0d...",
      "kind": "diagram",
      "title": "Architecture",
      "description": "Imported PowerPoint template",
      "slideCount": 1,
      "elementCount": 12,
      "previewUrl": "/templates/8f0d.../preview?appId=DiligenceStudio_WestMonroe"
    }
  ]
}
```

`previewUrl` is `null` when no preview was generated. It is an API-relative path; clients using a
non-default API base should resolve it against that base.

### List previews

`GET /api/v1/templates/previews?page=1`

The page defaults to `1`, must be a positive integer, and contains at most 10 previews. Templates
without a stored preview are excluded.

```json
{
  "pagination": {
    "page": 1,
    "pageSize": 10,
    "totalItems": 1,
    "totalPages": 1,
    "hasPreviousPage": false,
    "hasNextPage": false
  },
  "previews": [
    {
      "templateId": "8f0d...",
      "previewUrl": "/templates/8f0d.../preview?appId=DiligenceStudio_WestMonroe",
      "contentType": "image/png",
      "width": 1600,
      "height": 900,
      "dataUrl": "data:image/png;base64,..."
    }
  ]
}
```

### Retrieve, preview, and delete a template

`GET /api/v1/templates/:templateId` returns the canvas JSON document described in
[`canvas-json.md`](canvas-json.md). The legacy `GET /api/v1/import/:templateId` route returns the
same response. Image elements in the retrieved document contain embedded data URIs; the direct
asset route remains available when a binary response is preferred.

`GET /api/v1/templates/:templateId/preview` returns `image/png` bytes. The response is private and
not cacheable by shared clients. `GET /api/v1/import/:templateId/assets/:assetId` returns the stored
image bytes with the asset's content type.

`DELETE /api/v1/templates/:templateId` returns `204 No Content` when deletion succeeds and `404`
when the template is unavailable to the app. Deletion removes that app's access; the template and
its assets are removed when no app can access it.

### Export a PowerPoint

`POST /api/v1/export`

Send a canvas JSON document as the body:

```sh
curl --request POST 'http://localhost:43127/api/v1/export' \
  --header 'X-App-Id: DiligenceStudio_WestMonroe' \
  --header 'Content-Type: application/json' \
  --data-binary @slide.json \
  --output generated-slide.pptx
```

The response is `200 OK` with a PowerPoint OOXML binary, an attachment filename derived from the
presentation title, and `X-PowerPoint-Warning-Count`. The input may contain one or more slides.
Image elements must use embedded base64 data URIs. Stored asset paths are resolved when the
server can identify the referenced template asset.

### Insert into an existing deck

`POST /api/v1/export/insert`

Send `multipart/form-data` with exactly these parts:

| Part | Type | Required | Description |
| --- | --- | --- | --- |
| `presentation` | text | Yes | Canvas JSON to convert. |
| `insertAfterSlide` | text | Yes | Non-negative whole number. `0` inserts before the first slide. |
| `target` | file | Yes | Existing `.pptx` with the PowerPoint MIME type. |

```sh
curl --request POST 'http://localhost:43127/api/v1/export/insert' \
  --header 'X-App-Id: DiligenceStudio_WestMonroe' \
  --form 'presentation=<slide.json' \
  --form 'insertAfterSlide=1' \
  --form 'target=@target.pptx;type=application/vnd.openxmlformats-officedocument.presentationml.presentation' \
  --output target-with-slide.pptx
```

The response has the same PowerPoint content type, download headers, and warning-count header as
`POST /api/v1/export`.

## Errors

Errors are sanitized JSON and use the same shape for validation, missing resources, and server
failures:

```json
{
  "error": {
    "code": "invalid_powerpoint",
    "message": "The uploaded file is not a supported PowerPoint OOXML presentation.",
    "requestId": "f3a1..."
  }
}
```

Common status and code combinations include:

| Status | Codes | Typical cause |
| --- | --- | --- |
| `400` | `invalid_app_id`, `conflicting_app_id`, `invalid_json`, `invalid_template_kind`, `invalid_preview_page`, `invalid_insert_position`, `missing_multipart_field` | Malformed or incomplete request. |
| `404` | `route_not_found`, `template_not_found`, `template_asset_not_found`, `template_preview_not_found` | Resource or route does not exist. |
| `413` | `payload_too_large`, `embedded_images_too_large` | Configured request or processing limit exceeded. |
| `415` | `unsupported_media_type`, `unsupported_target_media_type`, `unsupported_content_encoding` | Unsupported content type or compressed body. |
| `422` | `invalid_powerpoint`, `template_must_have_one_slide`, `powerpoint_has_no_slides`, `invalid_presentation_json`, `unsupported_image_source`, `invalid_target_powerpoint` | Input has the right transport format but cannot be processed. |
| `499` | `request_cancelled` | The client disconnected or the request timed out during import. |
| `500` | `internal_error` | Unexpected server failure. |

Clients should use `error.code` for handling and display `error.message` as sanitized user-facing
diagnostic text. The full implementation is in [`server/src/errors.ts`](../../server/src/errors.ts)
and the endpoint wiring is in [`server/src/app.ts`](../../server/src/app.ts).

## Related references

- [Canvas JSON contract](canvas-json.md)
- [Server setup and configuration](../../server/README.md)
- [Browser API client](../../web/src/lib/api/templateApi.ts)
