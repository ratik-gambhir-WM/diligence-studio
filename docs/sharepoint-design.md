# SharePoint integration design

This document describes the SharePoint file-ingestion integration implemented by Diligence Studio.
It is intended for developers configuring, reviewing, troubleshooting, or extending the integration.

The implementation supports two related capabilities:

1. Microsoft Entra sign-in establishes a server-side session and gives the server access to a Microsoft
   Graph access token.
2. A signed-in user can paste an approved SharePoint file or folder link, select supported files, and
   add the downloaded files to the existing browser-side diagram-generation workflow.

The integration keeps Microsoft-specific behavior behind the `MICROSOFT_SUPPORT` feature flag. The
web and server environments should use the same value: the web flag controls rendering, while the
server flag controls SharePoint route registration and versioned API authentication. The SharePoint
API always requires a valid Microsoft session when enabled.

## Scope and non-goals

### In scope

- Resolving a SharePoint sharing link through Microsoft Graph.
- Resolving either one file or a folder tree.
- Filtering folder contents to supported diligence file types.
- Selecting files from a resolved folder.
- Downloading selected file bytes through the Express API.
- Reusing the existing OpenAI attachment pipeline.
- Hiding the SharePoint controls and short-circuiting their handlers when Microsoft support is disabled.
- Validating external URLs, Graph responses, identifiers, response headers, and client response shapes.

### Out of scope

- Replacing the existing in-memory Microsoft session store.
- Persisting SharePoint files or metadata in SQLite.
- Adding SharePoint authorization independent of Microsoft Graph permissions.
- Passing SharePoint URLs directly to OpenAI.
- Exposing Microsoft access or refresh tokens to browser JavaScript.

When Microsoft support is enabled, the server endpoints are available only to authenticated clients.
When it is disabled, the server does not register the SharePoint routes and leaves the existing
anonymous demo API behavior in place.

## Architecture

The integration preserves the existing runtime boundary:

| Layer | Responsibility | Main implementation |
| --- | --- | --- |
| Vite/browser configuration | Reads the frontend feature flag at build time | web/vite.config.ts |
| Browser feature helper | Strictly parses the exposed boolean | web/src/lib/microsoftSupport.ts |
| Top-level workflow | Controls feature visibility, resolution, downloads, and attachment handoff | web/src/App.tsx |
| SharePoint UI | Accepts a link and lets users select folder files | web/src/pages/PromptPage.tsx, web/src/components/SharePointFilePicker.tsx |
| Browser API client | Calls the versioned SharePoint endpoints and validates responses | web/src/lib/api/sharepointApi.ts |
| Attachment session | Stores browser File objects and their source metadata | web/src/hooks/useDiagramSession.ts |
| Express routes | Authenticates requests, validates inputs, and maps service errors | server/src/routes/sharePointRoutes.ts |
| Graph integration | Resolves links, traverses folders, and downloads bytes | server/src/services/SharePointResourceService.ts |
| Microsoft session | Exchanges OAuth codes and refreshes Graph access tokens | server/src/auth/microsoftAuth.ts |

The end-to-end flow is:

    Browser link input
      -> POST /api/v1/sharepoint/resolve
      -> Express session lookup
      -> Microsoft Graph /shares/{encoded-sharing-url}/driveItem
      -> File metadata returned to browser
      -> Optional folder selection in the browser
      -> GET /api/v1/sharepoint/files/{driveId}/{fileId}/content
      -> Express session lookup
      -> Microsoft Graph drive item metadata and content
      -> Browser File object
      -> useDiagramSession attachment record
      -> Existing model-selection and slide-generation OpenAI calls

The Graph access token and Graph URLs remain server-side. The browser receives only sanitized
resource metadata and the downloaded file bytes.

## Feature flag and configuration

The flag is configured in both the server `.env` and `web/.env`:

    MICROSOFT_SUPPORT=false

Vite loads the web workspace environment and converts the value to a public boolean:

- web/vite.config.ts reads MICROSOFT_SUPPORT.
- server/src/config.ts reads MICROSOFT_SUPPORT for route registration and API authentication.
- Only a case-insensitive, whitespace-trimmed value of true enables the feature.
- The Vite definition is exposed as VITE_MICROSOFT_SUPPORT.
- web/src/lib/microsoftSupport.ts applies the same strict true parsing in the browser.
- Unset, empty, false, and any other value disable the feature.

The web value is build-time configuration and the server value is read at startup. Restart/rebuild
the corresponding process after changing either value.

When MICROSOFT_SUPPORT is false:

- App.tsx does not request /api/auth/me.
- The /login route redirects to the anonymous demo app.
- SharePoint URL, loading, picker, and error controls are not rendered.
- SharePoint resolve, file-selection, and download handlers return before making requests.
- The server does not register the SharePoint routes, and `/api/v1` retains anonymous demo behavior.

When MICROSOFT_SUPPORT is true:

- App.tsx performs the Microsoft session bootstrap.
- Protected routes redirect unauthenticated users to /login.
- The SharePoint controls are rendered.
- SharePoint handlers may call the browser API client.
- The server registers SharePoint routes and requires the Microsoft session for `/api/v1`.

The root .env.example contains server-side Microsoft and SharePoint settings. It is separate from
web/.env.example, which contains browser-facing Vite settings. Microsoft client secrets must remain
in the root server environment and must never be placed in a VITE_* variable.

## Microsoft session prerequisite

SharePoint calls use the access token associated with the current Microsoft session. The session
flow is documented in docs/microsoft-sign-in-flow-design.md; the relevant contract is:

1. The browser starts the Entra authorization-code flow with PKCE.
2. The Express callback exchanges the code and calls Microsoft Graph /me.
3. The server stores the access token, expiry, optional refresh token, and normalized user in memory.
4. The browser receives only an HttpOnly session cookie.
5. getMicrosoftAccessToken() retrieves or refreshes the token for SharePoint requests.

The authorization request includes:

    openid profile email offline_access User.Read Files.Read Sites.Read.All

Files.Read is used for file content access. Sites.Read.All supports access to SharePoint site
content in the configured tenant. Tenant administrators may need to grant consent.

## Express API

The routes are mounted under /api/v1/sharepoint by server/src/app.ts when a
SharePointResourceServiceLike dependency is provided. The production server always constructs the
service with the configured allowlist.

### Resolve a file or folder

    POST /api/v1/sharepoint/resolve
    Content-Type: application/json
    Cookie: diligence_studio_session=...

Request body:

    { "url": "https://relentlessblue.sharepoint.com/:f:/s/Site/folder?e=..." }

The route:

1. Requires application/json with a small 16 KiB body limit.
2. Requires a valid Microsoft access token.
3. Passes the trimmed URL to SharePointResourceService.
4. Returns a private, non-cacheable JSON response.

Successful response shape:

    {
      "files": [
        {
          "driveId": "drive-id",
          "fileId": "file-id",
          "lastModifiedDateTime": "2026-09-23T00:00:00Z",
          "mimeType": "application/pdf",
          "name": "diagram.pdf",
          "path": "Architecture",
          "size": 12345,
          "webUrl": "https://relentlessblue.sharepoint.com/..."
        }
      ],
      "kind": "folder",
      "name": "Architecture",
      "resourceUrl": "https://relentlessblue.sharepoint.com/...",
      "webUrl": "https://relentlessblue.sharepoint.com/..."
    }

For a single file, kind is file and files contains one record. For a folder, kind is folder and
files contains all supported files discovered below the folder.

### Download a file

    GET /api/v1/sharepoint/files/{driveId}/{fileId}/content
    Cookie: diligence_studio_session=...

The route validates both path identifiers, requires a Microsoft access token, and returns the
downloaded bytes with:

- Content-Type from the Graph metadata or response fallback.
- Content-Disposition with a sanitized attachment filename.
- Content-Length.
- Cache-Control: private, no-store.
- X-Content-Type-Options: nosniff.

Downloads are limited to the server's configured upload limit (25 MiB by default) and are read with
a bounded stream. The Graph content URL is never sent to the browser. The browser calls the
application endpoint and receives a File constructed from the response bytes.

### Error envelope

Authentication errors become authentication_required. SharePointResourceError values preserve their
sanitized status and code through the common Express ApiError handler. Provider response bodies and
tokens are not returned to the browser or logged.

Common error codes include:

| Code | Meaning |
| --- | --- |
| invalid_sharepoint_request | Resolve request did not contain a usable URL |
| invalid_sharepoint_url | URL could not be parsed |
| unsupported_sharepoint_host | Host is not in the configured allowlist |
| sharepoint_access_denied | Graph returned 401 or 403 |
| sharepoint_unsupported_file_type | A shared file has an unsupported extension |
| sharepoint_missing_drive | Graph item did not identify a document-library drive |
| sharepoint_file_limit_exceeded | Folder traversal exceeded the supported-file limit |
| sharepoint_folder_item_limit_exceeded | Folder traversal exceeded the total-item limit |
| sharepoint_folder_request_limit_exceeded | Folder traversal exceeded the Graph request limit |
| sharepoint_file_too_large | A downloaded file exceeded the byte limit |
| sharepoint_folder_too_deep | Folder traversal exceeded the nesting limit |
| invalid_sharepoint_response | Graph returned a shape the service could not safely interpret |

## Graph resource resolution

SharePointResourceService owns all Graph-specific behavior.

### Host validation

The resource URL must:

- Use HTTPS.
- Have a hostname in SHAREPOINT_ALLOWED_HOSTS.
- Match a hostname allowlist entry case-insensitively.

The default allowlist is:

    relentlessblue.sharepoint.com

The list is parsed as comma-separated hostnames by server/src/config.ts. Values are normalized,
deduplicated, and rejected if they contain characters outside the hostname grammar.

This validation prevents the application from treating arbitrary URLs as Microsoft Graph resources.
The service does not follow a user-provided URL directly.

### Shared-link encoding

Graph's shares endpoint requires a sharing URL encoded as:

1. UTF-8 bytes.
2. Base64.
3. Padding removed.
4. Slash replaced with underscore.
5. Plus replaced with hyphen.
6. The u! prefix added.

The service requests:

    /shares/{encoded-sharing-url}/driveItem?$select=id,name,size,file,folder,parentReference,webUrl,lastModifiedDateTime

The response is narrowed from unknown before use. A resource must have a usable id, name, webUrl,
and parent drive reference.

### Single-file resources

A resource with a file facet is accepted only when its filename extension is supported. The current
supported set is:

    docx, pdf, ppt, pptx, png, jpg, jpeg, md, markdown, txt, rtf

The result contains one SharePointFile record with the drive ID, item ID, filename, MIME type, size,
modified timestamp, web URL, and an empty root path.

### Folder resources

Folder traversal uses the drive item's children endpoint and follows Graph @odata.nextLink values.
The implementation:

- Requests up to 200 children per page.
- Recurses through nested folders.
- Retains only supported file extensions.
- Builds a display path such as Architecture/Archive.
- Stops with an error after 500 supported files, 5,000 total items, or 1,000 Graph list requests.
- Stops with an error beyond depth 8.
- Accepts only next links that remain under the Microsoft Graph v1.0 base URL.

Unsupported files are omitted from a folder result rather than causing the entire folder to fail.
If no supported files remain, the browser presents an empty picker state.

### Download behavior

Before downloading content, the service fetches Graph metadata again. This confirms that the selected
item is still a file and supplies a reliable filename and MIME type. It then requests:

    /drives/{driveId}/items/{fileId}/content

Drive and item identifiers are validated for length and control characters, then URL-encoded before
being included in the Graph path. Graph access tokens are normalized to avoid duplicating a Bearer
prefix.

Non-success Graph responses are converted to sanitized application errors. The provider response is
consumed only for internal error handling and is not returned or written to application logs.

## Browser workflow

### Link resolution

PromptPage owns the visible SharePoint controls, while App.tsx owns the async workflow state:

- sharePointUrl
- sharePointResource
- sharePointError
- isSharePointResolving
- isSharePointDownloading

The user enters a link and selects Load SharePoint. The browser client posts the URL with
credentials included. A single-file result starts downloading immediately. A folder result opens
SharePointFilePicker.

### Folder selection

SharePointFilePicker:

- Resets selection whenever a new resource is loaded.
- Tracks selected file IDs in a Set.
- Supports Select all and Clear all.
- Displays the relative path and formatted size.
- Disables controls during downloads.
- Requires at least one selection before confirming.

The picker returns metadata records to App.tsx. It does not download bytes itself.

### Bounded downloads

App.tsx downloads selected files with mapWithConcurrency and a concurrency of four. Each result is
converted to:

    {
      file: File,
      fileId: string,
      path: string,
      sourceUrl: string
    }

The files are then handed to useDiagramSession.addSharePointAttachments().

If any download in the bounded batch fails, the operation reports the sanitized error to the UI and
does not add the completed batch through the normal success path.

### Attachment session policy

useDiagramSession stores an AttachmentRecord containing:

- A browser File object.
- A stable local record ID.
- The attachment mode: template-context or upload-only.
- The source: upload or sharepoint.
- SharePoint identifiers and display metadata when applicable.

SharePoint files are added as upload-only attachments so they participate in the same submit flow as
files selected through the local upload control.

The current policy treats local upload-only files and SharePoint upload-only files as mutually
exclusive sources. Adding one source replaces the existing upload-only source while preserving
template-context attachments. This avoids ambiguous mixed-source behavior in the single upload-only
submit control.

### OpenAI handoff

The downloaded SharePoint File objects are not sent as SharePoint links. They follow the existing
browser attachment path:

1. App.tsx snapshots the current upload-only File array at submit time.
2. Template mode sends the snapshot to model selection and then to template text generation.
3. Create mode sends the snapshot to diagram generation.
4. OpenAI.ts reads each File with arrayBuffer().
5. Non-image files become input_file content with a data URL.
6. PNG/JPG/JPEG files become input_image content with a data URL.

The frontend OpenAI adapter is intentionally a development/demo trust model. VITE_* configuration
is public browser code, and the OpenAI key is not a server secret boundary.

## Client response validation

web/src/lib/api/sharepointApi.ts treats browser responses as untrusted:

- It checks HTTP success before consuming the result.
- It parses sanitized JSON error codes and messages.
- It requires JSON content for resolve responses.
- It validates resource kind, names, URLs, and file arrays.
- It validates file IDs, paths, MIME types, timestamps, URLs, sizes, and non-negative integer sizes.
- It constructs a browser File only after a successful download response.

The client does not trust arbitrary server JSON merely because it has a TypeScript type.

## Security and privacy controls

### Token boundary

Microsoft access and refresh tokens remain in the server's in-memory session map. The browser holds
only the HttpOnly session cookie. Graph calls are made by Express.

### Link and path boundary

SharePoint hosts are allowlisted. Graph IDs are validated and encoded. Folder pagination links must
remain under the Graph v1.0 origin. Download filenames are sanitized before entering a response
header.

### Cache and response boundary

SharePoint metadata and downloaded bytes are marked private and no-store. Download responses include
nosniff. API errors use sanitized application messages rather than provider response bodies.

### UI feature boundary

MICROSOFT_SUPPORT controls browser behavior and server route registration. Defensive checks remain
in the App.tsx handlers so an accidentally retained UI callback cannot start SharePoint work while
the flag is false.

### Logging

The SharePoint service does not log tokens, file bytes, Graph response bodies, or document contents.
The browser attachment diagnostics are development-only and report counts, byte totals, file type
categories, request operation, serialized payload size, and empty payload count. They intentionally
omit filenames, file contents, prompts, keys, and generated diagram content.

Useful browser events include:

- diligence-studio:file-selection
- diligence-studio:attachment-state
- diligence-studio:upload-only-submit-started
- diligence-studio:openai-request-started
- diligence-studio:openai-request-attachments-built
- diligence-studio:openai-response-received

For the template-text-update event, an uploaded or SharePoint file is confirmed at the OpenAI
boundary when expectedAttachmentCount and attachmentContentCount include the same file count,
serializedAttachmentPayloadChars is greater than zero, and emptySerializedAttachmentCount is zero.
The model-selection event also includes generated candidate preview images, so its expected count is
the sum of the user attachments and those candidate images.

## Configuration reference

### Server configuration

| Variable | Default | Purpose |
| --- | --- | --- |
| MICROSOFT_TENANT_ID | organizations | Entra tenant or tenant alias |
| MICROSOFT_CLIENT_ID | none | Confidential client ID |
| MICROSOFT_CLIENT_SECRET | none | Confidential client secret value |
| MICROSOFT_REDIRECT_URI | http://localhost:43127/api/auth/callback | Entra callback |
| MICROSOFT_FRONTEND_ORIGIN | http://localhost:5173 | Post-login frontend origin |
| MICROSOFT_SUPPORT | false | Enables the web feature and server-side API/SharePoint protection |
| MICROSOFT_COOKIE_SECURE | true | Adds Secure to the session cookie; set false only for local HTTP development |
| SHAREPOINT_ALLOWED_HOSTS | relentlessblue.sharepoint.com | Comma-separated approved SharePoint hosts |

These values belong in the root server environment.

### Browser configuration

| Variable | Default | Purpose |
| --- | --- | --- |
| MICROSOFT_SUPPORT | false | Frontend feature flag; keep aligned with the server value |
| VITE_API_BASE_URL | /api/v1 | Browser API prefix |
| VITE_API_PROXY_TARGET | http://127.0.0.1:43127 | Local Vite proxy target |

These values belong in web/.env. The browser-facing flag is exposed to code as
VITE_MICROSOFT_SUPPORT by Vite.

## Testing strategy

The implementation has tests at each boundary:

| Area | Tests |
| --- | --- |
| Microsoft auth flow | server/test/microsoftAuth.test.ts |
| Graph resource service | server/test/sharePointResourceService.test.ts |
| Express SharePoint routes | server/test/sharePointRoutes.test.ts |
| Browser SharePoint API client | web/src/lib/api/sharepointApi.test.ts |
| Feature flag parsing | web/src/lib/microsoftSupport.test.ts |
| SharePoint control visibility | web/src/pages/PromptPage.test.tsx |
| Anonymous Microsoft-disabled routing | web/src/App.test.tsx |
| File identity through the attachment hook | web/src/hooks/useDiagramSession.test.tsx |

The service tests cover shared-file resolution, nested folders, downloads, and host rejection.
Route tests cover authenticated resolve and download behavior. Client tests cover response validation,
filename handling, and sanitized errors. UI tests verify that SharePoint controls and alerts are
hidden when the feature is disabled and present when enabled.

## Known limitations and extension points

- Sessions are process-local and disappear on server restart.
- The web and server flags must be changed together; restarting/rebuilding is required after changing
  either environment.
- The browser OpenAI adapter sends file data directly from the client under the existing demo trust
  model.
- Folder selection downloads the chosen files into browser memory before model submission.
- The current attachment policy replaces one upload-only source with the other.
- SharePoint resource metadata is not persisted for later reuse.
- The Graph permission scope and host allowlist are intentionally broad enough for the configured
  demo tenant but should be reviewed for a production tenant.

Potential future work includes durable session/token storage, server-side document staging, richer
audit logging without document content, and a first-class mixed local/SharePoint attachment model.
