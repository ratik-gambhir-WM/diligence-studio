# Simple Synchronous Slide Classification Plan

## Goal

Replace the asynchronous classification workflow for `POST /api/v2/import` with one synchronous
request flow:

1. Validate and convert the uploaded single-slide PowerPoint.
2. Save the template, assets, preview, and a classification record in SQLite.
3. Classify the saved slide with OpenAI while the HTTP request remains open.
4. Validate and save the returned classification metadata.
5. Return the template and completed classification to the caller.

There is no background worker, in-memory queue, database job queue, polling loop, lease, scheduled
retry, or separate embedding stage in this plan.

## Core decision

Classification is part of the import request rather than eventual background work. A successful
response means both of these facts are true:

- the template is stored and available through the existing template APIs; and
- validated classification metadata is stored for that template.

The server must not return `201 Created` with a pending classification. It waits for the provider
call and returns only after classification persistence succeeds.

This is intentionally optimized for a simple, low-volume first release. Each concurrent import
can create one concurrent OpenAI request. If traffic, latency, or provider rate limits later make
that unsuitable, background jobs can be reconsidered from measured evidence.

## Scope

This plan includes:

- the existing single-slide PowerPoint validation and conversion flow;
- template, asset, preview, and app-scope persistence;
- bounded classification input construction;
- one synchronous OpenAI Responses API classification call;
- strict runtime validation and normalization of `SlideRetrievalMetadata`;
- classification persistence and a completed import response;
- request cancellation, provider timeout, and sanitized failure behavior; and
- focused service, repository, and HTTP contract tests.

This plan does not include:

- `SlideClassificationWorker` or any other long-lived worker;
- classification or embedding job claims;
- polling, notification callbacks, leases, backoff, jitter, or automatic retries;
- embeddings, semantic search, hybrid ranking, or vector storage;
- OpenAI Agents SDK tools;
- public classification status, retry, backfill, or query endpoints;
- batch classification; or
- changes to existing v1 import behavior.

## Current repository findings

- `ImportService.importV2` currently saves durable pending work and calls a best-effort worker
  notification callback.
- `SlideClassificationWorker` currently claims classification and embedding jobs, manages retry
  timing, and drains work during shutdown.
- `SlideClassificationService` currently accepts claimed jobs, calls the provider, and persists
  classification or embedding results.
- `SqliteTemplateRepository` already stores templates, previews, assets, classification metadata,
  facets, and provider provenance.
- `createImportHandlers` already keeps an `AbortSignal` tied to the request and response lifetime.
- The current default HTTP timeout is 30 seconds, while classification alone may be configured for
  45 seconds. A synchronous endpoint therefore requires an explicit compatible timeout budget.
- Existing v1 import and batch import must remain classification-free.

## Target architecture

```mermaid
flowchart LR
    REQUEST["POST /api/v2/import"] --> VALIDATE["Validate and convert one slide"]
    VALIDATE --> SAVE["Save template and classification state"]
    SAVE --> INPUT["Build bounded classification input"]
    INPUT --> OPENAI["Call OpenAI and wait"]
    OPENAI --> PARSE["Validate and normalize metadata"]
    PARSE --> COMPLETE["Save ready classification"]
    COMPLETE --> RESPONSE["Return template plus classification"]
```

The important transaction boundary is:

```text
SQLite transaction 1: save template + assets + preview + processing classification row
External call:        classify the saved template
SQLite transaction 2: save validated metadata + mark classification ready
HTTP response:        return 201 with template and classification
```

Do not hold a SQLite transaction open while waiting for OpenAI. A database transaction cannot make
an external provider call atomic, and holding the write lock would block unrelated repository
work.

## Request lifecycle

### 1. Reject unusable requests before persistence

Before writing a template:

- require the PowerPoint content type and enforce the existing upload byte limit;
- require the classification provider and its server-side configuration to be available;
- convert the PowerPoint with the existing bounded importer;
- require exactly one slide;
- generate the preview when configured;
- normalize the canvas and externalize image assets; and
- prepare the canonical template record that will be saved.

Provider configuration can be checked before conversion because an unconfigured server cannot
complete this endpoint. No OpenAI call occurs during validation or conversion.

### 2. Save before classification

In one SQLite transaction:

- ensure the app scope exists;
- insert the template, metadata, assets, preview, and app association; and
- insert a classification row with `status = 'processing'`, its input fingerprint, schema version,
  prompt version, and no classification metadata.

The commit must finish before the OpenAI request starts. This satisfies the required ordering:
the template exists in the database before it is classified.

### 3. Classify synchronously

After the save commits, build the bounded classification input from the canonical template and
preview, then call the injected `SlideClassifier` directly from the import use case. Forward the
request's `AbortSignal` and enforce a provider timeout shorter than the endpoint's total timeout.

The classifier receives only:

- a bounded deterministic slide digest; and
- the validated PNG preview when it is within the classification image limit.

It must not receive raw OOXML, asset bytes, base64 image sources, database data, absolute paths,
request headers, app metadata, or secrets.

There is one provider attempt per import request. Do not add SDK retries or application retries in
this synchronous version because they increase latency and can multiply billable calls.

### 4. Validate and persist the result

Parse the provider result with the existing strict `SlideRetrievalMetadataSchema`, then normalize it
deterministically. A TypeScript cast is not sufficient validation.

In a second short SQLite transaction:

- set `status = 'ready'`;
- store normalized metadata JSON and queryable facet columns;
- store the input fingerprint, schema version, prompt version, model, and completion timestamp;
- clear any prior sanitized error code.

This plan does not create an FTS row, embedding document, or vector.

### 5. Return the completed result

Return `201 Created` only after the ready classification record is committed. Preserve the existing
location, template ID, preview status, and warning headers.

Suggested response body:

```json
{
  "templateId": "template-id",
  "templateJson": {},
  "previewAvailable": true,
  "warnings": [],
  "classification": {
    "status": "ready",
    "metadata": {
      "slide_type": "architecture",
      "slide_purpose": "Explain a system architecture"
    }
  }
}
```

The example metadata is abbreviated. The real response uses the complete validated
`SlideRetrievalMetadata` contract.

## Failure and cancellation behavior

Saving first means a classification failure can occur after the template has committed. The code
must acknowledge that partial outcome instead of claiming the entire workflow is atomic.

| Failure point | Database result | HTTP result |
| --- | --- | --- |
| Upload, conversion, or single-slide validation fails | Nothing is saved | Existing sanitized 4xx/5xx error |
| Initial template transaction fails | Nothing is saved | Sanitized 500 error |
| Provider refuses or returns invalid output | Template remains; classification becomes `failed` | Sanitized 502 error |
| Provider is unavailable or rate limited | Template remains; classification becomes `failed` | Sanitized 502 or 503 error |
| Provider timeout | Template remains; classification becomes `failed` | Sanitized 504 error |
| Client disconnects after persistence | Template remains; provider call is aborted; classification becomes `failed` when possible | No response can be delivered |
| Final classification update fails | Template remains; classification is not ready | Sanitized 500 error |

For failures after persistence, set `X-Template-Id` when an HTTP response can still be written so
the caller can identify the saved template. Store only a sanitized error code such as
`classification_timeout`; never store raw provider bodies, prompts, slide text, or secrets.

There is no automatic retry in this plan. Retrying the HTTP import can create a second template, so
the UI should not retry automatically. Idempotency keys or a dedicated retry endpoint are separate
features and should be added only if duplicate imports become a real problem.

Do not delete the saved template automatically when classification fails. Deletion would introduce
another fallible compensation step and would contradict the explicit save-before-classify order.

## Timeout policy

The endpoint timeout must cover conversion, preview generation, classification, validation, and
both database transactions.

Use one explicit budget such as:

```text
preview timeout:         up to 15 seconds
classification timeout: up to 45 seconds
request timeout:         at least 70 seconds
```

The exact values remain configuration, but the invariant is:

```text
request timeout > preview timeout + classification timeout + persistence buffer
```

Configure any reverse proxy or browser client timeout consistently. The provider timeout and
request cancellation must both abort the OpenAI call. Do not allow classification to continue as
detached work after the request ends.

## Classification contract

Reuse the existing `SlideRetrievalMetadataSchema` and normalized snake_case fields from the larger
classification plan. Keeping the schema unchanged allows a future retrieval feature to consume the
stored metadata without reclassifying slides.

Continue to enforce:

- required non-empty descriptive fields;
- bounded strings and arrays;
- trimmed and case-insensitively de-duplicated array items;
- exact `content_density` values;
- required boolean facets;
- no unknown properties; and
- versioned schema, prompt, and input fingerprints.

The response exposes the application metadata contract, not raw OpenAI response objects or SDK
types.

## OpenAI boundary

Keep the existing server-only `SlideClassifier` interface and `OpenAISlideClassifier` adapter. The
adapter should:

- use the configured explicit model;
- use schema-constrained Responses API output;
- set `store: false`;
- send no tools or model-selected external URLs;
- bound output tokens and input sizes;
- forward cancellation and enforce its timeout; and
- translate provider errors into sanitized application errors.

The OpenAI API key stays in server configuration. It must never enter browser code, `VITE_*`
values, logs, response bodies, stored error messages, or test fixtures.

## Repository contract

Replace job-oriented repository methods in the synchronous path with direct state transitions:

```ts
insertWithProcessingClassification(
  record: TemplateInsert,
  classification: ProcessingSlideClassification,
  appId?: string,
): void

completeSlideClassification(result: CompleteSlideClassification): void

failSlideClassification(templateId: string, errorCode: string): void
```

The completion update must require the expected input fingerprint and `processing` status so a
stale response cannot overwrite a newer classification.

Remove synchronous-path dependencies on:

- `claimNextSlideClassification`;
- `claimNextSlideEmbedding`;
- retry timestamps;
- lease expiration timestamps;
- attempt counters; and
- embedding completion or failure methods.

If schema version 4 already exists in a developer database, do not add a destructive migration
solely to remove now-unused queue columns. Stop reading and writing them in the application first;
physical schema cleanup can be a later migration after compatibility needs are clear.

## Service ownership

`ImportService.importV2` owns the complete synchronous use case:

```text
convert -> prepare -> insert -> classify -> complete -> respond
```

Refactor `SlideClassificationService` into a narrow collaborator that builds provider input, calls
the classifier, validates the result, and returns normalized metadata. It should not accept a
claimed job or expose embedding behavior in this version.

This keeps HTTP details in the handler, import orchestration in `ImportService`, provider details in
the integration adapter, and persistence details in the repository.

## File-level implementation map

| File | Change |
| --- | --- |
| `server/src/services/ImportTemplateService.ts` | Replace worker notification with synchronous classification after the template commit and return ready metadata. |
| `server/src/services/SlideClassificationService.ts` | Accept saved slide input, call the classifier, and return validated normalized metadata; remove embedding-stage behavior. |
| `server/src/services/SlideClassificationWorker.ts` | Delete after all startup, test, and import references are removed. |
| `server/src/repositories/TemplateRepository.ts` | Replace claim/retry job methods with direct processing, completion, and failure transitions. |
| `server/src/repositories/SqliteTemplateRepository.ts` | Persist processing/ready/failed states without queue selection, leases, or retry scheduling. |
| `server/src/server.ts` | Construct and inject the classifier service directly into `ImportService`; remove worker startup and shutdown. |
| `server/src/config.ts` | Remove worker concurrency, retry, lease, and embedding settings; retain classification model, input, provider, and timeout settings. |
| `server/src/handlers/importHandlers.ts` | Return the completed classification contract and expose the saved template ID on post-persistence errors when possible. |
| `server/test/importV2Api.test.ts` | Cover synchronous success, provider failures, cancellation, response shape, and v1 isolation. |
| `server/test/slideClassification.test.ts` | Replace queue/worker tests with direct service and repository transition tests. |
| `server/README.md` and `docs/api/README.md` | Document that v2 waits for classification and may leave a saved failed template. |

Remove embedding and retrieval files only if they have no remaining callers. Do not broaden this
change into unrelated cleanup.

## Test plan

### Happy path

- Save the template before the fake classifier is invoked.
- Pass the exact saved template, kind, title, bounded digest, optional preview, and request signal.
- Persist normalized valid metadata and provenance.
- Return `201` only after the classification update commits.
- Return the exact template ID, hydrated template JSON, preview state, warnings, and ready metadata.
- Preserve `Location`, `X-Template-Id`, preview, and warning headers.

### Failure paths

- Do not save anything when conversion or single-slide validation fails.
- Keep the template and record a sanitized failure when the classifier refuses, times out, is rate
  limited, is unavailable, or returns invalid output.
- Do not return a successful import response when classification fails.
- Abort the provider call when the request disconnects or times out.
- Do not continue classification in the background after cancellation.
- Do not expose raw provider errors, prompt text, slide content, preview bytes, or API keys.
- Do not retry provider calls automatically.

### Compatibility

- Existing `POST /api/v1/import` remains classification-free.
- Existing batch import remains classification-free.
- Built-in seeding does not call the classifier.
- App scoping, template retrieval, preview retrieval, asset retrieval, and deletion remain unchanged.
- A stale provider completion cannot overwrite a different input fingerprint.

Use injected fake classifiers for all routine tests. Do not make live OpenAI calls.

## Implementation sequence

1. Lock the synchronous `POST /api/v2/import` success and error response contracts in request tests.
2. Add direct repository transitions for processing, ready, and failed classification states.
3. Refactor `SlideClassificationService` to perform one classification and return normalized
   metadata without job or embedding concepts.
4. Update `ImportService.importV2` to save, classify, complete, and return within one request.
5. Inject the classification service directly from `server.ts`.
6. Remove worker startup, notification, shutdown, claim, lease, retry, and embedding-stage code that
   has no remaining callers.
7. Align request and provider timeout configuration.
8. Update API and server documentation.
9. Run focused tests, broad verification, and inspect the final diff for unrelated changes.

## Verification gate

Run from the repository root:

```sh
npm run test --workspace @diligence-studio/server -- test/importV2Api.test.ts
npm run test --workspace @diligence-studio/server -- test/slideClassification.test.ts
npm run typecheck
npm test
npm run build
git diff --check
git status --short
```

No live OpenAI request is required for routine verification. If a live smoke test is later
authorized, use a synthetic non-sensitive slide, a strict spending boundary, and a disposable
database.

## Definition of done

- A successful v2 import saves the template before calling the classifier.
- The request waits until validated classification metadata is committed.
- A successful response contains the template and `classification.status = 'ready'`.
- No background worker, queue, polling, lease, scheduled retry, or embedding stage remains in this
  workflow.
- Failures after the initial commit leave the template stored and mark classification failed.
- Request cancellation stops the provider call rather than detaching work.
- Existing v1 and batch import contracts remain unchanged.
- Server tests, repository tests, typechecking, the complete test suite, the browser build, and
  `git diff --check` pass.
- No secrets, raw provider output, sensitive slide content, generated artifacts, or unrelated
  changes enter the diff.

## Accepted tradeoffs

This synchronous design is simpler, but it deliberately accepts:

- slower import responses;
- a long-lived HTTP request during classification;
- no automatic recovery when the process exits mid-classification;
- no automatic retry for transient provider failures;
- possible saved-but-failed templates;
- one OpenAI call per concurrent import without worker-level concurrency control; and
- the need to revisit queues if production traffic or reliability requirements grow.

Those tradeoffs are acceptable only while the feature is low volume and the product values a
straightforward completed-or-error response more than asynchronous reliability.
