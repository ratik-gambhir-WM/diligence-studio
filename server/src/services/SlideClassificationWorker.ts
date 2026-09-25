import { SlideProviderError } from '../integrations/SlideProvider'
import type { TemplateRepository } from '../repositories/TemplateRepository'
import type { SlideClassificationService } from './SlideClassificationService'

export type SlideClassificationWorkerOptions = {
  classificationMaxAttempts: number
  classificationTimeoutMs: number
  concurrency: number
  embeddingMaxAttempts: number
  embeddingTimeoutMs: number
  leasePaddingMs?: number
  now?: () => Date
  random?: () => number
  repository: TemplateRepository
  retryBaseMs?: number
  service: SlideClassificationService
}

/**
 * Dormant compatibility implementation. The server intentionally does not construct this worker;
 * import v2 now awaits classification, embedding, and atomic in-memory storage itself.
 */
export class SlideClassificationWorker {
  readonly #abortController = new AbortController()
  readonly #active = new Set<Promise<void>>()
  readonly #leasePaddingMs: number
  readonly #now: () => Date
  readonly #random: () => number
  readonly #retryBaseMs: number
  #pollTimer: NodeJS.Timeout | undefined
  #preferredStage: 'classification' | 'embedding' = 'classification'
  #running = false
  #stopping = false

  constructor(private readonly options: SlideClassificationWorkerOptions) {
    this.#leasePaddingMs = options.leasePaddingMs ?? 10_000
    this.#now = options.now ?? (() => new Date())
    this.#random = options.random ?? Math.random
    this.#retryBaseMs = options.retryBaseMs ?? 1_000
  }

  /** Start lightweight polling and immediately check for queued work. */
  start() {
    if (this.#pollTimer || this.#stopping) return
    this.#pollTimer = setInterval(() => this.notify(), 1_000)
    this.#pollTimer.unref()
    this.notify()
  }

  /** Schedule a non-overlapping drain after an import or polling tick reports possible work. */
  notify() {
    if (this.#stopping || this.#running) return
    this.#running = true
    queueMicrotask(() => {
      void this.#drain()
        .catch(() => undefined)
        .finally(() => {
          this.#running = false
        })
    })
  }

  /** Stop claiming work, allow a bounded drain, then cancel remaining provider calls. */
  async stop(drainTimeoutMs = 10_000) {
    this.#stopping = true
    if (this.#pollTimer) clearInterval(this.#pollTimer)
    this.#pollTimer = undefined
    const drain = Promise.allSettled([...this.#active])
    let timeout: NodeJS.Timeout | undefined
    await Promise.race([
      drain,
      new Promise<void>((resolve) => {
        timeout = setTimeout(() => {
          this.#abortController.abort()
          resolve()
        }, drainTimeoutMs)
        timeout.unref()
      }),
    ])
    if (timeout) clearTimeout(timeout)
  }

  async #drain() {
    while (!this.#stopping) {
      while (this.#active.size < this.options.concurrency && !this.#stopping) {
        const work = this.#claimAndRun()
        if (!work) break
        this.#active.add(work)
        void work.then(
          () => this.#active.delete(work),
          () => this.#active.delete(work),
        )
      }
      if (this.#active.size === 0) return
      await Promise.race(this.#active)
    }
  }

  #claimAndRun(): Promise<void> | undefined {
    const now = this.#now()
    const classification = () => this.#claimClassification(now)
    const embedding = () => this.#claimEmbedding(now)
    const preferredStage = this.#preferredStage
    const first = preferredStage === 'classification' ? classification() : embedding()
    if (first) {
      this.#preferredStage = preferredStage === 'classification' ? 'embedding' : 'classification'
      return first
    }
    return preferredStage === 'classification' ? embedding() : classification()
  }

  #claimClassification(now: Date): Promise<void> | undefined {
    const classification = this.options.repository.claimNextSlideClassification(
      now.toISOString(),
      new Date(now.getTime() + this.options.classificationTimeoutMs + this.#leasePaddingMs).toISOString(),
    )
    if (classification) return this.#runClassification(classification.template.templateId, classification.attemptCount, () => (
      this.options.service.classify(classification, this.#abortController.signal)
    ))

    return undefined
  }

  #claimEmbedding(now: Date): Promise<void> | undefined {
    const embedding = this.options.repository.claimNextSlideEmbedding(
      now.toISOString(),
      new Date(now.getTime() + this.options.embeddingTimeoutMs + this.#leasePaddingMs).toISOString(),
    )
    if (embedding) return this.#runEmbedding(embedding.templateId, embedding.attemptCount, () => (
      this.options.service.embed(embedding, this.#abortController.signal)
    ))
    return undefined
  }

  async #runClassification(templateId: string, attempt: number, operation: () => Promise<void>) {
    try {
      await operation()
    } catch (error) {
      const retryable = error instanceof SlideProviderError && error.retryable
      const retryAt = retryable && attempt < this.options.classificationMaxAttempts
        ? this.#retryAt(attempt)
        : null
      this.options.repository.recordSlideClassificationFailure(
        templateId,
        error instanceof SlideProviderError ? error.code : 'classification_internal',
        retryAt,
      )
    }
  }

  async #runEmbedding(templateId: string, attempt: number, operation: () => Promise<void>) {
    try {
      await operation()
    } catch (error) {
      const retryable = error instanceof SlideProviderError && error.retryable
      const retryAt = retryable && attempt < this.options.embeddingMaxAttempts
        ? this.#retryAt(attempt)
        : null
      this.options.repository.recordSlideEmbeddingFailure(
        templateId,
        error instanceof SlideProviderError ? error.code : 'embedding_internal',
        retryAt,
      )
    }
  }

  #retryAt(attempt: number) {
    const exponential = this.#retryBaseMs * (2 ** Math.max(0, attempt - 1))
    const jitter = Math.floor(exponential * 0.25 * this.#random())
    return new Date(this.#now().getTime() + exponential + jitter).toISOString()
  }
}
