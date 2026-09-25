import type { SlideClassificationInput } from '../lib/retrieval/SlideClassificationInput'
import type { SlideRetrievalMetadata } from '../lib/retrieval/SlideRetrievalMetadata'

export type SlideProviderErrorCode =
  | 'classification_incomplete'
  | 'classification_invalid_output'
  | 'classification_missing_output'
  | 'classification_rate_limited'
  | 'classification_refused'
  | 'classification_timeout'
  | 'classification_unavailable'
  | 'embedding_invalid_vector'
  | 'embedding_rate_limited'
  | 'embedding_timeout'
  | 'embedding_unavailable'

export class SlideProviderError extends Error {
  constructor(
    readonly code: SlideProviderErrorCode,
    readonly retryable: boolean,
    options?: ErrorOptions,
  ) {
    super('The slide provider operation failed.', options)
    this.name = 'SlideProviderError'
  }
}

export interface SlideClassifier {
  classify(input: SlideClassificationInput, signal?: AbortSignal): Promise<SlideRetrievalMetadata>
}

export interface SlideEmbedder {
  embed(text: string, signal?: AbortSignal): Promise<readonly number[]>
}

