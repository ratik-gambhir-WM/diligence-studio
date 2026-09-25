import type OpenAI from 'openai'

import { validateEmbedding } from '../lib/retrieval/SlideVector'
import { SlideProviderError, type SlideEmbedder } from './SlideProvider'
import { translateOpenAIError } from './OpenAISlideClassifier'

export type OpenAISlideEmbedderOptions = {
  client: SlideEmbeddingsClient
  dimensions?: number
  maxInputBytes: number
  model: string
  timeoutMs: number
}

export type SlideEmbeddingsClient = {
  embeddings: {
    create(
      body: OpenAI.Embeddings.EmbeddingCreateParams,
      options?: OpenAI.RequestOptions,
    ): Promise<Pick<OpenAI.Embeddings.CreateEmbeddingResponse, 'data'>>
  }
}

export class OpenAISlideEmbedder implements SlideEmbedder {
  constructor(private readonly options: OpenAISlideEmbedderOptions) {}

  /** Create and validate exactly one embedding for a bounded canonical document or query. */
  async embed(text: string, signal?: AbortSignal): Promise<readonly number[]> {
    if (!text.trim()) throw new SlideProviderError('embedding_invalid_vector', false)
    if (Buffer.byteLength(text, 'utf8') > this.options.maxInputBytes) {
      throw new SlideProviderError('embedding_invalid_vector', false)
    }
    try {
      const response = await this.options.client.embeddings.create({
        input: text,
        model: this.options.model,
        ...(this.options.dimensions === undefined ? {} : { dimensions: this.options.dimensions }),
        encoding_format: 'float',
      }, { maxRetries: 0, signal, timeout: this.options.timeoutMs })
      if (response.data.length !== 1) throw new SlideProviderError('embedding_invalid_vector', false)
      const vector = response.data[0]?.embedding
      if (!vector) throw new SlideProviderError('embedding_invalid_vector', false)
      try {
        validateEmbedding(vector, this.options.dimensions)
      } catch (error) {
        throw new SlideProviderError('embedding_invalid_vector', false, { cause: error })
      }
      return vector
    } catch (error) {
      if (error instanceof SlideProviderError) throw error
      throw translateOpenAIError(error, 'embedding')
    }
  }
}
