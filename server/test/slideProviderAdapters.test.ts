import type OpenAI from 'openai'
import { describe, expect, it } from 'vitest'

import {
  OpenAISlideClassifier,
  type SlideResponsesClient,
} from '../src/integrations/OpenAISlideClassifier'
import {
  OpenAISlideEmbedder,
  type SlideEmbeddingsClient,
} from '../src/integrations/OpenAISlideEmbedder'
import { SlideProviderError } from '../src/integrations/SlideProvider'
import { VALID_METADATA } from './slideTestFixtures'

describe('OpenAI slide provider adapters', () => {
  it('sends a bounded non-stored schema response with no tools and forwards cancellation', async () => {
    let body: OpenAI.Responses.ResponseCreateParamsNonStreaming | undefined
    let requestOptions: OpenAI.RequestOptions | undefined
    const client: SlideResponsesClient = {
      responses: {
        parse: async (receivedBody, receivedOptions) => {
          body = receivedBody
          requestOptions = receivedOptions
          return { output: [], output_parsed: VALID_METADATA, status: 'completed' }
        },
      },
    }
    const signal = new AbortController().signal
    const result = await new OpenAISlideClassifier({
      client,
      maxOutputTokens: 900,
      model: 'configured-classifier',
      timeoutMs: 1_234,
    }).classify({
      digest: 'bounded digest',
      preview: { bytes: Buffer.from('png'), contentType: 'image/png' },
    }, signal)

    expect(result.slide_type).toBe('architecture-overview')
    expect(body).toMatchObject({
      max_output_tokens: 900,
      model: 'configured-classifier',
      store: false,
      tools: [],
    })
    expect(body?.input).toEqual([{
      role: 'user',
      content: [
        { text: 'bounded digest', type: 'input_text' },
        { detail: 'high', image_url: 'data:image/png;base64,cG5n', type: 'input_image' },
      ],
    }])
    expect(requestOptions).toMatchObject({ maxRetries: 0, signal, timeout: 1_234 })
  })

  it('maps incomplete and aborted classification attempts to sanitized codes', async () => {
    const incomplete = new OpenAISlideClassifier({
      client: { responses: { parse: async () => ({ output: [], status: 'incomplete' }) } },
      model: 'configured-classifier',
      timeoutMs: 100,
    })
    await expect(incomplete.classify({ digest: 'private slide text' }))
      .rejects.toMatchObject({ code: 'classification_incomplete', message: 'The slide provider operation failed.' })

    const aborted = new OpenAISlideClassifier({
      client: {
        responses: {
          parse: async () => {
            const error = new Error('raw private provider body')
            error.name = 'AbortError'
            throw error
          },
        },
      },
      model: 'configured-classifier',
      timeoutMs: 100,
    })
    await expect(aborted.classify({ digest: 'private slide text' }))
      .rejects.toMatchObject({ code: 'classification_timeout', retryable: true })
  })

  it('embeds one bounded document with exact model/dimensions and validates its vector', async () => {
    let body: OpenAI.Embeddings.EmbeddingCreateParams | undefined
    const client: SlideEmbeddingsClient = {
      embeddings: {
        create: async (receivedBody) => {
          body = receivedBody
          return { data: [{ embedding: [1, 0], index: 0, object: 'embedding' }] }
        },
      },
    }
    const embedder = new OpenAISlideEmbedder({
      client,
      dimensions: 2,
      maxInputBytes: 100,
      model: 'configured-embedding',
      timeoutMs: 200,
    })
    await expect(embedder.embed('canonical document')).resolves.toEqual([1, 0])
    expect(body).toEqual({
      dimensions: 2,
      encoding_format: 'float',
      input: 'canonical document',
      model: 'configured-embedding',
    })

    const invalid = new OpenAISlideEmbedder({
      client: { embeddings: { create: async () => ({
        data: [{ embedding: [0, 0], index: 0, object: 'embedding' }],
      }) } },
      dimensions: 2,
      maxInputBytes: 100,
      model: 'configured-embedding',
      timeoutMs: 200,
    })
    await expect(invalid.embed('canonical document')).rejects.toBeInstanceOf(SlideProviderError)
    await expect(embedder.embed(' '.repeat(101))).rejects.toMatchObject({ code: 'embedding_invalid_vector' })
  })
})
