import { SlideProviderError, type SlideClassifier, type SlideEmbedder } from '../integrations/SlideProvider'
import {
  buildSlideClassificationInput,
  type SlideClassificationInputLimits,
} from '../lib/retrieval/SlideClassificationInput'
import {
  buildSlideEmbeddingDocuments,
  buildSlideEmbeddingFingerprint,
} from '../lib/retrieval/SlideEmbeddingDocument'
import {
  assertContentSlotsExist,
  type SlideRetrievalMetadata,
} from '../lib/retrieval/SlideRetrievalMetadata'
import type {
  SlideClassificationJob,
  SlideEmbeddingJob,
  TemplateRepository,
} from '../repositories/TemplateRepository'

export type SlideClassificationServiceOptions = {
  classifier: SlideClassifier
  classificationModel: string
  embedder: SlideEmbedder
  embeddingDimensions?: number
  embeddingMaxTextBytes: number
  embeddingModel: string
  inputLimits: Partial<SlideClassificationInputLimits>
  repository: TemplateRepository
}

export class SlideClassificationService {
  constructor(private readonly options: SlideClassificationServiceOptions) {}

  /** Classify and embed one saved slide, then publish every retrieval artifact atomically. */
  async process(
    job: SlideClassificationJob,
    signal?: AbortSignal,
  ): Promise<SlideRetrievalMetadata> {
    const prepared = await this.#classify(job, signal)
    const [subject, capability] = await Promise.all([
      this.#embed(prepared.subjectEmbeddingDocument, signal),
      this.#embed(prepared.capabilityEmbeddingDocument, signal),
    ])
    if (subject.dimensions !== capability.dimensions) {
      throw new Error('Subject and capability embedding dimensions do not match.')
    }
    signal?.throwIfAborted()
    this.options.repository.completeSlideProcessing({
      ...prepared,
      capabilityVector: capability.vector,
      embeddingDimensions: subject.dimensions,
      embeddingModel: this.options.embeddingModel,
      subjectVector: subject.vector,
      templateId: job.template.templateId,
    })
    return prepared.metadata
  }

  /** Background-worker compatibility path; synchronous imports call process instead. */
  async classify(job: SlideClassificationJob, signal?: AbortSignal) {
    const result = await this.#classify(job, signal)
    this.options.repository.completeSlideClassification({
      ...result,
      templateId: job.template.templateId,
    })
  }

  async #classify(job: SlideClassificationJob, signal?: AbortSignal) {
    const input = buildSlideClassificationInput({
      kind: job.kind,
      limits: this.options.inputLimits,
      preview: job.preview,
      templateJson: job.template.templateJson,
      title: job.title,
    })
    const metadata = await this.options.classifier.classify(input, signal)
    const elementIds = new Set(job.template.templateJson.presentation.slides.flatMap((slide) => (
      slide.elements.map((element) => element.id)
    )))
    try {
      assertContentSlotsExist(metadata, elementIds)
    } catch (error) {
      throw new SlideProviderError('classification_invalid_output', false, { cause: error })
    }
    const documents = buildSlideEmbeddingDocuments(
      job.title,
      job.kind,
      metadata,
      this.options.embeddingMaxTextBytes,
    )
    const subjectEmbeddingFingerprint = buildSlideEmbeddingFingerprint(
      'subject',
      documents.subject,
      this.options.embeddingModel,
      this.options.embeddingDimensions,
    )
    const capabilityEmbeddingFingerprint = buildSlideEmbeddingFingerprint(
      'capability',
      documents.capability,
      this.options.embeddingModel,
      this.options.embeddingDimensions,
    )
    return {
      capabilityEmbeddingDocument: documents.capability,
      capabilityEmbeddingFingerprint,
      classifiedAt: new Date().toISOString(),
      classifiedFingerprint: job.inputFingerprint,
      metadata,
      model: this.options.classificationModel,
      subjectEmbeddingDocument: documents.subject,
      subjectEmbeddingFingerprint,
    }
  }

  /** Background-worker compatibility path; synchronous imports call process instead. */
  async embed(job: SlideEmbeddingJob, signal?: AbortSignal) {
    const [subject, capability] = await Promise.all([
      this.#embed(job.subjectDocument, signal),
      this.#embed(job.capabilityDocument, signal),
    ])
    if (subject.dimensions !== capability.dimensions) {
      throw new Error('Subject and capability embedding dimensions do not match.')
    }
    this.options.repository.completeSlideEmbedding(
      job.templateId,
      job.subjectFingerprint,
      job.capabilityFingerprint,
      subject.vector,
      capability.vector,
      this.options.embeddingModel,
      subject.dimensions,
    )
  }

  async #embed(document: string, signal?: AbortSignal) {
    const vector = await this.options.embedder.embed(document, signal)
    return {
      dimensions: this.options.embeddingDimensions ?? vector.length,
      vector,
    }
  }
}
