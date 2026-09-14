import type { SlideClassifier, SlideEmbedder } from '../integrations/SlideProvider'
import {
  buildSlideClassificationInput,
  type SlideClassificationInputLimits,
} from '../lib/retrieval/SlideClassificationInput'
import {
  buildSlideEmbeddingDocument,
  buildSlideEmbeddingFingerprint,
} from '../lib/retrieval/SlideEmbeddingDocument'
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

  /** Build provider input, classify it, and atomically publish metadata plus pending embedding work. */
  async classify(job: SlideClassificationJob, signal?: AbortSignal) {
    const input = buildSlideClassificationInput({
      kind: job.kind,
      limits: this.options.inputLimits,
      preview: job.preview,
      templateJson: job.template.templateJson,
      title: job.title,
    })
    const metadata = await this.options.classifier.classify(input, signal)
    const embeddingDocument = buildSlideEmbeddingDocument(
      job.title,
      job.kind,
      metadata,
      this.options.embeddingMaxTextBytes,
    )
    const embeddingFingerprint = buildSlideEmbeddingFingerprint(
      embeddingDocument,
      this.options.embeddingModel,
      this.options.embeddingDimensions,
    )
    this.options.repository.completeSlideClassification({
      classifiedAt: new Date().toISOString(),
      classifiedFingerprint: job.inputFingerprint,
      embeddingDocument,
      embeddingFingerprint,
      metadata,
      model: this.options.classificationModel,
      templateId: job.template.templateId,
    })
  }

  /** Embed a previously classified document without repeating the classification call. */
  async embed(job: SlideEmbeddingJob, signal?: AbortSignal) {
    const vector = await this.options.embedder.embed(job.document, signal)
    const dimensions = this.options.embeddingDimensions ?? vector.length
    this.options.repository.completeSlideEmbedding(
      job.templateId,
      job.fingerprint,
      vector,
      this.options.embeddingModel,
      dimensions,
    )
  }
}
