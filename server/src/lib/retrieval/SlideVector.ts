/** Validate and encode a vector as a portable little-endian Float32 SQLite BLOB. */
export function serializeEmbedding(vector: readonly number[]): Buffer {
  validateEmbedding(vector)
  const bytes = Buffer.allocUnsafe(vector.length * Float32Array.BYTES_PER_ELEMENT)
  vector.forEach((value, index) => bytes.writeFloatLE(value, index * Float32Array.BYTES_PER_ELEMENT))
  return bytes
}

/** Decode a SQLite Float32 BLOB while enforcing its recorded dimensions and norm. */
export function deserializeEmbedding(bytes: Uint8Array, dimensions: number): number[] {
  if (!Number.isSafeInteger(dimensions) || dimensions < 1 || bytes.byteLength !== dimensions * 4) {
    throw new Error('Stored embedding dimensions are invalid.')
  }
  const buffer = Buffer.from(bytes)
  const vector = Array.from({ length: dimensions }, (_, index) => buffer.readFloatLE(index * 4))
  validateEmbedding(vector, dimensions)
  return vector
}

/** Reject empty, non-finite, zero-norm, or dimensionally inconsistent vectors. */
export function validateEmbedding(vector: readonly number[], dimensions?: number) {
  if (vector.length === 0 || (dimensions !== undefined && vector.length !== dimensions)) {
    throw new Error('Embedding dimensions do not match the configured dimensions.')
  }
  let normSquared = 0
  for (const value of vector) {
    if (!Number.isFinite(value)) throw new Error('Embedding contains a non-finite value.')
    normSquared += value * value
  }
  if (!Number.isFinite(normSquared) || normSquared === 0) {
    throw new Error('Embedding must have a finite non-zero norm.')
  }
}

/** Calculate cosine similarity after validating both vectors share a safe shape. */
export function cosineSimilarity(left: readonly number[], right: readonly number[]) {
  validateEmbedding(left)
  validateEmbedding(right, left.length)
  let dot = 0
  let leftNorm = 0
  let rightNorm = 0
  for (let index = 0; index < left.length; index += 1) {
    const leftValue = left[index] ?? 0
    const rightValue = right[index] ?? 0
    dot += leftValue * rightValue
    leftNorm += leftValue * leftValue
    rightNorm += rightValue * rightValue
  }
  return dot / Math.sqrt(leftNorm * rightNorm)
}
