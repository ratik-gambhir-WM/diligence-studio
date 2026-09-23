import { getExtension } from '../utils/files'

export type AttachmentDebugSummary = {
  count: number
  totalBytes: number
  imageCount: number
  fileCount: number
  extensions: string[]
}

export function summarizeAttachments(files: readonly File[]): AttachmentDebugSummary {
  const extensions = new Set<string>()

  for (const file of files) {
    const extension = getExtension(file.name)
    if (extension) {
      extensions.add(extension)
    }
  }

  const imageCount = files.filter((file) => {
    if (file.type.startsWith('image/')) {
      return true
    }

    return ['jpg', 'jpeg', 'png'].includes(getExtension(file.name))
  }).length

  return {
    count: files.length,
    totalBytes: files.reduce((total, file) => total + file.size, 0),
    imageCount,
    fileCount: files.length - imageCount,
    extensions: [...extensions].sort(),
  }
}

export function logAttachmentDebug(event: string, details: Record<string, unknown>) {
  if (!import.meta.env.DEV) {
    return
  }

  console.info('[diligence-studio:' + event + ']', details)
}
