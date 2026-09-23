import { useEffect, useState } from 'react'
import type { ChangeEvent } from 'react'

import {
  logAttachmentDebug,
  summarizeAttachments,
} from '../lib/attachmentDiagnostics'
import { getExtension } from '../utils/files'

export type AttachmentMode = 'template-context' | 'upload-only'

const ALLOWED_EXTENSIONS = new Set([
  'docx',
  'pdf',
  'ppt',
  'pptx',
  'png',
  'jpg',
  'jpeg',
  'md',
  'markdown',
  'txt',
  'rtf',
])

export const ACCEPT_ATTR = [
  '.docx',
  '.ppt',
  '.rtf',
  '.pdf',
  '.pptx',
  '.png',
  '.jpg',
  '.jpeg',
  '.md',
  '.markdown',
  '.txt',
  'text/plain',
  'image/png',
  'image/jpeg',
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  'application/pdf',
  'application/vnd.ms-powerpoint',
  'application/vnd.openxmlformats-officedocument.presentationml.presentation',
].join(',')

type AttachmentRecord = {
  id: string
  file: File
  mode: AttachmentMode
  source: 'upload' | 'sharepoint'
  sharePointFileId?: string
  sharePointPath?: string
  sharePointSourceUrl?: string
}

export type SharePointAttachmentInput = {
  file: File
  fileId: string
  path: string
  sourceUrl: string
}

export type UploadOnlyAttachmentItem = {
  id: string
  name: string
  path?: string
  size: number
  source: 'upload' | 'sharepoint'
}

let attachmentSequence = 0

function createAttachmentId() {
  attachmentSequence += 1
  return `attachment-${attachmentSequence}`
}

function getAttachmentCountLabel(count: number) {
  if (count === 0) return 'No files attached'
  if (count === 1) return '1 file attached'
  return `${count} files attached`
}

function partitionAttachments(files: File[]) {
  const validAttachments: File[] = []
  const invalidFileNames: string[] = []

  for (const file of files) {
    const ext = getExtension(file.name)

    if (ALLOWED_EXTENSIONS.has(ext)) {
      validAttachments.push(file)
    } else {
      invalidFileNames.push(file.name)
    }
  }

  return { invalidFileNames, validAttachments }
}

function formatUnsupportedFilesMessage(fileNames: string[]) {
  return fileNames.length > 0
    ? `Unsupported file type: ${fileNames.join(', ')}. Allowed: DOCX, PDF, PPT/PPTX, PNG/JPG, Markdown, TXT.`
    : ''
}

export function useDiagramSession() {
  const [attachmentRecords, setAttachmentRecords] = useState<AttachmentRecord[]>([])
  const [error, setError] = useState('')

  const attachments = attachmentRecords.map((attachment) => attachment.file)
  const uploadOnlyAttachments = attachmentRecords
    .filter((attachment) => attachment.mode === 'upload-only')
    .map((attachment) => attachment.file)

  const attachmentCountLabel = getAttachmentCountLabel(attachments.length)

  useEffect(() => {
    logAttachmentDebug('attachment-state', {
      total: summarizeAttachments(attachments),
      uploadOnly: summarizeAttachments(uploadOnlyAttachments),
      modeCounts: {
        templateContext: attachmentRecords.filter((attachment) => attachment.mode === 'template-context').length,
        uploadOnly: attachmentRecords.filter((attachment) => attachment.mode === 'upload-only').length,
      },
      sourceCounts: {
        upload: attachmentRecords.filter((attachment) => attachment.source === 'upload').length,
        sharePoint: attachmentRecords.filter((attachment) => attachment.source === 'sharepoint').length,
      },
    })
  }, [attachmentRecords])

  function handleFiles(event: ChangeEvent<HTMLInputElement>, mode: AttachmentMode = 'template-context') {
    const incomingFiles = Array.from(event.target.files ?? [])

    if (incomingFiles.length === 0) {
      return []
    }

    const { invalidFileNames, validAttachments } = partitionAttachments(incomingFiles)

    logAttachmentDebug('file-selection', {
      mode,
      selectedCount: incomingFiles.length,
      acceptedCount: validAttachments.length,
      rejectedCount: invalidFileNames.length,
      accepted: summarizeAttachments(validAttachments),
    })

    if (validAttachments.length > 0) {
      setAttachmentRecords((previousAttachments) => {
        const hasSharePointAttachments = previousAttachments.some(
          (attachment) => attachment.mode === 'upload-only' && attachment.source === 'sharepoint',
        )
        const retainedAttachments = mode === 'upload-only' && hasSharePointAttachments
          ? previousAttachments.filter((attachment) => attachment.mode !== 'upload-only')
          : previousAttachments

        return [
          ...retainedAttachments,
          ...validAttachments.map((file) => ({
            file,
            id: createAttachmentId(),
            mode,
            source: 'upload' as const,
          })),
        ]
      })
    }

    setError(formatUnsupportedFilesMessage(invalidFileNames))

    event.target.value = ''

    return validAttachments
  }

  function removeAttachment(index: number) {
    setAttachmentRecords((previousAttachments) =>
      previousAttachments.filter((_, currentIndex) => currentIndex !== index),
    )
  }

  function addSharePointAttachments(
    attachmentsToAdd: SharePointAttachmentInput[],
  ) {
    if (attachmentsToAdd.length === 0) return

    logAttachmentDebug('sharepoint-files-added', {
      requestedCount: attachmentsToAdd.length,
      files: summarizeAttachments(attachmentsToAdd.map((attachment) => attachment.file)),
    })

    setAttachmentRecords((previousAttachments) => {
      const retainedAttachments = previousAttachments.filter(
        (attachment) => attachment.mode !== 'upload-only',
      )
      const existingSharePointIds = new Set(
        retainedAttachments
          .map((attachment) => attachment.sharePointFileId)
          .filter((fileId): fileId is string => fileId !== undefined),
      )

      return [
        ...retainedAttachments,
        ...attachmentsToAdd
          .filter((attachment) => !existingSharePointIds.has(attachment.fileId))
          .map((attachment) => ({
            file: attachment.file,
            id: createAttachmentId(),
            mode: 'upload-only' as const,
            sharePointFileId: attachment.fileId,
            sharePointPath: attachment.path,
            sharePointSourceUrl: attachment.sourceUrl,
            source: 'sharepoint' as const,
          })),
      ]
    })
  }

  function removeUploadOnlyAttachment(id: string) {
    setAttachmentRecords((previousAttachments) => {
      return previousAttachments.filter((attachment) => attachment.id !== id)
    })
  }

  const uploadOnlyAttachmentItems: UploadOnlyAttachmentItem[] = attachmentRecords
    .filter((attachment) => attachment.mode === 'upload-only')
    .map((attachment) => ({
      id: attachment.id,
      name: attachment.file.name,
      path: attachment.sharePointPath,
      size: attachment.file.size,
      source: attachment.source,
    }))

  return {
    attachmentCountLabel,
    attachments,
    addSharePointAttachments,
    error,
    handleFiles,
    removeAttachment,
    removeUploadOnlyAttachment,
    uploadOnlyAttachmentItems,
    uploadOnlyAttachments,
  }
}
