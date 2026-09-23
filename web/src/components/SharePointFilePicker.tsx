import { useEffect, useMemo, useState } from 'react'

import { Modal } from './Modal'
import { Button } from './Button'
import type { SharePointFile, SharePointResolvedResource } from '../lib/api/sharepointApi'
import { formatFileSize } from '../utils/files'

type SharePointFilePickerProps = {
  error: string
  isDownloading: boolean
  onClose: () => void
  onConfirm: (files: SharePointFile[]) => void
  resource: SharePointResolvedResource
}

export function SharePointFilePicker({
  error,
  isDownloading,
  onClose,
  onConfirm,
  resource,
}: SharePointFilePickerProps) {
  const [selectedFileIds, setSelectedFileIds] = useState<Set<string>>(new Set())

  useEffect(() => {
    setSelectedFileIds(new Set())
  }, [resource])

  const selectedFiles = useMemo(
    () => resource.files.filter((file) => selectedFileIds.has(file.fileId)),
    [resource.files, selectedFileIds],
  )

  function toggleFile(fileId: string) {
    setSelectedFileIds((current) => {
      const next = new Set(current)
      if (next.has(fileId)) {
        next.delete(fileId)
      } else {
        next.add(fileId)
      }
      return next
    })
  }

  function toggleAll() {
    setSelectedFileIds((current) => (
      current.size === resource.files.length
        ? new Set()
        : new Set(resource.files.map((file) => file.fileId))
    ))
  }

  return (
    <Modal
      labelledBy="sharepoint-file-picker-title"
      onClose={isDownloading ? () => undefined : onClose}
      title={`Select files from ${resource.name}`}
    >
      <div className="grid gap-4">
        <div className="flex flex-wrap items-center justify-between gap-3 text-[0.84rem] text-[#a8afc4]">
          <span>{resource.files.length} file{resource.files.length === 1 ? '' : 's'} found</span>
          <Button
            type="button"
            variant="link"
            onClick={toggleAll}
            disabled={isDownloading || resource.files.length === 0}
          >
            {selectedFileIds.size === resource.files.length ? 'Clear all' : 'Select all'}
          </Button>
        </div>

        {resource.files.length === 0 ? (
          <p className="text-[0.9rem] leading-6 text-[#a8afc4]">
            No supported files were found in this SharePoint folder.
          </p>
        ) : (
          <ul className="grid max-h-[24rem] gap-2 overflow-y-auto pr-1" aria-label="SharePoint files">
            {resource.files.map((file) => (
              <li key={file.fileId} className="rounded-[0.8rem] border border-[#28304a] bg-[#080c1c] p-3">
                <label className="flex cursor-pointer items-start gap-3">
                  <input
                    type="checkbox"
                    checked={selectedFileIds.has(file.fileId)}
                    onChange={() => toggleFile(file.fileId)}
                    disabled={isDownloading}
                    className="mt-1 h-4 w-4 accent-[#f3c316]"
                  />
                  <span className="min-w-0">
                    <span className="block truncate text-[0.88rem] font-semibold text-[#eef3ff]">
                      {file.name}
                    </span>
                    <span className="mt-1 block truncate text-[0.75rem] text-[#a8afc4]">
                      {file.path || 'Root'} · {formatFileSize(file.size)}
                    </span>
                  </span>
                </label>
              </li>
            ))}
          </ul>
        )}

        {error && (
          <p className="text-[0.84rem] leading-5 text-[#ffb5b5]" role="alert">
            {error}
          </p>
        )}

        <div className="flex flex-wrap justify-end gap-2">
          <Button type="button" variant="secondary" onClick={onClose} disabled={isDownloading}>
            Cancel
          </Button>
          <Button
            type="button"
            variant="primary"
            onClick={() => onConfirm(selectedFiles)}
            disabled={isDownloading || selectedFiles.length === 0}
          >
            {isDownloading ? 'Loading files...' : `Use ${selectedFiles.length || ''} selected file${selectedFiles.length === 1 ? '' : 's'}`}
          </Button>
        </div>
      </div>
    </Modal>
  )
}
