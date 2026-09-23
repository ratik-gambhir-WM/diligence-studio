import { Button } from './Button'

type FileListItem = {
  id?: string
  name: string
  path?: string
  size: number
  source?: 'upload' | 'sharepoint'
}

type FileListProps = {
  files: FileListItem[]
  formatFileSize: (bytes: number) => string
  label: string
  onRemove?: (index: number) => void
  removeDisabled?: boolean
  showCountHeader?: boolean
}

export function FileList({
  files,
  formatFileSize,
  label,
  onRemove,
  removeDisabled = false,
  showCountHeader = false,
}: FileListProps) {
  if (files.length === 0) {
    return null
  }

  return (
    <div className="file-list">
      {showCountHeader && (
        <div className="file-list-header">
          <span>{label}</span>
          <span>
            {files.length} file{files.length === 1 ? '' : 's'}
          </span>
        </div>
      )}
      <ul className="file-list-items" aria-label={label}>
        {files.map((file, index) => (
          <li
            className="file-list-item"
            key={file.id ?? `${file.name}-${file.size}-${index}`}
          >
            <span className="file-list-name">
              <span>
                {file.name}
                {!showCountHeader && ` (${formatFileSize(file.size)})`}
              </span>
              {file.source === 'sharepoint' && (
                <span className="file-list-source">SharePoint{file.path ? ` · ${file.path}` : ''}</span>
              )}
            </span>
            {showCountHeader ? (
              <span className="file-list-actions">
                <span className="file-list-size">{formatFileSize(file.size)}</span>
                {onRemove && (
                  <Button
                    type="button"
                    variant="ghost"
                    disabled={removeDisabled}
                    className="file-list-remove-button file-list-remove-button-icon"
                    aria-label={`Remove ${file.name}`}
                    onClick={() => onRemove(index)}
                  >
                    X
                  </Button>
                )}
              </span>
            ) : (
              onRemove && (
                <Button
                  type="button"
                  variant="ghost"
                  disabled={removeDisabled}
                  className="file-list-remove-button"
                  onClick={() => onRemove(index)}
                >
                  Remove
                </Button>
              )
            )}
          </li>
        ))}
      </ul>
    </div>
  )
}
