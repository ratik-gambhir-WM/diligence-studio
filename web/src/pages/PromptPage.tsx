import type { ChangeEvent } from 'react'

import { AppNav } from '../components/AppNav'
import { BrandLockup } from '../components/BrandLockup'
import { Button } from '../components/Button'
import { FileList } from '../components/FileList'
import { FileUploadButton } from '../components/FileUploadButton'
import { PageShell } from '../components/PageShell'
import { SharePointFilePicker } from '../components/SharePointFilePicker'
import { SnailLoader } from '../components/SnailLoader'
import { StudioPanel } from '../components/StudioPanel'
import type { SharePointFile, SharePointResolvedResource } from '../lib/api/sharepointApi'
import { formatFileSize } from '../utils/files'

type UploadOnlyFile = {
  id: string
  name: string
  path?: string
  size: number
  source: 'upload' | 'sharepoint'
}

type PromptPageProps = {
  acceptAttr: string
  createMode: boolean
  isUploadOnlySelecting: boolean
  onOpenCommentaryPicker: () => void
  onOpenDiagramPicker: () => void
  onOpenInputPage: () => void
  onOpenJsonInput: () => void
  onCreateModeChange: (createMode: boolean) => void
  onRemoveUploadOnlyFile: (id: string) => void
  onResolveSharePoint: () => void
  onSharePointUrlChange: (url: string) => void
  onCloseSharePointPicker: () => void
  onUseSharePointFiles: (files: SharePointFile[]) => void
  onUploadOnlyFileChange: (event: ChangeEvent<HTMLInputElement>) => void
  onUploadOnlySubmit: () => void
  selectedArchitectureDiagramId: string
  uploadOnlyFiles: UploadOnlyFile[]
  uploadOnlyFileCount: number
  uploadOnlyError: string
  sharePointError: string
  sharePointIsDownloading: boolean
  sharePointIsResolving: boolean
  sharePointResource: SharePointResolvedResource | null
  sharePointUrl: string
}

export function PromptPage({
  acceptAttr,
  createMode,
  isUploadOnlySelecting,
  onOpenCommentaryPicker,
  onOpenDiagramPicker,
  onOpenInputPage,
  onOpenJsonInput,
  onCreateModeChange,
  onRemoveUploadOnlyFile,
  onResolveSharePoint,
  onSharePointUrlChange,
  onCloseSharePointPicker,
  onUseSharePointFiles,
  onUploadOnlyFileChange,
  onUploadOnlySubmit,
  selectedArchitectureDiagramId,
  uploadOnlyFiles,
  uploadOnlyFileCount,
  uploadOnlyError,
  sharePointError,
  sharePointIsDownloading,
  sharePointIsResolving,
  sharePointResource,
  sharePointUrl,
}: PromptPageProps) {
  if (isUploadOnlySelecting) {
    return <SnailLoader />
  }

  return (
    <PageShell>
      <AppNav
        activePage="diagramming"
        onOpenCommentaryPicker={onOpenCommentaryPicker}
        onOpenDiagramPicker={onOpenDiagramPicker}
        onOpenInputPage={onOpenInputPage}
        onOpenJsonInput={onOpenJsonInput}
      />

      <div className="grid min-h-0 flex-1 place-items-center pt-8">
        <StudioPanel className="w-full max-w-[780px] max-[640px]:rounded-[1.5rem] max-[640px]:p-5">
          <label className="absolute top-5 right-5 z-10 flex cursor-pointer items-center gap-2 text-[0.72rem] font-bold tracking-[0.14em] text-[#eef3ff] uppercase">
            <span>Create</span>
            <span
              className={[
                'relative h-6 w-11 rounded-full border transition',
                createMode ? 'border-[#f3c316] bg-[#f3c316]' : 'border-[#28304a] bg-[#080c1c]',
              ]
                .filter(Boolean)
                .join(' ')}
            >
              <input
                aria-label="Create a new diagram from uploaded files"
                type="checkbox"
                checked={createMode}
                onChange={(event) => onCreateModeChange(event.currentTarget.checked)}
                disabled={isUploadOnlySelecting}
                className="sr-only"
              />
              <span
                className={[
                  'absolute top-1 h-4 w-4 rounded-full bg-white transition',
                  createMode ? 'left-6' : 'left-1',
                ]
                  .filter(Boolean)
                  .join(' ')}
              />
            </span>
          </label>

          <BrandLockup align="center" markSize="lg" title="west monroe" />

          <div
            className="relative mt-8 flex flex-wrap items-center justify-center gap-3"
            data-selected-diagram-id={selectedArchitectureDiagramId}
          >
            <FileUploadButton
              id="upload-only-attachments"
              accept={acceptAttr}
              onChange={onUploadOnlyFileChange}
              disabled={isUploadOnlySelecting}
              variant="secondary"
              className="px-4 py-2.5 text-[0.9rem]"
            >
              {isUploadOnlySelecting ? 'Selecting...' : 'Upload file'}
            </FileUploadButton>
            <div className="flex min-w-[18rem] flex-1 flex-wrap gap-2">
              <label className="sr-only" htmlFor="sharepoint-resource-url">
                SharePoint file or folder link
              </label>
              <input
                id="sharepoint-resource-url"
                type="url"
                value={sharePointUrl}
                onChange={(event) => onSharePointUrlChange(event.currentTarget.value)}
                placeholder="Paste a SharePoint file or folder link"
                disabled={isUploadOnlySelecting || sharePointIsResolving || sharePointIsDownloading}
                className="min-w-0 flex-1 rounded-full border border-[#28304a] bg-[#080c1c] px-4 py-2.5 text-[0.84rem] text-[#eef3ff] outline-none placeholder:text-[#737b95] focus:border-[#f3c316]"
              />
              <Button
                type="button"
                onClick={onResolveSharePoint}
                disabled={isUploadOnlySelecting || sharePointIsResolving || sharePointIsDownloading || !sharePointUrl.trim()}
                variant="secondary"
                className="px-4 py-2.5 text-[0.9rem]"
              >
                {sharePointIsResolving ? 'Loading...' : 'Load SharePoint'}
              </Button>
            </div>
            <Button
              type="button"
              onClick={onUploadOnlySubmit}
              disabled={isUploadOnlySelecting || uploadOnlyFileCount === 0}
              variant="primary"
              className="px-4 py-2.5 text-[0.9rem] disabled:opacity-50"
            >
              {isUploadOnlySelecting ? 'Working...' : createMode ? 'Create diagram' : 'Submit files'}
            </Button>
          </div>

          {uploadOnlyError && (
            <p className="relative mt-3 text-center text-[0.92rem] leading-5 text-[#ffb5b5]">
              {uploadOnlyError}
            </p>
          )}

          {sharePointError && (
            <p className="relative mt-3 text-center text-[0.92rem] leading-5 text-[#ffb5b5]" role="alert">
              {sharePointError}
            </p>
          )}

          <FileList
            files={uploadOnlyFiles}
            formatFileSize={formatFileSize}
            label="Added files"
            onRemove={(index) => {
              const file = uploadOnlyFiles[index]
              if (file) {
                onRemoveUploadOnlyFile(file.id)
              }
            }}
            showCountHeader
          />
        </StudioPanel>
      </div>

      {sharePointResource?.kind === 'folder' && (
        <SharePointFilePicker
          error={sharePointError}
          isDownloading={sharePointIsDownloading}
          onClose={onCloseSharePointPicker}
          onConfirm={onUseSharePointFiles}
          resource={sharePointResource}
        />
      )}
    </PageShell>
  )
}
