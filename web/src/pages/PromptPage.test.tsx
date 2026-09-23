import { render, screen } from '@testing-library/react'
import { describe, expect, it, vi } from 'vitest'

import { PromptPage } from './PromptPage'

function renderPromptPage(microsoftSupport: boolean) {
  return render(
    <PromptPage
      acceptAttr=".pdf"
      createMode={false}
      isUploadOnlySelecting={false}
      microsoftSupport={microsoftSupport}
      onOpenCommentaryPicker={vi.fn()}
      onOpenDiagramPicker={vi.fn()}
      onOpenInputPage={vi.fn()}
      onOpenJsonInput={vi.fn()}
      onCreateModeChange={vi.fn()}
      onRemoveUploadOnlyFile={vi.fn()}
      onResolveSharePoint={vi.fn()}
      onSharePointUrlChange={vi.fn()}
      onCloseSharePointPicker={vi.fn()}
      onUseSharePointFiles={vi.fn()}
      onUploadOnlyFileChange={vi.fn()}
      onUploadOnlySubmit={vi.fn()}
      selectedArchitectureDiagramId=""
      uploadOnlyFiles={[]}
      uploadOnlyFileCount={0}
      uploadOnlyError=""
      sharePointError="SharePoint error"
      sharePointIsDownloading={false}
      sharePointIsResolving={false}
      sharePointResource={null}
      sharePointUrl=""
    />,
  )
}

describe('PromptPage Microsoft features', () => {
  it('hides SharePoint controls when Microsoft support is disabled', () => {
    renderPromptPage(false)

    expect(screen.queryByLabelText('SharePoint file or folder link')).toBeNull()
    expect(screen.queryByRole('button', { name: 'Load SharePoint' })).toBeNull()
    expect(screen.queryByRole('alert')).toBeNull()
  })

  it('shows SharePoint controls when Microsoft support is enabled', () => {
    renderPromptPage(true)

    expect(screen.getByLabelText('SharePoint file or folder link')).not.toBeNull()
    expect(screen.getByRole('button', { name: 'Load SharePoint' })).not.toBeNull()
    expect(screen.getByRole('alert').textContent).toContain('SharePoint error')
  })
})
