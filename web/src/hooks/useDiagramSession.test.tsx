import { useEffect } from 'react'
import { render, screen, waitFor } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { describe, expect, it } from 'vitest'

import { useDiagramSession } from './useDiagramSession'

function AttachmentHarness({ onObserved }: { onObserved: (files: File[]) => void }) {
  const { handleFiles, uploadOnlyAttachments } = useDiagramSession()

  useEffect(() => {
    onObserved(uploadOnlyAttachments)
  }, [onObserved, uploadOnlyAttachments])

  return (
    <input
      aria-label="Upload source files"
      type="file"
      onChange={(event) => handleFiles(event, 'upload-only')}
    />
  )
}

describe('useDiagramSession', () => {
  it('preserves the selected File object for the upload-only OpenAI flow', async () => {
    const user = userEvent.setup()
    let observedFiles: File[] = []
    const sourceFile = new File(['source material'], 'source.pdf', { type: 'application/pdf' })

    render(<AttachmentHarness onObserved={(files) => { observedFiles = files }} />)

    await user.upload(screen.getByLabelText('Upload source files'), sourceFile)

    await waitFor(() => expect(observedFiles).toHaveLength(1))
    expect(observedFiles[0]).toBe(sourceFile)
  })
})
