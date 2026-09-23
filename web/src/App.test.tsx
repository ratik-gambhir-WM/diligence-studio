import { render, screen, waitFor } from '@testing-library/react'
import { MemoryRouter } from 'react-router-dom'
import { afterEach, describe, expect, it, vi } from 'vitest'

import App from './App'

afterEach(() => {
  vi.unstubAllGlobals()
})

describe('App when Microsoft support is disabled', () => {
  it('opens the app without checking for a Microsoft session', async () => {
    const fetchMock = vi.fn()
    vi.stubGlobal('fetch', fetchMock)

    render(
      <MemoryRouter initialEntries={['/']}>
        <App />
      </MemoryRouter>,
    )

    await waitFor(() => expect(screen.getByText('Upload file')).not.toBeNull())
    expect(fetchMock).not.toHaveBeenCalled()
  })

  it('redirects the login route to the anonymous app', async () => {
    const fetchMock = vi.fn()
    vi.stubGlobal('fetch', fetchMock)

    render(
      <MemoryRouter initialEntries={['/login']}>
        <App />
      </MemoryRouter>,
    )

    await waitFor(() => expect(screen.getByText('Upload file')).not.toBeNull())
    expect(screen.queryByRole('button', { name: 'Sign in with Microsoft' })).toBeNull()
    expect(fetchMock).not.toHaveBeenCalled()
  })
})
