import { useState } from 'react'

import { MICROSOFT_SUPPORT } from '../lib/microsoftSupport'
import { WestMonroeMark } from './WestMonroeMark'
import { classNames } from './classNames'

type AppNavPage = 'diagramming' | 'commentary' | 'json-input'

type AppNavProps = {
  activePage: AppNavPage
  onOpenCommentaryPicker?: () => void
  onOpenDiagramPicker?: () => void
  onOpenInputPage?: () => void
  onOpenJsonInput?: () => void
}

const navItems = [
  { id: 'diagramming', label: 'Diagramming' },
  { id: 'commentary', label: 'Commentary' },
  { id: 'json-input', label: 'JSON Input' },
  { id: 'industries', label: 'Industries' },
  { id: 'about', label: 'About' },
  { id: 'how-to', label: 'How-to' },
] as const

export function AppNav({
  activePage,
  onOpenCommentaryPicker,
  onOpenDiagramPicker,
  onOpenInputPage,
  onOpenJsonInput,
}: AppNavProps) {
  const [isLoggingOut, setIsLoggingOut] = useState(false)
  const [logoutError, setLogoutError] = useState('')

  async function handleLogout() {
    setIsLoggingOut(true)
    setLogoutError('')

    try {
      const response = await fetch('/api/auth/logout', {
        credentials: 'include',
        method: 'POST',
      })

      if (!response.ok) {
        throw new Error('Logout failed.')
      }

      window.location.assign('/login')
    } catch {
      setIsLoggingOut(false)
      setLogoutError('Unable to log out. Try again.')
    }
  }

  function getNavAction(itemId: (typeof navItems)[number]['id']) {
    if (itemId === 'diagramming') {
      return onOpenInputPage
    }

    if (itemId === 'commentary') {
      return onOpenCommentaryPicker ?? onOpenDiagramPicker
    }

    if (itemId === 'json-input') {
      return onOpenJsonInput
    }

    return undefined
  }

  return (
    <nav className="app-nav" aria-label="Main navigation">
      <button
        type="button"
        onClick={onOpenInputPage}
        className="app-nav-brand"
        aria-label="Open WM Diligence Studio"
      >
        <WestMonroeMark className="app-nav-mark" />
        <span className="app-nav-brand-text">WM Diligence Studio</span>
      </button>

      <div className="app-nav-items">
        {navItems.map((item) => {
          const isActive = item.id === activePage
          const action = getNavAction(item.id)
          const className = classNames(
            'app-nav-item',
            isActive && 'app-nav-active',
            action && !isActive && 'app-nav-action',
          )

          if (isActive) {
            return (
              <span key={item.id} className={className} aria-current="page">
                {item.label}
              </span>
            )
          }

          if (action) {
            return (
              <button key={item.id} type="button" onClick={action} className={className}>
                {item.label}
              </button>
            )
          }

          return (
            <a key={item.id} className={className} href={`#${item.id}`}>
              {item.label}
            </a>
          )
        })}
      </div>
      {logoutError && (
        <span className="app-nav-logout-error" role="alert">
          {logoutError}
        </span>
      )}
      {MICROSOFT_SUPPORT && (
        <button
          type="button"
          className="app-nav-logout-button"
          onClick={handleLogout}
          disabled={isLoggingOut}
        >
          {isLoggingOut ? 'Logging out...' : 'Log out'}
        </button>
      )}
      <button
        type="button"
        className="app-nav-menu-button"
        aria-label="Open navigation menu"
      >
        <span className="app-nav-menu-icon" aria-hidden="true">
          ≡
        </span>
      </button>
    </nav>
  )
}
