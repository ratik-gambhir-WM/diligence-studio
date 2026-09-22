import { BrandLockup } from '../components/BrandLockup'
import { Button } from '../components/Button'
import { PageShell } from '../components/PageShell'
import { StudioPanel } from '../components/StudioPanel'

type LoginPageProps = {
  authError?: string
  onSignInWithMicrosoft: () => void
}

export function LoginPage({ authError, onSignInWithMicrosoft }: LoginPageProps) {
  return (
    <PageShell layout="center">
      <StudioPanel className="w-full max-w-[520px]">
        <BrandLockup className="mb-8" title="Diligence Studio" />

        <div className="grid gap-5">
          <div>
            <h1 className="text-[1.45rem] font-semibold text-white">Sign in to continue</h1>
            <p className="mt-2 text-[0.95rem] leading-6 text-[#a8afc4]">
              Use your West Monroe Microsoft account to access SharePoint source material.
            </p>
          </div>

          {authError && (
            <p className="text-[0.84rem] leading-5 text-[#ffb5b5]" role="alert">
              {authError}
            </p>
          )}

          <Button
            type="button"
            variant="primary"
            onClick={onSignInWithMicrosoft}
            className="mt-1 rounded-[1.15rem] px-5 py-3 text-[0.88rem] tracking-[0.08em]"
          >
            Sign in with Microsoft
          </Button>
        </div>
      </StudioPanel>
    </PageShell>
  )
}
