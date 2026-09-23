import { useEffect, useState } from 'react'
import type { ChangeEvent } from 'react'
import { AnimatePresence, motion, useReducedMotion } from 'motion/react'
import { Navigate, Route, Routes, useLocation, useNavigate } from 'react-router-dom'

import type { CanvasTemplate } from './lib/canvas-model/templates'
import { getTemplate, listTemplates, type PickerTemplateSummary } from './lib/api/templateApi'
import {
  getSelectedArchitectureTemplate,
  selectArchitectureDiagramModel,
} from './lib/modelSelector'
import { generateSlidePromptOutput } from './lib/OpenAI'
import {
  downloadSharePointFile,
  resolveSharePointResource,
  type SharePointFile,
  type SharePointResolvedResource,
} from './lib/api/sharepointApi'
import { ACCEPT_ATTR, useDiagramSession } from './hooks/useDiagramSession'
import { JsonInputPage } from './pages/JsonInputPage'
import { LoginPage } from './pages/LoginPage'
import { PromptPage } from './pages/PromptPage'
import { CommentaryPicker } from './pages/CommentaryPicker'
import { SlidePickerPage } from './pages/SlidePickerPage'
import { TemplateCanvasPage } from './pages/TemplateCanvasPage'
import { SnailLoader } from './components/SnailLoader'
import {
  logAttachmentDebug,
  summarizeAttachments,
} from './lib/attachmentDiagnostics'
import type { ModelSelectorOutput } from './types/ModelSelectorOutput'
import { formatFileSize } from './utils/files'
import { MICROSOFT_SUPPORT } from './lib/microsoftSupport'

const EXPORTER_ROUTE = '/'
const COMMENTARY_PICKER_ROUTE = '/commentary-picker'
const DIAGRAM_PICKER_ROUTE = '/diagram-picker'
const DIAGRAM_CANVAS_ROUTE = '/diagram-template'
const JSON_INPUT_ROUTE = '/json-input'
const LOGIN_ROUTE = '/login'
type CanvasTemplateSource = 'commentary' | 'diagram'
type AuthSession = {
  email: string
  id: string
  name: string
}

export default function App() {
  const location = useLocation()
  const navigate = useNavigate()
  const shouldReduceMotion = useReducedMotion()
  const [session, setSession] = useState<AuthSession | null>(null)
  const [isSessionLoading, setIsSessionLoading] = useState(true)
  const [canvasTemplate, setCanvasTemplate] = useState<CanvasTemplate | null>(null)
  const [canvasTemplateSource, setCanvasTemplateSource] = useState<CanvasTemplateSource>('diagram')
  const [templateStatusMessage, setTemplateStatusMessage] = useState('')
  const [templateError, setTemplateError] = useState('')
  const [isTemplateSubmitting, setIsTemplateSubmitting] = useState(false)
  const [isCreateMode, setIsCreateMode] = useState(false)
  const [modelSelection, setModelSelection] = useState<ModelSelectorOutput | null>(null)
  const [modelSelectorError, setModelSelectorError] = useState('')
  const [isModelSelecting, setIsModelSelecting] = useState(false)
  const [isTemplateJsonOpenOnLoad, setIsTemplateJsonOpenOnLoad] = useState(false)
  const [sharePointUrl, setSharePointUrl] = useState('')
  const [sharePointResource, setSharePointResource] = useState<SharePointResolvedResource | null>(null)
  const [sharePointError, setSharePointError] = useState('')
  const [isSharePointResolving, setIsSharePointResolving] = useState(false)
  const [isSharePointDownloading, setIsSharePointDownloading] = useState(false)
  const {
    addSharePointAttachments,
    attachmentCountLabel,
    attachments,
    error,
    handleFiles,
    removeAttachment,
    removeUploadOnlyAttachment,
    uploadOnlyAttachmentItems,
    uploadOnlyAttachments,
  } = useDiagramSession()

  useEffect(() => {
    if (!MICROSOFT_SUPPORT) {
      setIsSessionLoading(false)
      return
    }

    let isActive = true

    fetch('/api/auth/me', { credentials: 'include' })
      .then(async (response) => {
        if (!response.ok) {
          return null
        }

        return (await response.json()) as { authenticated: boolean; user?: AuthSession }
      })
      .then((result) => {
        if (isActive && result?.authenticated && result.user) {
          setSession(result.user)
        }
      })
      .catch(() => undefined)
      .finally(() => {
        if (isActive) {
          setIsSessionLoading(false)
        }
      })

    return () => {
      isActive = false
    }
  }, [])

  function handleUploadOnlyFileChange(event: ChangeEvent<HTMLInputElement>) {
    const uploadedFiles = handleFiles(event, 'upload-only')

    logAttachmentDebug('upload-only-selection-handled', {
      accepted: summarizeAttachments(uploadedFiles),
    })

    if (uploadedFiles.length === 0) {
      return
    }

    setModelSelection(null)
    setModelSelectorError('')
    setSharePointResource(null)
    setSharePointError('')
  }

  async function downloadAndAddSharePointFiles(
    files: SharePointFile[],
    sourceUrl: string,
  ) {
    if (!MICROSOFT_SUPPORT) return

    if (files.length === 0) {
      setSharePointError('No supported files were found in that SharePoint resource.')
      return
    }

    setSharePointError('')
    setIsSharePointDownloading(true)

    try {
      const downloadedFiles = await mapWithConcurrency(files, 4, async (file) => ({
        file: await downloadSharePointFile(file),
        fileId: file.fileId,
        path: file.path,
        sourceUrl,
      }))
      addSharePointAttachments(downloadedFiles)
      setSharePointResource(null)
      setModelSelection(null)
      setModelSelectorError('')
    } catch (error) {
      setSharePointError(
        error instanceof Error
          ? error.message
          : 'SharePoint files could not be loaded.',
      )
    } finally {
      setIsSharePointDownloading(false)
    }
  }

  async function handleResolveSharePoint() {
    if (!MICROSOFT_SUPPORT) return

    const url = sharePointUrl.trim()
    if (!url) {
      setSharePointError('Paste a SharePoint file or folder link first.')
      return
    }

    setSharePointError('')
    setSharePointResource(null)
    setIsSharePointResolving(true)

    try {
      const resource = await resolveSharePointResource(url)
      if (resource.kind === 'file') {
        await downloadAndAddSharePointFiles(resource.files, resource.resourceUrl)
      } else {
        setSharePointResource(resource)
      }
    } catch (error) {
      setSharePointError(
        error instanceof Error
          ? error.message
          : 'The SharePoint link could not be resolved.',
      )
    } finally {
      setIsSharePointResolving(false)
    }
  }

  function handleUseSharePointFiles(files: SharePointFile[]) {
    if (!MICROSOFT_SUPPORT || !sharePointResource) return
    void downloadAndAddSharePointFiles(files, sharePointResource.resourceUrl)
  }

  async function handleUploadOnlySubmit() {
    const submittedAttachments = [...uploadOnlyAttachments]

    logAttachmentDebug('upload-only-submit-started', {
      createMode: isCreateMode,
      attachments: summarizeAttachments(submittedAttachments),
    })

    if (submittedAttachments.length === 0) {
      setModelSelectorError('Add at least one file before submitting.')
      return
    }

    setModelSelectorError('')
    setIsModelSelecting(true)

    try {
      const catalog = await listTemplates('diagram')
      if (isCreateMode) {
        const { generateArchitectureDiagramFromExamples } = await import('./lib/diagramGenerator')
        const candidates = await Promise.all(
          catalog.templates
            .filter((template) => template.previewUrl !== null)
            .map(async (template) => toCanvasTemplate(template, await getTemplate(template.templateId))),
        )
        const createdDiagramJson = await generateArchitectureDiagramFromExamples({
          candidates,
          uploadedFiles: submittedAttachments,
        })

        logAttachmentDebug('upload-only-create-completed', {
          attachments: summarizeAttachments(submittedAttachments),
          slideCount: createdDiagramJson.presentation.slides.length,
          elementCounts: createdDiagramJson.presentation.slides.map((slide) => slide.elements.length),
        })

        setModelSelection(null)
        setCanvasTemplate({
          id: 'created-architecture-diagram',
          name: createdDiagramJson.presentation.title || 'Created Architecture Diagram',
          description: 'A generated architecture diagram created from uploaded source material.',
          image: '',
          relatedAlt: 'Created architecture diagram',
          jsonSpec: createdDiagramJson,
        })
        setCanvasTemplateSource('diagram')
        setTemplateStatusMessage(
          `Created a new architecture diagram from ${uploadOnlyAttachments.length} uploaded file${
            uploadOnlyAttachments.length === 1 ? '' : 's'
          }.`,
        )
        setIsTemplateJsonOpenOnLoad(true)
        navigate(DIAGRAM_CANVAS_ROUTE)
        return
      }

      const selection = await selectArchitectureDiagramModel({
        candidates: catalog.templates,
        uploadedFiles: submittedAttachments,
      })
      const selectedTemplate = getSelectedArchitectureTemplate(selection, catalog.templates)

      if (!selectedTemplate) {
        throw new Error(`No template JSON found for selected diagram id: ${selection.selectedDiagramId}`)
      }

      const selectedTemplateJson = await getTemplate(selectedTemplate.templateId)
      const generatedSlideJson = await generateSlidePromptOutput({
        attachments: submittedAttachments,
        templateJson: selectedTemplateJson,
        prompt: [
          `Selected template: ${selectedTemplate.title}.`,
          'Use the attached technical context files to update the architecture diagram text.',
          'Keep the template layout and all non-text JSON values unchanged.',
        ].join(' '),
      })

      logAttachmentDebug('upload-only-submit-completed', {
        attachments: summarizeAttachments(submittedAttachments),
        slideCount: generatedSlideJson.presentation.slides.length,
        elementCounts: generatedSlideJson.presentation.slides.map((slide) => slide.elements.length),
      })

      setModelSelection(selection)
      setCanvasTemplate({
        ...toCanvasTemplate(selectedTemplate, generatedSlideJson),
        jsonSpec: generatedSlideJson,
      })
      setCanvasTemplateSource('diagram')
      setTemplateStatusMessage(
        `Generated ${selectedTemplate.title} from uploaded diligence material.`,
      )
      setIsTemplateJsonOpenOnLoad(true)
      navigate(DIAGRAM_CANVAS_ROUTE)
    } catch (selectionError) {
      setModelSelectorError(
        selectionError instanceof Error
          ? selectionError.message
          : isCreateMode
            ? 'Failed to create an architecture diagram.'
            : 'Failed to select an architecture diagram.',
      )
    } finally {
      setIsModelSelecting(false)
    }
  }

  async function handleSubmitTemplate(template: PickerTemplateSummary, signal: AbortSignal) {
    const submittedAttachments = [...attachments]

    logAttachmentDebug('template-submit-started', {
      attachments: summarizeAttachments(submittedAttachments),
    })

    setTemplateError('')
    setIsTemplateJsonOpenOnLoad(false)
    setIsTemplateSubmitting(true)

    try {
      const templateJson = await getTemplate(template.templateId, signal)
      const generatedSlideJson = await generateSlidePromptOutput({
        attachments: submittedAttachments,
        templateJson,
        prompt: [
          `Selected template: ${template.title}.`,
          'Use the attached technical context files to update the architecture diagram text.',
          'Keep the template layout and all non-text JSON values unchanged.',
        ].join(' '),
      })

      logAttachmentDebug('template-submit-completed', {
        attachments: summarizeAttachments(submittedAttachments),
        slideCount: generatedSlideJson.presentation.slides.length,
        elementCounts: generatedSlideJson.presentation.slides.map((slide) => slide.elements.length),
      })

      setCanvasTemplate({
        ...toCanvasTemplate(template, generatedSlideJson),
      })
      setCanvasTemplateSource('diagram')
      setTemplateStatusMessage(
        `Generated ${template.title} from ${
          attachments.length === 0
            ? 'the selected template.'
            : `${attachments.length} context file${attachments.length === 1 ? '' : 's'}.`
        }`,
      )
      navigate(DIAGRAM_CANVAS_ROUTE)
    } catch (submissionError) {
      setTemplateError(
        submissionError instanceof Error
          ? submissionError.message
          : 'Failed to generate the slide diagram JSON.',
      )
    } finally {
      setIsTemplateSubmitting(false)
    }
  }

  async function handleSelectCommentaryTemplate(template: PickerTemplateSummary, signal: AbortSignal) {
    setTemplateError('')
    setIsTemplateSubmitting(true)
    try {
      const templateJson = await getTemplate(template.templateId, signal)
      setCanvasTemplate(toCanvasTemplate(template, templateJson))
      setCanvasTemplateSource('commentary')
      setTemplateStatusMessage(`${template.title} is rendered from its stored template JSON.`)
      setIsTemplateJsonOpenOnLoad(false)
      navigate(DIAGRAM_CANVAS_ROUTE)
    } catch (selectionError) {
      const message = selectionError instanceof Error
        ? selectionError.message
        : 'Failed to load the commentary template.'
      setTemplateError(message)
      throw selectionError
    } finally {
      setIsTemplateSubmitting(false)
    }
  }

  function handleSignInWithMicrosoft() {
    window.location.assign('/api/auth/login')
  }

  if (isSessionLoading) {
    return <SnailLoader />
  }

  return (
    <div className="min-h-screen overflow-x-hidden bg-[#070a1b]">
      <AnimatePresence mode="popLayout" initial={false}>
        <motion.div
          key={location.pathname}
          className="min-h-screen bg-[#070a1b] will-change-transform"
          initial={shouldReduceMotion ? { opacity: 0 } : { opacity: 0, x: 140 }}
          animate={shouldReduceMotion ? { opacity: 1 } : { opacity: 1, x: 0 }}
          exit={shouldReduceMotion ? { opacity: 0 } : { opacity: 0, x: -140 }}
          transition={{ duration: shouldReduceMotion ? 0.12 : 0.46, ease: [0.22, 1, 0.36, 1] }}
        >
          <Routes location={location}>
            <Route
              path={LOGIN_ROUTE}
              element={
                !MICROSOFT_SUPPORT ? (
                  <Navigate replace to={EXPORTER_ROUTE} />
                ) : session ? (
                  <Navigate replace to={EXPORTER_ROUTE} />
                ) : (
                  <LoginPage
                    authError={new URLSearchParams(location.search).get('authError') ?? ''}
                    onSignInWithMicrosoft={handleSignInWithMicrosoft}
                  />
                )
              }
            />
            <Route
              path={EXPORTER_ROUTE}
              element={
                !MICROSOFT_SUPPORT || session ? (
                  <PromptPage
                    acceptAttr={ACCEPT_ATTR}
                    createMode={isCreateMode}
                    isUploadOnlySelecting={isModelSelecting}
                    microsoftSupport={MICROSOFT_SUPPORT}
                    onOpenCommentaryPicker={() => navigate(COMMENTARY_PICKER_ROUTE)}
                    onCreateModeChange={setIsCreateMode}
                    onOpenDiagramPicker={() => navigate(DIAGRAM_PICKER_ROUTE)}
                    onOpenInputPage={() => navigate(EXPORTER_ROUTE)}
                    onOpenJsonInput={() => navigate(JSON_INPUT_ROUTE)}
                    onRemoveUploadOnlyFile={removeUploadOnlyAttachment}
                    onResolveSharePoint={() => { void handleResolveSharePoint() }}
                    onSharePointUrlChange={(url) => {
                      setSharePointUrl(url)
                      setSharePointError('')
                    }}
                    onCloseSharePointPicker={() => {
                      if (!isSharePointDownloading) {
                        setSharePointResource(null)
                      }
                    }}
                    onUseSharePointFiles={handleUseSharePointFiles}
                    onUploadOnlyFileChange={handleUploadOnlyFileChange}
                    onUploadOnlySubmit={handleUploadOnlySubmit}
                    selectedArchitectureDiagramId={modelSelection?.selectedDiagramId ?? ''}
                    uploadOnlyFiles={uploadOnlyAttachmentItems}
                    uploadOnlyFileCount={uploadOnlyAttachments.length}
                    uploadOnlyError={modelSelectorError || error}
                    sharePointError={sharePointError}
                    sharePointIsDownloading={isSharePointDownloading}
                    sharePointIsResolving={isSharePointResolving}
                    sharePointResource={sharePointResource}
                    sharePointUrl={sharePointUrl}
                  />
                ) : (
                  <Navigate replace to={LOGIN_ROUTE} />
                )
              }
            />
            <Route
              path={DIAGRAM_PICKER_ROUTE}
              element={
                !MICROSOFT_SUPPORT || session ? (
                  <SlidePickerPage
                    acceptAttr={ACCEPT_ATTR}
                    attachmentCountLabel={attachmentCountLabel}
                    attachments={attachments}
                    error={templateError}
                    isSubmitting={isTemplateSubmitting}
                    onFileChange={handleFiles}
                    onOpenCommentaryPicker={() => navigate(COMMENTARY_PICKER_ROUTE)}
                    onOpenInputPage={() => navigate(EXPORTER_ROUTE)}
                    onOpenJsonInput={() => navigate(JSON_INPUT_ROUTE)}
                    onRemoveAttachment={removeAttachment}
                    onSelectTemplate={handleSubmitTemplate}
                    renderFileSize={formatFileSize}
                  />
                ) : (
                  <Navigate replace to={LOGIN_ROUTE} />
                )
              }
            />
            <Route
              path={COMMENTARY_PICKER_ROUTE}
              element={
                !MICROSOFT_SUPPORT || session ? (
                  <CommentaryPicker
                    error={templateError}
                    isSubmitting={isTemplateSubmitting}
                    onOpenInputPage={() => navigate(EXPORTER_ROUTE)}
                    onOpenJsonInput={() => navigate(JSON_INPUT_ROUTE)}
                    onSelectTemplate={handleSelectCommentaryTemplate}
                  />
                ) : (
                  <Navigate replace to={LOGIN_ROUTE} />
                )
              }
            />
            <Route
              path={JSON_INPUT_ROUTE}
              element={
                !MICROSOFT_SUPPORT || session ? (
                  <JsonInputPage
                    onOpenCommentaryPicker={() => navigate(COMMENTARY_PICKER_ROUTE)}
                    onOpenDiagramPicker={() => navigate(DIAGRAM_PICKER_ROUTE)}
                    onOpenInputPage={() => navigate(EXPORTER_ROUTE)}
                    onOpenJsonInput={() => navigate(JSON_INPUT_ROUTE)}
                  />
                ) : (
                  <Navigate replace to={LOGIN_ROUTE} />
                )
              }
            />
            <Route
              path={DIAGRAM_CANVAS_ROUTE}
              element={
                !MICROSOFT_SUPPORT || session ? (
                  canvasTemplate ? <TemplateCanvasPage
                    template={canvasTemplate}
                    statusMessage={templateStatusMessage}
                    showJsonByDefault={isTemplateJsonOpenOnLoad}
                    onTemplateJsonChange={(jsonSpec) =>
                      setCanvasTemplate((currentTemplate) => currentTemplate
                        ? { ...currentTemplate, jsonSpec }
                        : currentTemplate)
                    }
                    onOpenPicker={() =>
                      navigate(
                        canvasTemplateSource === 'commentary'
                          ? COMMENTARY_PICKER_ROUTE
                          : DIAGRAM_PICKER_ROUTE,
                      )
                    }
                    onOpenInputPage={() => navigate(EXPORTER_ROUTE)}
                  /> : <Navigate replace to={DIAGRAM_PICKER_ROUTE} />
                ) : (
                  <Navigate replace to={LOGIN_ROUTE} />
                )
              }
            />
          </Routes>
        </motion.div>
      </AnimatePresence>
    </div>
  )
}

function toCanvasTemplate(
  template: PickerTemplateSummary,
  jsonSpec: CanvasTemplate['jsonSpec'],
): CanvasTemplate {
  return {
    description: template.description,
    id: template.templateId,
    image: template.previewUrl ?? '',
    jsonSpec,
    name: template.title,
    relatedAlt: `${template.title} template preview`,
  }
}

async function mapWithConcurrency<Input, Output>(
  values: Input[],
  concurrency: number,
  mapper: (value: Input) => Promise<Output>,
) {
  const results = new Array<Output>(values.length)
  let nextIndex = 0

  async function worker() {
    while (nextIndex < values.length) {
      const index = nextIndex
      nextIndex += 1
      results[index] = await mapper(values[index])
    }
  }

  await Promise.all(
    Array.from({ length: Math.min(concurrency, values.length) }, () => worker()),
  )
  return results
}
