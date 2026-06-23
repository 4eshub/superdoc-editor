export const SUPERDOC_IFRAME_MESSAGE_NAMESPACE = 'superdoc-editor'

export const DOCX_MIME =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document'

export type SuperDocDocumentSource = string | File | Blob | null

export type SuperDocUser = {
  name?: string
  email?: string
}

export type SuperDocLabels = {
  openDocxFile: string
  differentFirstPage: string
  differentFirstPageActive: string
  editingFirstPageHeader: string
  editingFirstPageFooter: string
  editingOtherPagesHeader: string
  editingOtherPagesFooter: string
  pageBreak: string
}

export type SuperDocInitPayload = {
  document: SuperDocDocumentSource
  user?: SuperDocUser
  documentMode: string
  role: string
  hideToolbar: boolean
  showOpenDocx: boolean
  showDifferentFirstPage: boolean
  showPageBreak: boolean
  trackChangesVisible?: boolean
  canComment?: boolean
  locale?: string
  labels?: Partial<SuperDocLabels>
}

export type ParentToIframeMessage =
  | {
      namespace: typeof SUPERDOC_IFRAME_MESSAGE_NAMESPACE
      type: 'init'
      requestId: string
      payload: SuperDocInitPayload
    }
  | {
      namespace: typeof SUPERDOC_IFRAME_MESSAGE_NAMESPACE
      type: 'setDocument'
      requestId: string
      payload: Pick<SuperDocInitPayload, 'document'>
    }
  | {
      namespace: typeof SUPERDOC_IFRAME_MESSAGE_NAMESPACE
      type: 'exportDocx'
      requestId: string
      payload: { fileName: string }
    }
  | {
      namespace: typeof SUPERDOC_IFRAME_MESSAGE_NAMESPACE
      type: 'isEmpty'
      requestId: string
    }
  | {
      namespace: typeof SUPERDOC_IFRAME_MESSAGE_NAMESPACE
      type: 'runDiff'
      requestId: string
      payload: {
        original: SuperDocDocumentSource
        revised: SuperDocDocumentSource
        user: SuperDocUser
      }
    }
  | {
      namespace: typeof SUPERDOC_IFRAME_MESSAGE_NAMESPACE
      type: 'scrollToComment'
      requestId: string
      payload: { commentId: string }
    }

export type IframeToParentMessage =
  | {
      namespace: typeof SUPERDOC_IFRAME_MESSAGE_NAMESPACE
      type: 'ready'
      requestId?: string
    }
  | {
      namespace: typeof SUPERDOC_IFRAME_MESSAGE_NAMESPACE
      type: 'update'
      requestId?: string
    }
  | {
      namespace: typeof SUPERDOC_IFRAME_MESSAGE_NAMESPACE
      type: 'docxSelected'
      requestId?: string
      payload: { file: File }
    }
  | {
      namespace: typeof SUPERDOC_IFRAME_MESSAGE_NAMESPACE
      type: 'exportDocxResult'
      requestId: string
      payload: { file: File }
    }
  | {
      namespace: typeof SUPERDOC_IFRAME_MESSAGE_NAMESPACE
      type: 'isEmptyResult'
      requestId: string
      payload: { isEmpty: boolean }
    }
  | {
      namespace: typeof SUPERDOC_IFRAME_MESSAGE_NAMESPACE
      type: 'runDiffResult'
      requestId: string
      payload: { blob: Blob; hasChanges: boolean; changedComponents: string[] }
    }
  | {
      namespace: typeof SUPERDOC_IFRAME_MESSAGE_NAMESPACE
      type: 'scrollToCommentResult'
      requestId: string
      payload: { found: boolean }
    }
  | {
      namespace: typeof SUPERDOC_IFRAME_MESSAGE_NAMESPACE
      type: 'commentSelected'
      requestId?: string
      payload: { commentId: string }
    }
  | {
      namespace: typeof SUPERDOC_IFRAME_MESSAGE_NAMESPACE
      type: 'commentSaved'
      requestId?: string
      payload: { type: string; commentId?: string }
    }
  | {
      namespace: typeof SUPERDOC_IFRAME_MESSAGE_NAMESPACE
      type: 'error'
      requestId?: string
      payload: { message: string }
    }

export function isSuperDocIframeMessage(
  value: unknown,
): value is ParentToIframeMessage | IframeToParentMessage {
  return (
    typeof value === 'object' &&
    value !== null &&
    (value as { namespace?: unknown }).namespace === SUPERDOC_IFRAME_MESSAGE_NAMESPACE
  )
}
