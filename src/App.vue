<template>
  <main class="superdoc-host">
    <SuperDocEditor
      v-if="config"
      ref="editorRef"
      :key="editorKey"
      :document="config.document"
      :user="config.user"
      :document-mode="config.documentMode"
      :role="config.role"
      :hide-toolbar="config.hideToolbar"
      :show-open-docx="config.showOpenDocx"
      :show-different-first-page="config.showDifferentFirstPage"
      :show-page-break="config.showPageBreak"
      :show-page-number="config.showPageNumber ?? true"
      :track-changes-visible="config.trackChangesVisible ?? false"
      :can-comment="config.canComment ?? false"
      :show-highlight-toolbar="config.showHighlightToolbar ?? false"
      :labels="labels"
      @ready="postToParent({ type: 'ready' })"
      @update="postToParent({ type: 'update' })"
      @docx-selected="onDocxSelected"
      @comment-selected="onCommentSelected"
      @comment-saved="onCommentSaved"
      @highlight-saved="onHighlightSaved"
      @copy-comment-link-request="onCopyCommentLinkRequest"
      @copy-node-link-request="onCopyNodeLinkRequest"
    />
    <div v-else class="superdoc-waiting" role="status">Waiting for editor configuration...</div>
  </main>
</template>

<script setup lang="ts">
import { computed, onMounted, onUnmounted, ref } from 'vue'
import SuperDocEditor from '@/components/SuperDocEditor.vue'
import { resolveLabels } from '@/utils/labels'
import { buildRedlinedDocx } from '@/utils/superdoc-diff'
import {
  SUPERDOC_IFRAME_MESSAGE_NAMESPACE,
  isSuperDocIframeMessage,
  type IframeToParentMessage,
  type ParentToIframeMessage,
  type SuperDocInitPayload,
} from '@/types/messages'

type EditorComponent = InstanceType<typeof SuperDocEditor>
type WithoutNamespace<T> = T extends { namespace: typeof SUPERDOC_IFRAME_MESSAGE_NAMESPACE }
  ? Omit<T, 'namespace'>
  : never
type IframeOutboundMessage = WithoutNamespace<IframeToParentMessage>

const editorRef = ref<EditorComponent | null>(null)
const config = ref<SuperDocInitPayload | null>(null)
const editorKey = ref(0)
const parentOrigin = ref<string | null>(null)

const allowedParentOrigins = (import.meta.env.VITE_ALLOWED_PARENT_ORIGINS || '')
  .split(',')
  .map((origin: string) => origin.trim())
  .filter(Boolean)

const labels = computed(() => resolveLabels(config.value?.locale, config.value?.labels))

function canTrustOrigin(origin: string) {
  return allowedParentOrigins.length === 0 || allowedParentOrigins.includes(origin)
}

function rememberParentOrigin(origin: string) {
  if (!canTrustOrigin(origin)) return false
  if (!parentOrigin.value) {
    parentOrigin.value = origin
  }
  return parentOrigin.value === origin
}

function postToParent(
  message: IframeOutboundMessage,
  transfer?: Transferable[],
) {
  if (!parentOrigin.value) return
  window.parent.postMessage(
    {
      namespace: SUPERDOC_IFRAME_MESSAGE_NAMESPACE,
      ...message,
    } as IframeToParentMessage,
    parentOrigin.value,
    transfer ?? [],
  )
}

function postError(requestId: string | undefined, error: unknown) {
  const message = error instanceof Error ? error.message : String(error)
  postToParent({
    type: 'error',
    requestId,
    payload: { message },
  })
}

function onDocxSelected(file: File) {
  postToParent({
    type: 'docxSelected',
    payload: { file },
  })
}

function onCommentSelected(commentId: string) {
  postToParent({
    type: 'commentSelected',
    payload: { commentId },
  })
}

function onCommentSaved(payload: { type: string; commentId?: string }) {
  postToParent({
    type: 'commentSaved',
    payload,
  })
}

function onHighlightSaved() {
  postToParent({
    type: 'highlightSaved',
  })
}

function onCopyCommentLinkRequest(commentId: string) {
  postToParent({
    type: 'copyCommentLinkRequest',
    payload: { commentId },
  })
}

function onCopyNodeLinkRequest(payload: { nodeId: string; offset?: number }) {
  postToParent({
    type: 'copyNodeLinkRequest',
    payload,
  })
}

async function handleRequest(message: ParentToIframeMessage) {
  try {
    if (message.type === 'init') {
      config.value = message.payload
      editorKey.value += 1
      return
    }

    if (message.type === 'setDocument') {
      if (!config.value) throw new Error('Editor has not been initialized')
      config.value = {
        ...config.value,
        document: message.payload.document,
      }
      editorKey.value += 1
      return
    }

    if (message.type === 'runDiff') {
      const { original, revised, user } = message.payload
      if (!original || !revised) throw new Error('runDiff: original and revised documents are required')
      console.info('[superdoc-editor] runDiff message received', {
        requestId: message.requestId,
        origin: parentOrigin.value,
        originalType: original instanceof Blob ? 'Blob' : typeof original,
        revisedType: revised instanceof Blob ? 'Blob' : typeof revised,
      })
      console.info('[superdoc-editor] runDiff build started', { requestId: message.requestId })
      const result = await buildRedlinedDocx(
        original,
        revised,
        user as { name: string; email: string },
      )
      console.info('[superdoc-editor] runDiff build completed', {
        requestId: message.requestId,
        hasChanges: result.summary.hasChanges,
        changedComponents: result.summary.changedComponents,
        blobSize: result.blob.size,
      })
      postToParent({
        type: 'runDiffResult',
        requestId: message.requestId,
        payload: {
          blob: result.blob,
          hasChanges: result.summary.hasChanges,
          changedComponents: result.summary.changedComponents,
        },
      })
      return
    }

    if (!editorRef.value) throw new Error('Editor is not ready')

    if (message.type === 'exportDocx') {
      const file = await editorRef.value.exportDocx(message.payload.fileName)
      postToParent({
        type: 'exportDocxResult',
        requestId: message.requestId,
        payload: { file },
      })
      return
    }

    if (message.type === 'isEmpty') {
      postToParent({
        type: 'isEmptyResult',
        requestId: message.requestId,
        payload: { isEmpty: editorRef.value.isEmpty() },
      })
      return
    }

    if (message.type === 'hasComments') {
      postToParent({
        type: 'hasCommentsResult',
        requestId: message.requestId,
        payload: { hasComments: editorRef.value.hasComments() },
      })
      return
    }

    if (message.type === 'scrollToComment') {
      const found = await editorRef.value.scrollToComment(message.payload.commentId)
      postToParent({
        type: 'scrollToCommentResult',
        requestId: message.requestId,
        payload: { found },
      })
      return
    }

    if (message.type === 'scrollToElement') {
      const found = await editorRef.value.scrollToElement(
        message.payload.nodeId,
        message.payload.offset
      )
      postToParent({
        type: 'scrollToElementResult',
        requestId: message.requestId,
        payload: { found },
      })
    }
  } catch (error) {
    if (message.type === 'runDiff') {
      console.error('[superdoc-editor] runDiff failed', {
        requestId: message.requestId,
        error,
      })
    }
    postError(message.requestId, error)
  }
}

function onMessage(event: MessageEvent) {
  if (!isSuperDocIframeMessage(event.data)) return
  if (!rememberParentOrigin(event.origin)) return

  handleRequest(event.data as ParentToIframeMessage)
}

onMounted(() => {
  window.addEventListener('message', onMessage)
})

onUnmounted(() => {
  window.removeEventListener('message', onMessage)
})
</script>

<style>
html,
body,
#app {
  width: 100%;
  height: 100%;
  margin: 0;
}

body {
  overflow: hidden;
}
</style>

<style scoped>
.superdoc-host {
  width: 100%;
  height: 100%;
  min-height: 0;
  background: #ffffff;
}

.superdoc-waiting {
  display: flex;
  align-items: center;
  justify-content: center;
  width: 100%;
  height: 100%;
  color: #64748b;
  font-size: 0.875rem;
}
</style>
