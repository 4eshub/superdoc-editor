import { Plugin, PluginKey } from 'prosemirror-state'

const HIGHLIGHT_ONLY_PLUGIN_KEY = new PluginKey('superdoc-highlight-only')

export const HIGHLIGHT_SWATCHES = [
  { color: '#FFFF00', label: 'Yellow' },
  { color: '#00FF00', label: 'Green' },
  { color: '#00FFFF', label: 'Cyan' },
  { color: '#FF00FF', label: 'Pink' },
  { color: '#0000FF', label: 'Blue' },
  { color: '#FF0000', label: 'Red' },
  { color: '#FCFCFC', label: 'Light gray' },
] as const

type GuardableEditor = {
  registerPlugin?: (plugin: Plugin) => void
  unregisterPlugin?: (nameOrPluginKey: string | PluginKey) => void
  state?: { selection?: { empty?: boolean } }
  commands?: {
    setHighlight?: (color: string) => boolean
    unsetHighlight?: () => boolean
  }
  chain?: () => {
    focus?: () => {
      setHighlight?: (color: string) => { run?: () => boolean }
      unsetHighlight?: () => { run?: () => boolean }
    }
  }
}

export function wireHighlightOnlyGuard(editor: GuardableEditor | null | undefined) {
  if (!editor?.registerPlugin) return

  editor.unregisterPlugin?.(HIGHLIGHT_ONLY_PLUGIN_KEY)

  editor.registerPlugin(
    new Plugin({
      key: HIGHLIGHT_ONLY_PLUGIN_KEY,
      filterTransaction(tr) {
        if (!tr.docChanged) return true
        const before = tr.before
        if (!before) return false
        return before.textContent === tr.doc.textContent
      },
    })
  )
}

export function hasHighlightableSelection(editor: GuardableEditor | null | undefined): boolean {
  const selection = editor?.state?.selection
  return Boolean(selection && !selection.empty)
}

export function applyTextHighlight(
  editor: GuardableEditor | null | undefined,
  color: string | null
): boolean {
  if (!editor?.commands) return false

  const chain = editor.chain?.()?.focus?.()
  if (color) {
    if (chain?.setHighlight) return chain.setHighlight(color).run?.() === true
    return editor.commands.setHighlight?.(color) === true
  }

  if (chain?.unsetHighlight) return chain.unsetHighlight().run?.() === true
  return editor.commands.unsetHighlight?.() === true
}
