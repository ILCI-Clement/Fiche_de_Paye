import { useEffect, type RefObject } from 'react'

/** Scope all focus work to the intended form; never select another person's dialog. */
export function useCompletionFocus(root: RefObject<HTMLElement | null>, field?: string, sequence?: number) {
  useEffect(() => {
    if (!field) return
    let frame = 0
    let highlighted: HTMLElement | undefined
    let clearHighlight = 0
    const deadline = performance.now() + 5000
    const focus = () => {
      const container = root.current
      const target = Array.from(container?.querySelectorAll<HTMLElement>('[data-completion-field]') || [])
        .find(element => element.dataset.completionField === field)
      const input = target?.querySelector<HTMLElement>('input:not([type="hidden"]):not(:disabled), textarea:not(:disabled), [role="combobox"], [role="spinbutton"], button:not(:disabled)')
      if (!target || !input || !input.getClientRects().length) {
        if (performance.now() < deadline) frame = requestAnimationFrame(focus)
        return
      }
      target.scrollIntoView({ block: 'center', behavior: 'auto' })
      input.focus({ preventScroll: true })
      target.classList.add('completion-focus')
      highlighted = target
      clearHighlight = window.setTimeout(() => target.classList.remove('completion-focus'), 2000)
    }
    frame = requestAnimationFrame(focus)
    return () => { cancelAnimationFrame(frame); clearTimeout(clearHighlight); highlighted?.classList.remove('completion-focus') }
  }, [root, field, sequence])
}
