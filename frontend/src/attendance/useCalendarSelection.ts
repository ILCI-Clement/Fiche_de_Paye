import { useEffect, useRef, useState } from 'react'
import type { MouseEvent, PointerEvent } from 'react'
import type { Cell } from './types'

type Stroke = { pointerId: number; select: boolean; visited: Set<string>; x: number; y: number }
const keyOf = (cell: Cell) => `${cell.date}:${cell.period}`

/** Paint visited half-days, matching the original calendar's add/remove gesture. */
export function useCalendarSelection(disabled: boolean) {
  const [selection, setSelection] = useState<Cell[]>([])
  const stroke = useRef<Stroke | null>(null)

  useEffect(() => {
    const stop = () => { stroke.current = null }
    window.addEventListener('blur', stop)
    return () => window.removeEventListener('blur', stop)
  }, [])
  useEffect(() => { if (disabled) stroke.current = null }, [disabled])

  function paint(cell: Cell, gesture: Stroke) {
    const key = keyOf(cell)
    if (gesture.visited.has(key)) return
    gesture.visited.add(key)
    setSelection(previous => {
      const exists = previous.some(item => keyOf(item) === key)
      if (exists === gesture.select) return previous
      return gesture.select ? [...previous, cell] : previous.filter(item => keyOf(item) !== key)
    })
  }

  function start(event: PointerEvent<HTMLButtonElement>, cell: Cell) {
    if (disabled || !event.isPrimary || event.button !== 0 || stroke.current) return
    event.preventDefault()
    event.currentTarget.focus({ preventScroll: true })
    const gesture = { pointerId: event.pointerId, select: !selection.some(item => keyOf(item) === keyOf(cell)), visited: new Set<string>(), x: event.clientX, y: event.clientY }
    stroke.current = gesture
    event.currentTarget.setPointerCapture(event.pointerId)
    paint(cell, gesture)
  }

  function move(event: PointerEvent<HTMLButtonElement>) {
    const gesture = stroke.current
    if (!gesture || gesture.pointerId !== event.pointerId || disabled) return
    const grid = event.currentTarget.closest('.attendance-calendar-grid')
    if (!grid) return
    // Fill the pointer's path so a quick stroke does not skip cells between move events.
    const steps = Math.min(512, Math.max(1, Math.ceil(Math.hypot(event.clientX - gesture.x, event.clientY - gesture.y) / 8)))
    for (let step = 1; step <= steps; step++) {
      const x = gesture.x + (event.clientX - gesture.x) * step / steps
      const y = gesture.y + (event.clientY - gesture.y) * step / steps
      const hit = document.elementFromPoint(x, y)?.closest<HTMLButtonElement>('button[data-calendar-date]')
      if (!hit || !grid.contains(hit) || hit.disabled) continue
      const date = hit.dataset.calendarDate
      const period = hit.dataset.calendarPeriod
      if (date && (period === 'matin' || period === 'aprem')) paint({ date, period }, gesture)
    }
    gesture.x = event.clientX; gesture.y = event.clientY
  }

  function stop(event: PointerEvent<HTMLButtonElement>) {
    if (stroke.current?.pointerId !== event.pointerId) return
    stroke.current = null
    if (event.currentTarget.hasPointerCapture(event.pointerId)) event.currentTarget.releasePointerCapture(event.pointerId)
  }

  function cellProps(cell: Cell) {
    return {
      'data-calendar-date': cell.date,
      'data-calendar-period': cell.period,
      onPointerDown: (event: PointerEvent<HTMLButtonElement>) => start(event, cell),
      onPointerMove: move,
      onPointerUp: stop,
      onPointerCancel: stop,
      onLostPointerCapture: () => { stroke.current = null },
      onClick: (event: MouseEvent<HTMLButtonElement>) => {
        // Pointer input was already applied on down; keyboard/assistive clicks still toggle.
        if (disabled || event.detail !== 0) return
        setSelection(previous => previous.some(item => keyOf(item) === keyOf(cell))
          ? previous.filter(item => keyOf(item) !== keyOf(cell)) : [...previous, cell])
      },
    }
  }

  function clear() { stroke.current = null; setSelection([]) }
  return { selection, clear, cellProps }
}
