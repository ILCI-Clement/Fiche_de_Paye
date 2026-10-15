// Shared by half-day cells, the status selector and the calendar legend.
const definitions = {
  'Travail': { short: 'Travail', light: ['#e3edfc', '#214a84'], dark: ['#233853', '#bbd7ff'] },
  'Congé payé': { short: 'CP', light: ['#eee5fa', '#63368b'], dark: ['#433052', '#e1c5ff'] },
  'Absence': { short: 'ABS', light: ['#fde4e8', '#9e2943'], dark: ['#532b38', '#ffc2cf'] },
  'Arrêt maladie': { short: 'AM', light: ['#ffebdb', '#914415'], dark: ['#503929', '#ffd0a5'] },
  'Férié': { short: 'Férié', light: ['#fff2cb', '#70550b'], dark: ['#494126', '#ffe391'] },
  'Autre': { short: 'Autre', light: ['#e6e7ff', '#46459b'], dark: ['#343653', '#cacbff'] },
  'Repos': { short: 'Repos', light: ['#e9ebef', '#515967'], dark: ['#343740', '#c3c7d0'] },
} as const

export const calendarStatuses = Object.keys(definitions)

export function calendarStatus(status: string, mode: 'light' | 'dark') {
  const definition = definitions[status as keyof typeof definitions] ?? definitions.Repos
  const [background, foreground] = definition[mode]
  return { label: status, short: status in definitions ? definition.short : status, background, foreground }
}
