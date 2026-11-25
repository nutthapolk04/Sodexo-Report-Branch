import { ref } from 'vue'
import { messages, type Lang, type MsgKey } from '../i18n'

const STORAGE_KEY = 'dcbk.lang'

function loadLang(): Lang {
  try {
    const v = localStorage.getItem(STORAGE_KEY)
    if (v === 'th' || v === 'en') return v
  } catch {
    /* storage unavailable */
  }
  return 'th'
}

const lang = ref<Lang>(loadLang())
if (typeof document !== 'undefined') document.documentElement.lang = lang.value

function setLang(next: Lang): void {
  lang.value = next
  if (typeof document !== 'undefined') document.documentElement.lang = next
  try {
    localStorage.setItem(STORAGE_KEY, next)
  } catch {
    /* storage unavailable */
  }
}

type Params = Record<string, string | number>

/** Look up a key and fill {param} placeholders. Reads lang.value so callers re-render on switch. */
function t(key: MsgKey, params?: Params): string {
  const raw = messages[lang.value][key] ?? key
  return raw.replace(/\{(\w+)\}/g, (_m, k: string) => String(params?.[k] ?? ''))
}

export interface RichPart {
  text: string
  bold: boolean
}

/** Split a constant i18n string on <b>…</b> into safe text parts (no HTML injection). */
function richParts(s: string): RichPart[] {
  const out: RichPart[] = []
  const re = /<b>(.*?)<\/b>/g
  let last = 0
  for (let m = re.exec(s); m; m = re.exec(s)) {
    if (m.index > last) out.push({ text: s.slice(last, m.index), bold: false })
    out.push({ text: m[1] ?? '', bold: true })
    last = m.index + m[0].length
  }
  if (last < s.length) out.push({ text: s.slice(last), bold: false })
  return out
}

export function useI18n() {
  return { lang, setLang, t, richParts }
}
