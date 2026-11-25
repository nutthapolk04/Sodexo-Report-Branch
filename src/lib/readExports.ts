import * as XLSX from 'xlsx'
import {
  CARD_COLUMNS, ITEM_COLUMNS, TEMPLATE_START, TOPUP_COLUMNS,
} from '../types'
import type {
  Cell, ErrorCode, Issue, ParsedSlot, ReadOptions, Row, TableKey, WarningCode,
} from '../types'

export interface InputFile { name: string; data: ArrayBuffer }

type Raw = unknown[][]
type FileResult = {
  rows: Row[]
  org: string | null
  from: string | null
  to: string | null
  errors: Issue<ErrorCode>[]
  warnings: Issue<WarningCode>[]
}

const HEADER_SCAN = 15
const DAY = 86400000

// ---------- helpers ----------

const norm = (v: unknown): string => String(v ?? '').replace(/ /g, ' ').trim()
const normCard = (v: unknown): string => norm(v).replace(/\s*\*\s*$/, '').trim()

function toCell(v: unknown): Cell {
  if (typeof v === 'number' || typeof v === 'string') return v
  if (v === null || v === undefined) return ''
  return String(v)
}

/** dd/mm/yyyy -> UTC ms, or null. */
function parseDmy(s: unknown): number | null {
  const m = /^\s*(\d{1,2})\/(\d{1,2})\/(\d{4})/.exec(String(s ?? ''))
  if (!m) return null
  return Date.UTC(Number(m[3]), Number(m[2]) - 1, Number(m[1]))
}

function fmtDmy(ms: number): string {
  const d = new Date(ms)
  const p = (n: number): string => String(n).padStart(2, '0')
  return `${p(d.getUTCDate())}/${p(d.getUTCMonth() + 1)}/${d.getUTCFullYear()}`
}

function parseTime(s: unknown): number {
  const m = /(\d{1,2}):(\d{2})(?::(\d{2}))?/.exec(String(s ?? ''))
  return m ? Number(m[1]) * 3600 + Number(m[2]) * 60 + Number(m[3] ?? 0) : 0
}

function readRaw(data: ArrayBuffer): Raw {
  // SheetJS would happily parse arbitrary text as CSV; require a zip (xlsx) or OLE (xls) signature.
  const b = new Uint8Array(data.slice(0, 4))
  const isZip = b[0] === 0x50 && b[1] === 0x4b
  const isOle = b[0] === 0xd0 && b[1] === 0xcf && b[2] === 0x11 && b[3] === 0xe0
  if (!isZip && !isOle) throw new Error('not a workbook')
  const wb = XLSX.read(data, { type: 'array' })
  const name = wb.SheetNames[0]
  if (name === undefined) throw new Error('no sheet')
  const ws = wb.Sheets[name]
  if (!ws) throw new Error('no sheet')
  return XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, raw: true, defval: '' })
}

const ORG_RE = /^\s*(?:สาขา\/องค์กร|องค์กร\/สาขา|Org\/Branch|Organization\/Branch)\s*:\s*(.*)$/i

function findOrg(raw: Raw): string | null {
  for (let i = 0; i < Math.min(raw.length, HEADER_SCAN); i++) {
    const row = raw[i] ?? []
    for (const c of row) {
      if (typeof c !== 'string') continue
      const m = ORG_RE.exec(c.replace(/ /g, ' '))
      if (m) return (m[1] ?? '').trim() || null
    }
  }
  return null
}

function findHeaderRow(raw: Raw, pred: (cells: string[]) => boolean, norm_: (v: unknown) => string = norm): number {
  for (let i = 0; i < Math.min(raw.length, HEADER_SCAN); i++) {
    const cells = (raw[i] ?? []).map(norm_)
    if (pred(cells)) return i
  }
  return -1
}

function isEmptyRow(r: unknown[]): boolean {
  return r.every((c) => norm(c) === '')
}

function range(dates: (number | null)[]): { from: number | null; to: number | null } {
  let from: number | null = null
  let to: number | null = null
  for (const d of dates) {
    if (d === null) continue
    if (from === null || d < from) from = d
    if (to === null || d > to) to = d
  }
  return { from, to }
}

const empty = (): FileResult => ({ rows: [], org: null, from: null, to: null, errors: [], warnings: [] })

// ---------- per-kind file parsers ----------

const OPTIONAL_ITEM = ['เคาน์เตอร์', 'หมายเหตุการชำระเงิน', 'ประเภทผู้ใช้']

function parseItems(raw: Raw, file: string): FileResult {
  const res = empty()
  res.org = findOrg(raw)

  // English export
  if (findHeaderRow(raw, (c) => c.includes('Receipt No.') && (c.includes('Type') || c.includes('Date'))) >= 0) {
    res.errors.push({ code: 'SALES_NOT_THAI', params: { file } })
    return res
  }
  const h = findHeaderRow(raw, (c) => c.includes('วันที่') && c.includes('เลขที่ใบเสร็จ'))
  if (h < 0) {
    res.errors.push({ code: 'HEADER_NOT_FOUND', params: { file, missing: 'วันที่, เลขที่ใบเสร็จ' } })
    return res
  }
  const header = (raw[h] ?? []).map(norm)
  const dateIdx = header.indexOf('วันที่')
  const receiptIdx = header.indexOf('เลขที่ใบเสร็จ')
  const typeBeforeDate = header.some((c, i) => c === 'ประเภท' && i < dateIdx)
  if (!typeBeforeDate) {
    res.errors.push({ code: 'BILLS_NOT_TRANSACTIONS', params: { file } })
    return res
  }

  const colOf = new Map<string, number>()
  header.forEach((c, i) => {
    if (c === '') return
    const key = c === 'ประเภท' ? (i < receiptIdx ? 'ประเภท' : 'ประเภท2') : c
    if (!colOf.has(key)) colOf.set(key, i)
  })

  const missingOpt: string[] = []
  const missing: string[] = []
  for (const col of ITEM_COLUMNS) {
    if (colOf.has(col)) continue
    if (OPTIONAL_ITEM.includes(col)) missingOpt.push(col)
    else missing.push(col)
  }
  if (missing.length > 0) {
    res.errors.push({ code: 'HEADER_NOT_FOUND', params: { file, missing: missing.join(', ') } })
    return res
  }
  if (missingOpt.length > 0) {
    res.warnings.push({ code: 'MISSING_OPTIONAL_COLUMNS', params: { file, columns: missingOpt.join(', ') } })
  }

  const dates: (number | null)[] = []
  for (let i = h + 1; i < raw.length; i++) {
    const r = raw[i] ?? []
    if (isEmptyRow(r)) continue
    res.rows.push(ITEM_COLUMNS.map((col) => {
      const idx = colOf.get(col)
      return idx === undefined ? '-' : toCell(r[idx])
    }))
    dates.push(parseDmy(r[dateIdx]))
  }
  const rg = range(dates)
  res.from = rg.from === null ? null : fmtDmy(rg.from)
  res.to = rg.to === null ? null : fmtDmy(rg.to)
  return res
}

function parseTopups(raw: Raw, file: string): FileResult {
  const res = empty()
  res.org = findOrg(raw)

  if (findHeaderRow(raw, (c) => c.includes('วันเวลาที่ชำระ') || (c.includes('เลขที่ใบเสร็จ') && c.includes('เลขบัตร'))) >= 0) {
    res.errors.push({ code: 'TOPUP_NOT_ENGLISH', params: { file } })
    return res
  }
  const h = findHeaderRow(raw, (c) => c.includes('Date/Time') && c.includes('Receipt No.'))
  if (h < 0) {
    res.errors.push({ code: 'HEADER_NOT_FOUND', params: { file, missing: TOPUP_COLUMNS.join(', ') } })
    return res
  }
  const header = (raw[h] ?? []).map(norm)
  const missing = TOPUP_COLUMNS.filter((c) => !header.includes(c))
  if (missing.length > 0) {
    res.errors.push({ code: 'HEADER_NOT_FOUND', params: { file, missing: missing.join(', ') } })
    return res
  }
  const idxs = TOPUP_COLUMNS.map((c) => header.indexOf(c))
  const dates: (number | null)[] = []
  for (let i = h + 1; i < raw.length; i++) {
    const r = raw[i] ?? []
    if (isEmptyRow(r)) continue
    const row = idxs.map((ix) => toCell(r[ix]))
    res.rows.push(row)
    dates.push(parseDmy(String(row[0]).split(',')[0]))
  }
  const rg = range(dates)
  res.from = rg.from === null ? null : fmtDmy(rg.from)
  res.to = rg.to === null ? null : fmtDmy(rg.to)
  return res
}

const MEMBER_MAP: [string, number][] = [
  ['รหัสสมาชิก', 0], ['ชื่อ (TH)', 1], ['นามสกุล (TH)', 2], ['ชื่อ (EN)', 3], ['นามสกุล (EN)', 4],
  ['ประเภท', 10], ['สถานะ', 11], ['เลขบัตร', 12], ['วันหมดอายุบัตร', 13],
]

function parseCards(raw: Raw, file: string): FileResult {
  const res = empty()
  const h = findHeaderRow(raw, (c) => c.includes('ชื่อ (TH)') && c.includes('นามสกุล (TH)'), normCard)
  if (h < 0) {
    res.errors.push({ code: 'HEADER_NOT_FOUND', params: { file, missing: CARD_COLUMNS.join(', ') } })
    return res
  }
  const header = (raw[h] ?? []).map(normCard)
  const target = CARD_COLUMNS.map(normCard)

  let idxs: number[] // per target column; -1 = fill ''
  if (header.includes(target[0] ?? '')) {
    const missing = CARD_COLUMNS.filter((_, i) => !header.includes(target[i] ?? ''))
    if (missing.length > 0) {
      res.errors.push({ code: 'HEADER_NOT_FOUND', params: { file, missing: missing.join(', ') } })
      return res
    }
    idxs = target.map((t) => header.indexOf(t))
  } else if (header.includes('รหัสสมาชิก')) {
    const missing = MEMBER_MAP.filter(([n]) => !header.includes(n)).map(([n]) => n)
    if (missing.length > 0) {
      res.errors.push({ code: 'HEADER_NOT_FOUND', params: { file, missing: missing.join(', ') } })
      return res
    }
    idxs = CARD_COLUMNS.map(() => -1)
    for (const [n, t] of MEMBER_MAP) idxs[t] = header.indexOf(n)
    res.warnings.push({ code: 'MEMBER_FILE_MAPPED', params: { file } })
  } else {
    res.errors.push({ code: 'HEADER_NOT_FOUND', params: { file, missing: CARD_COLUMNS[0] } })
    return res
  }

  for (let i = h + 1; i < raw.length; i++) {
    const r = raw[i] ?? []
    if (isEmptyRow(r)) continue
    res.rows.push(idxs.map((ix) => (ix < 0 ? '' : toCell(r[ix]))))
  }
  return res
}

// ---------- slot level ----------

function sortKey(kind: TableKey, row: Row): number {
  if (kind === 'items') {
    const d = parseDmy(row[1])
    return (d ?? Number.MAX_SAFE_INTEGER / 2) + parseTime(row[2]) * 1000
  }
  const s = String(row[0])
  const d = parseDmy(s.split(',')[0])
  return (d ?? Number.MAX_SAFE_INTEGER / 2) + parseTime(s.split(',')[1] ?? '') * 1000
}

function fullName(a: Cell | undefined, b: Cell | undefined): string {
  return `${norm(a)} ${norm(b)}`.trim()
}

function duplicateNames(rows: Row[]): string[] {
  const out: string[] = []
  for (const [a, b] of [[1, 2], [3, 4]] as const) {
    const count = new Map<string, number>()
    for (const r of rows) {
      const n = fullName(r[a], r[b])
      if (n === '') continue
      count.set(n, (count.get(n) ?? 0) + 1)
    }
    for (const [n, c] of count) if (c > 1 && !out.includes(n)) out.push(n)
  }
  return out
}

export function readSlot(kind: TableKey, files: InputFile[], opts: ReadOptions): ParsedSlot {
  const slot: ParsedSlot = {
    kind, files: files.map((f) => f.name), rows: [], org: null, from: null, to: null,
    duplicatesRemoved: 0, errors: [], warnings: [],
  }
  const results: { name: string; res: FileResult }[] = []
  for (const f of files) {
    let raw: Raw
    try {
      raw = readRaw(f.data)
    } catch {
      slot.errors.push({ code: 'NOT_XLSX', params: { file: f.name } })
      continue
    }
    const res = kind === 'items' ? parseItems(raw, f.name) : kind === 'topups' ? parseTopups(raw, f.name) : parseCards(raw, f.name)
    slot.errors.push(...res.errors)
    slot.warnings.push(...res.warnings)
    if (res.errors.length === 0) results.push({ name: f.name, res })
  }

  // org
  const orgs: string[] = []
  for (const { res } of results) if (res.org !== null && !orgs.includes(res.org)) orgs.push(res.org)
  slot.org = orgs[0] ?? null
  if (orgs.length > 1) slot.errors.push({ code: 'ORG_MISMATCH', params: { orgs: orgs.join(', ') } })

  // merge
  let rows: Row[] = results.flatMap(({ res }) => res.rows)
  if (opts.replaceEmDash) {
    rows = rows.map((r) => r.map((c) => (typeof c === 'string' ? c.replace(/—/g, '-') : c)))
  }
  if (kind !== 'cards') {
    rows = rows
      .map((r, i) => ({ r, i, k: sortKey(kind, r) }))
      .sort((a, b) => a.k - b.k || a.i - b.i)
      .map((x) => x.r)
  }
  if (opts.dedupe) {
    const seen = new Set<string>()
    const kept: Row[] = []
    for (const r of rows) {
      const key = JSON.stringify(r)
      if (seen.has(key)) { slot.duplicatesRemoved++; continue }
      seen.add(key)
      kept.push(r)
    }
    rows = kept
  }
  slot.rows = rows

  // overall range
  const froms = results.map(({ res }) => parseDmy(res.from)).filter((n): n is number => n !== null)
  const tos = results.map(({ res }) => parseDmy(res.to)).filter((n): n is number => n !== null)
  if (froms.length > 0) slot.from = fmtDmy(Math.min(...froms))
  if (tos.length > 0) slot.to = fmtDmy(Math.max(...tos))

  // gaps / overlaps (items)
  if (kind === 'items') {
    const ranges = results
      .map(({ res }) => ({ from: parseDmy(res.from), to: parseDmy(res.to) }))
      .filter((r): r is { from: number; to: number } => r.from !== null && r.to !== null)
      .sort((a, b) => a.from - b.from)
    let prev: { from: number; to: number } | undefined
    for (const cur of ranges) {
      if (prev) {
        if (cur.from > prev.to + DAY) {
          slot.warnings.push({ code: 'DATE_GAP', params: { from: fmtDmy(prev.to + DAY), to: fmtDmy(cur.from - DAY) } })
        } else if (cur.from <= prev.to) {
          slot.warnings.push({ code: 'DATE_OVERLAP', params: { from: fmtDmy(cur.from), to: fmtDmy(Math.min(prev.to, cur.to)) } })
        }
        if (cur.to > prev.to) prev = cur
      } else {
        prev = cur
      }
    }
  }

  // duplicate names (cards)
  if (kind === 'cards') {
    const names = duplicateNames(slot.rows)
    if (names.length > 0) {
      slot.warnings.push({ code: 'DUPLICATE_NAMES', params: { names: names.join(', '), count: names.length } })
    }
  }
  return slot
}

export function crossCheck(
  slots: Partial<Record<TableKey, ParsedSlot>>,
): { errors: Issue<ErrorCode>[]; warnings: Issue<WarningCode>[] } {
  const errors: Issue<ErrorCode>[] = []
  const warnings: Issue<WarningCode>[] = []
  const items = slots.items
  const topups = slots.topups

  if (items?.org && topups?.org && items.org !== topups.org) {
    errors.push({ code: 'ORG_MISMATCH', params: { orgs: `${items.org}, ${topups.org}` } })
  }

  const [y, m, d] = TEMPLATE_START.split('-').map(Number)
  const start = Date.UTC(y ?? 0, (m ?? 1) - 1, d ?? 1)
  for (const kind of ['items', 'topups'] as const) {
    const slot = slots[kind]
    if (!slot || !slot.from) continue
    const f = parseDmy(slot.from)
    if (f === null) continue
    if (f > start) warnings.push({ code: 'STARTS_AFTER_TEMPLATE_START', params: { kind, from: slot.from } })
    else if (f < start) warnings.push({ code: 'BEFORE_TEMPLATE_START', params: { kind, from: slot.from } })
  }
  return { errors, warnings }
}
