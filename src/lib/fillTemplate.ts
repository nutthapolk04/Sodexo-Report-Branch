// Fills the customer's Excel template with rows from UPOS exports.
// HARD RULE: no SheetJS/ExcelJS round-trip. Only jszip + XML string edits, so charts,
// data validation, dynamic-array formulas and macros stay byte-identical.
import JSZip from 'jszip'
import { CARD_COLUMNS, ITEM_COLUMNS, TABLE_NAMES, TOPUP_COLUMNS } from '../types'
import type { Cell, ErrorCode, FillInput, FillResult, Issue, Row, TableKey, TemplateInfo } from '../types'

/** Thrown by fillTemplate for unusable templates; carries the i18n-ready issue. */
export class TemplateError extends Error {
  issue: Issue<ErrorCode>
  constructor(issue: Issue<ErrorCode>) {
    super(issue.code)
    this.name = 'TemplateError'
    this.issue = issue
  }
}

// ---------------------------------------------------------------- small helpers

const TABLE_KEYS: TableKey[] = ['cards', 'topups', 'items']
const COLUMN_COUNT: Record<TableKey, number> = {
  cards: CARD_COLUMNS.length,
  topups: TOPUP_COLUMNS.length,
  items: ITEM_COLUMNS.length,
}

const colCache: string[] = []
function idxToCol(i: number): string {
  const hit = colCache[i]
  if (hit !== undefined) return hit
  let n = i + 1
  let s = ''
  while (n > 0) {
    const m = (n - 1) % 26
    s = String.fromCharCode(65 + m) + s
    n = Math.floor((n - 1) / 26)
  }
  colCache[i] = s
  return s
}

function colToIdx(letters: string): number {
  let n = 0
  for (let i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) - 64)
  return n - 1
}

interface ParsedRef {
  c0: number
  r0: number
  c1: number
  r1: number
}

function parseRef(ref: string): ParsedRef | null {
  const m = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(ref)
  if (!m) return null
  return { c0: colToIdx(m[1]!), r0: Number(m[2]), c1: colToIdx(m[3]!), r1: Number(m[4]) }
}

function attr(tag: string, name: string): string | null {
  const m = new RegExp(`(?:^|\\s)${name}="([^"]*)"`).exec(tag)
  return m ? m[1]! : null
}

function unescapeXml(s: string): string {
  return s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&')
}

// XML 1.0 illegal chars + lone surrogates
const ILLEGAL_XML = /[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/g

function escapeText(s: string): string {
  return s
    .replace(ILLEGAL_XML, '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

function cellXml(ref: string, v: Cell | undefined, style: string | undefined): string {
  const s = style === undefined ? '' : ` s="${style}"`
  if (typeof v === 'number' && Number.isFinite(v)) return `<c r="${ref}"${s}><v>${v}</v></c>`
  const text = typeof v === 'string' ? v : ''
  return `<c r="${ref}"${s} t="inlineStr"><is><t xml:space="preserve">${escapeText(text)}</t></is></c>`
}

/** Resolve a relationship Target against the directory of the part owning the rels file (dir ends with '/' or is ''). */
function resolveTarget(baseDir: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1)
  const out: string[] = []
  for (const seg of (baseDir + target).split('/')) {
    if (seg === '' || seg === '.') continue
    if (seg === '..') out.pop()
    else out.push(seg)
  }
  return out.join('/')
}

function dirOf(path: string): string {
  const i = path.lastIndexOf('/')
  return i < 0 ? '' : path.slice(0, i + 1)
}

function relsPathFor(partPath: string): string {
  const i = partPath.lastIndexOf('/')
  return `${partPath.slice(0, i + 1)}_rels/${partPath.slice(i + 1)}.rels`
}

interface Rel {
  id: string
  type: string
  target: string
}

function parseRels(xml: string): Rel[] {
  const out: Rel[] = []
  const re = /<Relationship\b[^>]*>/g
  let m: RegExpExecArray | null
  while ((m = re.exec(xml)) !== null) {
    const tag = m[0]
    out.push({ id: attr(tag, 'Id') ?? '', type: attr(tag, 'Type') ?? '', target: attr(tag, 'Target') ?? '' })
  }
  return out
}

async function readText(zip: JSZip, path: string): Promise<string | null> {
  const f = zip.file(path)
  return f ? f.async('string') : null
}

// ---------------------------------------------------------------- template structure

interface TableInfo {
  displayName: string
  tablePath: string
  ref: string
}

async function readTables(zip: JSZip): Promise<TableInfo[]> {
  const out: TableInfo[] = []
  const paths = Object.keys(zip.files).filter((p) => /^xl\/tables\/[^/]+\.xml$/.test(p))
  for (const tablePath of paths) {
    const xml = await readText(zip, tablePath)
    if (xml === null) continue
    const root = /<table\b[^>]*>/.exec(xml)
    if (!root) continue
    const displayName = attr(root[0], 'displayName')
    const ref = attr(root[0], 'ref')
    if (displayName && ref) out.push({ displayName: unescapeXml(displayName), tablePath, ref })
  }
  return out
}

/** Worksheet part paths listed in workbook.xml (resolved through workbook.xml.rels). */
async function readSheetPaths(zip: JSZip, workbookXml: string): Promise<string[]> {
  const relsXml = (await readText(zip, 'xl/_rels/workbook.xml.rels')) ?? ''
  const rels = parseRels(relsXml)
  const out: string[] = []
  const re = /<sheet\s[^>]*>/g
  let m: RegExpExecArray | null
  while ((m = re.exec(workbookXml)) !== null) {
    const idm = /\s\w+:id="([^"]*)"/.exec(m[0])
    const rel = idm ? rels.find((r) => r.id === idm[1]) : undefined
    if (rel) out.push(resolveTarget('xl/', rel.target))
  }
  return out
}

/** Map table part path -> worksheet path that references it. */
async function mapTablesToSheets(zip: JSZip, sheetPaths: string[]): Promise<Map<string, string>> {
  const map = new Map<string, string>()
  for (const sheetPath of sheetPaths) {
    const relsXml = await readText(zip, relsPathFor(sheetPath))
    if (relsXml === null) continue
    for (const rel of parseRels(relsXml)) {
      if (!/\/table$/.test(rel.type) && !/tables\//.test(rel.target)) continue
      map.set(resolveTarget(dirOf(sheetPath), rel.target), sheetPath)
    }
  }
  return map
}

async function openTemplate(fileName: string, data: ArrayBuffer): Promise<{ zip: JSZip; workbookXml: string }> {
  let zip: JSZip
  try {
    zip = await JSZip.loadAsync(data)
  } catch {
    throw new TemplateError({ code: 'NOT_XLSX', params: { file: fileName } })
  }
  const workbookXml = await readText(zip, 'xl/workbook.xml')
  if (workbookXml === null) throw new TemplateError({ code: 'NOT_XLSX', params: { file: fileName } })
  return { zip, workbookXml }
}

export async function inspectTemplate(fileName: string, data: ArrayBuffer): Promise<TemplateInfo> {
  const info: TemplateInfo = { fileName, sizeBytes: data.byteLength, sheetCount: 0, tables: [], errors: [] }
  let opened: { zip: JSZip; workbookXml: string }
  try {
    opened = await openTemplate(fileName, data)
  } catch (e) {
    if (e instanceof TemplateError) {
      info.errors.push(e.issue)
      return info
    }
    throw e
  }
  info.sheetCount = (opened.workbookXml.match(/<sheet\s/g) ?? []).length
  info.tables = (await readTables(opened.zip)).map((t) => t.displayName)
  for (const key of TABLE_KEYS) {
    const name = TABLE_NAMES[key]
    if (!info.tables.includes(name)) info.errors.push({ code: 'TEMPLATE_TABLE_MISSING', params: { tables: name } })
  }
  return info
}

// ---------------------------------------------------------------- sheet rewriting

interface PlainCell {
  col: number
  start: number
  end: number
  head: string
}

/** Cells of one <row> content, in document order. */
function parseCells(content: string): PlainCell[] {
  const out: PlainCell[] = []
  let pos = 0
  for (;;) {
    const i = content.indexOf('<c', pos)
    if (i < 0) break
    const next = content.charCodeAt(i + 2)
    if (next !== 32 && next !== 62 && next !== 47) {
      pos = i + 2
      continue
    }
    const gt = content.indexOf('>', i)
    if (gt < 0) break
    const head = content.slice(i, gt + 1)
    const end = content.charCodeAt(gt - 1) === 47 ? gt + 1 : content.indexOf('</c>', gt) + 4
    const rm = /\sr="([A-Z]+)\d+"/.exec(head)
    out.push({ col: rm ? colToIdx(rm[1]!) : -1, start: i, end, head })
    pos = end
  }
  return out
}

interface RowAt {
  r: number
  openTag: string
  selfClosed: boolean
  contentStart: number
  contentEnd: number
  end: number
}

function readRowAt(body: string, idx: number): RowAt {
  const gt = body.indexOf('>', idx)
  const selfClosed = body.charCodeAt(gt - 1) === 47
  const openTag = body.slice(idx, gt + 1)
  const closeAt = selfClosed ? gt + 1 : body.indexOf('</row>', gt)
  return {
    r: Number(attr(openTag, 'r') ?? 0),
    openTag,
    selfClosed,
    contentStart: gt + 1,
    contentEnd: selfClosed ? gt + 1 : closeAt,
    end: selfClosed ? gt + 1 : closeAt + 6,
  }
}

function fixRowOpen(tag: string, r: number, lo: number, hi: number): string {
  let t = tag.endsWith('/>') ? `${tag.slice(0, -2)}>` : tag
  t = t.replace(/(\s)r="\d+"/, `$1r="${r}"`)
  if (/\sspans="/.test(t) && hi >= 0) t = t.replace(/(\s)spans="[^"]*"/, `$1spans="${lo + 1}:${hi + 1}"`)
  return t
}

interface SheetPlan {
  c0: number
  c1: number
  headerRow: number
  oldLast: number
  newLast: number
  n: number
  rows: Row[]
  /** row number -> replacement text for the column-A cell of that (pre-header) row */
  headerText: Map<number, string>
  onProgress?: (done: number, total: number) => void
}

interface Kept {
  col: number
  xml: string
}

const tick = (): Promise<void> => new Promise((resolve) => setTimeout(resolve, 0))

async function rewriteSheet(xml: string, p: SheetPlan): Promise<string> {
  const sdOpen = xml.indexOf('<sheetData')
  if (sdOpen < 0) throw new Error('sheetData missing')
  const sdGt = xml.indexOf('>', sdOpen)
  let prefix: string
  let suffix: string
  let body: string
  if (xml.charCodeAt(sdGt - 1) === 47) {
    prefix = `${xml.slice(0, sdOpen)}<sheetData>`
    suffix = `</sheetData>${xml.slice(sdGt + 1)}`
    body = ''
  } else {
    const bodyEnd = xml.indexOf('</sheetData>', sdGt)
    prefix = xml.slice(0, sdGt + 1)
    suffix = xml.slice(bodyEnd)
    body = xml.slice(sdGt + 1, bodyEnd)
  }

  // Template (first data row): row attributes, per-column style, calculated-column cells.
  let tplOpen = '<row>'
  const styles = new Map<number, string>()
  const calc: Kept[] = []
  const ti = body.indexOf(`<row r="${p.headerRow + 1}"`)
  if (ti >= 0) {
    const rr = readRowAt(body, ti)
    tplOpen = rr.openTag
    const content = body.slice(rr.contentStart, rr.contentEnd)
    for (const c of parseCells(content)) {
      if (c.col < p.c0 || c.col > p.c1) continue
      if (c.col < p.c0 + p.n) {
        const s = attr(c.head, 's')
        if (s !== null) styles.set(c.col, s)
      } else {
        calc.push({ col: c.col, xml: content.slice(c.start, c.end) })
      }
    }
  }

  const parts: string[] = []
  let maxRow = 0
  let next = p.headerRow + 1
  const rangeEnd = Math.max(p.oldLast, p.newLast)
  const total = p.newLast - p.headerRow
  let written = 0

  const buildRow = (r: number, srcOpen: string, kept: Kept[], withData: boolean): string | null => {
    if (!withData && kept.length === 0) return null
    let lo = Infinity
    let hi = -1
    for (const k of kept) {
      if (k.col < lo) lo = k.col
      if (k.col > hi) hi = k.col
    }
    const dataEnd = p.c0 + p.n - 1
    if (withData) {
      lo = Math.min(lo, p.c0)
      hi = Math.max(hi, dataEnd)
      for (const c of calc) hi = Math.max(hi, c.col)
    }
    const cells: string[] = []
    for (const k of kept) if (k.col < p.c0) cells.push(k.xml)
    if (withData) {
      const row = p.rows[r - p.headerRow - 1]
      for (let i = 0; i < p.n; i++) {
        const col = p.c0 + i
        cells.push(cellXml(`${idxToCol(col)}${r}`, row?.[i], styles.get(col)))
      }
      for (const c of calc) {
        cells.push(
          c.xml
            .replace(/(\s)(r|ref)="([A-Z]+)\d+"/g, `$1$2="$3${r}"`)
            .replace(/<v>[\s\S]*?<\/v>|<v\s*\/>/, '<v></v>'),
        )
      }
    }
    for (const k of kept) if (k.col > p.c1) cells.push(k.xml)
    maxRow = Math.max(maxRow, r)
    return `${fixRowOpen(srcOpen, r, lo, hi)}${cells.join('')}</row>`
  }

  const noteDone = async (): Promise<void> => {
    written++
    if (written % 4000 === 0) {
      p.onProgress?.(written, total)
      await tick()
    }
  }

  const flushTo = async (upto: number): Promise<void> => {
    while (next <= upto) {
      const s = buildRow(next, tplOpen, [], true)
      if (s !== null) parts.push(s)
      next++
      await noteDone()
    }
  }

  let pos = 0
  for (;;) {
    const idx = body.indexOf('<row ', pos)
    if (idx < 0) break
    const rr = readRowAt(body, idx)
    pos = rr.end
    const r = rr.r
    if (r <= p.headerRow) {
      maxRow = Math.max(maxRow, r)
      const text = p.headerText.get(r)
      if (text === undefined) {
        parts.push(body.slice(idx, rr.end))
        continue
      }
      const content = body.slice(rr.contentStart, rr.contentEnd)
      const cells = parseCells(content)
      const a = cells.find((c) => c.col === 0)
      const style = a ? attr(a.head, 's') : null
      const repl = cellXml(`A${r}`, text, style ?? undefined)
      let newContent: string
      if (a) newContent = content.slice(0, a.start) + repl + content.slice(a.end)
      else {
        const first = cells[0]
        const at = first ? first.start : 0
        newContent = content.slice(0, at) + repl + content.slice(at)
      }
      const open = rr.openTag.endsWith('/>') ? `${rr.openTag.slice(0, -2)}>` : rr.openTag
      parts.push(`${open}${newContent}</row>`)
      continue
    }
    if (r <= rangeEnd) {
      await flushTo(Math.min(r - 1, p.newLast))
      const content = body.slice(rr.contentStart, rr.contentEnd)
      const kept: Kept[] = []
      for (const c of parseCells(content)) {
        if (c.col < p.c0 || c.col > p.c1) kept.push({ col: c.col, xml: content.slice(c.start, c.end) })
      }
      const withData = r <= p.newLast
      const s = buildRow(r, rr.openTag, kept, withData)
      if (s !== null) parts.push(s)
      if (r >= next) next = r + 1
      if (withData) await noteDone()
      continue
    }
    await flushTo(p.newLast)
    maxRow = Math.max(maxRow, r)
    parts.push(body.slice(idx, rr.end))
  }
  await flushTo(p.newLast)
  p.onProgress?.(total, total)

  // <dimension>: keep original start and last column, new last row.
  const dim = /<dimension\s+ref="([^"]*)"\s*\/>/.exec(prefix)
  if (dim) {
    const parts2 = dim[1]!.split(':')
    const startRef = parts2[0]!
    const endRef = parts2[1] ?? parts2[0]!
    const em = /^([A-Z]+)\d+$/.exec(endRef)
    const endCol = em ? Math.max(colToIdx(em[1]!), p.c1) : p.c1
    const newRef = `${startRef}:${idxToCol(endCol)}${Math.max(maxRow, 1)}`
    prefix = prefix.replace(dim[0], `<dimension ref="${newRef}"/>`)
  }
  return prefix + parts.join('') + suffix
}

function rewriteRowsOfRef(ref: string, firstRow: number, lastRow: number): string {
  const m = /^([A-Z]+)\d+(?::([A-Z]+)\d+)?$/.exec(ref)
  if (!m) return ref
  return `${m[1]}${firstRow}:${m[2] ?? m[1]}${lastRow}`
}

function updateTableXml(xml: string, c0: number, c1: number, headerRow: number, newLast: number): string {
  const full = `${idxToCol(c0)}${headerRow}:${idxToCol(c1)}${newLast}`
  let out = xml.replace(/<table\b[^>]*>/, (tag) => tag.replace(/(\s)ref="[^"]*"/, `$1ref="${full}"`))
  out = out.replace(/(<autoFilter\b[^>]*?\sref=")[^"]*(")/, `$1${full}$2`)
  out = out.replace(/(<sortState\b[^>]*?\sref=")([^"]*)(")/g, (_m, a: string, ref: string, b: string) => `${a}${rewriteRowsOfRef(ref, headerRow + 1, newLast)}${b}`)
  out = out.replace(/(<sortCondition\b[^>]*?\sref=")([^"]*)(")/g, (_m, a: string, ref: string, b: string) => `${a}${rewriteRowsOfRef(ref, headerRow + 1, newLast)}${b}`)
  return out
}

function setFullCalcOnLoad(xml: string): string {
  const m = /<calcPr\b([^>]*?)(\/?)>/.exec(xml)
  if (m) {
    let attrs = m[1]!
    if (/\sfullCalcOnLoad="[^"]*"/.test(attrs)) attrs = attrs.replace(/(\s)fullCalcOnLoad="[^"]*"/, '$1fullCalcOnLoad="1"')
    else attrs = `${attrs} fullCalcOnLoad="1"`
    return xml.replace(m[0], `<calcPr${attrs}${m[2]}>`)
  }
  // schema order: ... definedNames, calcPr, oleSize, customWorkbookViews, pivotCaches, ..., extLst
  const after = /<(oleSize|customWorkbookViews|pivotCaches|smartTagPr|smartTagTypes|webPublishing|fileRecoveryPr|webPublishObjects|extLst)\b/.exec(xml)
  const tag = '<calcPr fullCalcOnLoad="1"/>'
  if (after) return xml.slice(0, after.index) + tag + xml.slice(after.index)
  const end = xml.lastIndexOf('</workbook>')
  return end < 0 ? xml + tag : xml.slice(0, end) + tag + xml.slice(end)
}

// ---------------------------------------------------------------- fillTemplate

// Replacing an existing part must not add folder entries to the package.
const NO_DIRS = { createFolders: false }
const pad2 = (n: number): string => String(n).padStart(2, '0')
const fmtInt = (n: number): string => n.toLocaleString('en-US')

function ymd(ddmmyyyy: string): string {
  const m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(ddmmyyyy.trim())
  if (!m) return ddmmyyyy.replace(/\D/g, '')
  return `${m[3]}${pad2(Number(m[2]))}${pad2(Number(m[1]))}`
}

export async function fillTemplate(input: FillInput): Promise<FillResult> {
  const step = input.onStep
  step?.(0)
  const { zip, workbookXml } = await openTemplate(input.templateName, input.template)

  step?.(1)
  const tables = await readTables(zip)
  for (const key of TABLE_KEYS) {
    if (!tables.some((t) => t.displayName === TABLE_NAMES[key])) {
      throw new TemplateError({ code: 'TEMPLATE_TABLE_MISSING', params: { tables: TABLE_NAMES[key] } })
    }
  }
  const sheetPaths = await readSheetPaths(zip, workbookXml)
  const tableToSheet = await mapTablesToSheets(zip, sheetPaths)

  const now = new Date()
  const exportedAt = `${pad2(now.getDate())}/${pad2(now.getMonth() + 1)}/${now.getFullYear()}, ${pad2(now.getHours())}:${pad2(now.getMinutes())}`
  const dataByKey: Record<TableKey, Row[]> = { cards: input.cards, topups: input.topups, items: input.items }
  const headerTextByKey: Record<TableKey, Map<number, string>> = {
    cards: new Map(),
    topups: new Map([
      [2, `Org/Branch: ${input.org}`],
      [3, `Date Range: ${input.from} - ${input.to}`],
      [4, `Exported At: ${exportedAt}`],
    ]),
    items: new Map([[2, `องค์กร/สาขา: ${input.org}`]]),
  }

  step?.(2)
  for (const key of TABLE_KEYS) {
    const table = tables.find((t) => t.displayName === TABLE_NAMES[key])!
    const sheetPath = tableToSheet.get(table.tablePath)
    const ref = parseRef(table.ref)
    if (!sheetPath || !ref) throw new TemplateError({ code: 'TEMPLATE_TABLE_MISSING', params: { tables: TABLE_NAMES[key] } })
    const sheetXml = await readText(zip, sheetPath)
    if (sheetXml === null) throw new TemplateError({ code: 'TEMPLATE_TABLE_MISSING', params: { tables: TABLE_NAMES[key] } })

    const rows = dataByKey[key]
    const width = ref.c1 - ref.c0 + 1
    const n = Math.min(COLUMN_COUNT[key], width)
    // A table needs at least one data row; with no input rows we keep one empty row.
    const newLast = ref.r0 + Math.max(1, rows.length)
    const label = TABLE_NAMES[key]
    const newSheet = await rewriteSheet(sheetXml, {
      c0: ref.c0,
      c1: ref.c1,
      headerRow: ref.r0,
      oldLast: ref.r1,
      newLast,
      n,
      rows,
      headerText: headerTextByKey[key],
      onProgress: (done, total) => {
        if (total > 4000) step?.(2, `${label} ${fmtInt(Math.min(done, rows.length))} / ${fmtInt(rows.length)}`)
      },
    })
    zip.file(sheetPath, newSheet, NO_DIRS)
    const tableXml = await readText(zip, table.tablePath)
    if (tableXml !== null) zip.file(table.tablePath, updateTableXml(tableXml, ref.c0, ref.c1, ref.r0, newLast), NO_DIRS)
  }

  step?.(3)
  zip.file('xl/workbook.xml', setFullCalcOnLoad(workbookXml), NO_DIRS)
  const relsPath = 'xl/_rels/workbook.xml.rels'
  const relsXml = await readText(zip, relsPath)
  let calcChainPath = 'xl/calcChain.xml'
  if (relsXml !== null) {
    const cc = parseRels(relsXml).find((r) => /\/calcChain$/.test(r.type) || /calcChain\.xml$/.test(r.target))
    if (cc) calcChainPath = resolveTarget('xl/', cc.target)
    zip.file(relsPath, relsXml.replace(/<Relationship\b[^>]*calcChain[^>]*\/>/g, ''), NO_DIRS)
  }
  zip.remove(calcChainPath)
  const ctXml = await readText(zip, '[Content_Types].xml')
  if (ctXml !== null) zip.file('[Content_Types].xml', ctXml.replace(/<Override\b[^>]*calcChain[^>]*\/>/g, ''), NO_DIRS)

  step?.(4)
  const ext = (/\.([A-Za-z0-9]+)$/.exec(input.templateName)?.[1] ?? 'xlsx').toLowerCase()
  const mimeType =
    ext === 'xlsm' ? 'application/vnd.ms-excel.sheet.macroEnabled.12' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  const blob = await zip.generateAsync({ type: 'blob', compression: 'DEFLATE', mimeType })
  return { blob, fileName: `DCBK_POS_Report_${ymd(input.from)}_${ymd(input.to)}.${ext}` }
}
