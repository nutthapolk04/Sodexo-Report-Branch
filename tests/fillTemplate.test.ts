import { describe, expect, test } from 'bun:test'
import JSZip from 'jszip'
import { fillTemplate, inspectTemplate } from '../src/lib/fillTemplate'
import { CARD_COLUMNS, ITEM_COLUMNS, TOPUP_COLUMNS } from '../src/types'
import type { Row } from '../src/types'

const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
const CHART = '<chart>untouched ไทย</chart>'
const DASH = `${HEAD}<worksheet ${NS}><sheetData><row r="1"><c r="A1"><f>COUNTA(WalletItems[ประเภท])</f><v>2</v></c></row></sheetData></worksheet>`

function strRow(r: number, n: number, ss0: number): string {
  let cells = ''
  for (let i = 0; i < n; i++) cells += `<c r="${String.fromCharCode(65 + i)}${r}" t="s"><v>${ss0 + i}</v></c>`
  return cells
}

function cardSheet(): string {
  let rows = `<row r="1" spans="1:19" ht="20">${strRow(1, 14, 0)}</row>`
  for (const r of [2, 3]) {
    rows += `<row r="${r}" spans="1:19" x14ac:dyDescent="0.25">${strRow(r, 14, 0).replace(/ t="s"><v>\d+<\/v>/g, ' s="4" t="s"><v>0</v>')}<c r="O${r}" s="7"/><c r="S${r}" s="7"/></row>`
  }
  rows += '<row r="4" spans="15:19"><c r="O4" s="7"/><c r="P4" s="7"/></row>'
  rows += '<row r="5" spans="1:19"><c r="A5" s="4" t="s"><v>0</v></c></row>'
  rows += '<row r="6" spans="15:19"><c r="R6" s="7"/></row>'
  return `${HEAD}<worksheet ${NS}><dimension ref="A1:S6"/><sheetData>${rows}</sheetData><pageMargins left="1"/></worksheet>`
}

function topupSheet(): string {
  const l = (r: number, t: number) => `<row r="${r}"><c r="A${r}" s="2" t="s"><v>${t}</v></c></row>`
  let rows = l(1, 0) + l(2, 1) + l(3, 2) + l(4, 3) + '<row r="5"/>' + `<row r="6">${strRow(6, 11, 0)}</row>`
  for (const r of [7, 8]) rows += `<row r="${r}">${strRow(r, 11, 0)}<c r="K${r}" s="9"><v>500</v></c></row>`
  return `${HEAD}<worksheet ${NS}><dimension ref="A1:K8"/><sheetData>${rows}</sheetData></worksheet>`
}

function calcCell(col: string, r: number): string {
  return `<c r="${col}${r}" t="str" cm="1"><f t="array" ref="${col}${r}">_xlfn.LET(x,1,x)</f><v>DCBK000147</v></c>`
}

function itemSheet(): string {
  const l = (r: number, t: number) => `<row r="${r}"><c r="A${r}" s="3" t="s"><v>${t}</v></c></row>`
  let rows = l(1, 0) + l(2, 1) + `<row r="5" spans="1:36">${strRow(5, 26, 0)}</row>`
  for (const r of [6, 7]) rows += `<row r="${r}" spans="1:36" ht="15">${strRow(r, 5, 0)}<c r="J${r}" s="5"><v>1</v></c>${calcCell('AG', r)}<c r="AH${r}" t="str" cm="1"><f t="array" ref="AH${r}">_xlfn.LET(y,2,y)</f><v/></c></row>`
  return `${HEAD}<worksheet ${NS}><dimension ref="A1:AJ7"/><sheetData>${rows}</sheetData></worksheet>`
}

const tableXml = (id: number, name: string, ref: string, extra = '') =>
  `${HEAD}<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="${id}" name="${name}" displayName="${name}" ref="${ref}" totalsRowShown="0"><autoFilter ref="${ref}"/>${extra}<tableColumns count="1"><tableColumn id="1" name="a"/></tableColumns></table>`

const relXml = (rels: string) => `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

async function buildTemplate(opts: { skipItemsTable?: boolean } = {}): Promise<ArrayBuffer> {
  const z = new JSZip()
  z.file('[Content_Types].xml', `${HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/><Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/></Types>`)
  z.file('_rels/.rels', relXml(`<Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/>`))
  z.file('xl/workbook.xml', `${HEAD}<workbook ${NS}><sheets><sheet name="Dashboard" sheetId="1" r:id="rId1"/><sheet name="Card" sheetId="2" r:id="rId2"/><sheet name="Card Top-up History" sheetId="3" r:id="rId3"/><sheet name="รายการสินค้า" sheetId="4" r:id="rId4"/></sheets><definedNames><definedName name="X">Card!$A$1</definedName></definedNames><calcPr calcId="191029"/></workbook>`)
  z.file('xl/_rels/workbook.xml.rels', relXml(`<Relationship Id="rId1" Type="${R}/worksheet" Target="worksheets/sheet9.xml"/><Relationship Id="rId2" Type="${R}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId3" Type="${R}/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId4" Type="${R}/worksheet" Target="/xl/worksheets/sheet3.xml"/><Relationship Id="rId5" Type="${R}/calcChain" Target="calcChain.xml"/><Relationship Id="rId6" Type="${R}/sharedStrings" Target="sharedStrings.xml"/>`))
  z.file('xl/calcChain.xml', `${HEAD}<calcChain ${NS}><c r="A1" i="1"/></calcChain>`)
  z.file('xl/sharedStrings.xml', `${HEAD}<sst ${NS} count="1" uniqueCount="1"><si><t>x</t></si></sst>`)
  z.file('xl/worksheets/sheet9.xml', DASH)
  z.file('xl/worksheets/sheet1.xml', cardSheet())
  z.file('xl/worksheets/sheet2.xml', topupSheet())
  z.file('xl/worksheets/sheet3.xml', itemSheet())
  z.file('xl/worksheets/_rels/sheet1.xml.rels', relXml(`<Relationship Id="rId1" Type="${R}/table" Target="../tables/table7.xml"/>`))
  z.file('xl/worksheets/_rels/sheet2.xml.rels', relXml(`<Relationship Id="rId1" Type="${R}/table" Target="../tables/table8.xml"/>`))
  z.file('xl/worksheets/_rels/sheet3.xml.rels', relXml(opts.skipItemsTable ? '' : `<Relationship Id="rId1" Type="${R}/table" Target="../tables/table9.xml"/>`))
  z.file('xl/tables/table7.xml', tableXml(1, 'WalletCards', 'A1:N3'))
  z.file('xl/tables/table8.xml', tableXml(2, 'WalletTopups', 'A6:K8'))
  if (!opts.skipItemsTable) z.file('xl/tables/table9.xml', tableXml(3, 'WalletItems', 'A5:AJ7', '<sortState ref="A6:AJ7"/>'))
  z.file('xl/charts/chart1.xml', CHART)
  z.file('xl/vbaProject.bin', new Uint8Array([0, 1, 2, 3, 255, 254]))
  return z.generateAsync({ type: 'arraybuffer' })
}

const cards: Row[] = [1, 2, 3, 4, 5].map((i) => Array.from({ length: CARD_COLUMNS.length }, (_, c) => (c === 0 ? `S${i}` : `v${i}-${c}`)))
const topups: Row[] = [1, 2, 3].map((i) => [`2026-08-1${i}`, `R${i}`, 'c', 'T', 'D', 'U', 'S', 'R', 'W', 'P', i * 100])
const items: Row[] = [1, 2, 3, 4].map((i) => Array.from({ length: ITEM_COLUMNS.length }, (_, c) => (c === 8 ? `A&B <"x"> ${i}` : c === 9 ? i : c === 3 ? '' : `s${i}`)))

async function run(over: Partial<{ cards: Row[]; topups: Row[]; items: Row[]; templateName: string }> = {}) {
  const template = await buildTemplate()
  const res = await fillTemplate({
    template, templateName: over.templateName ?? 'T.xlsm', cards: over.cards ?? cards, topups: over.topups ?? topups, items: over.items ?? items,
    org: 'DCI & Co', from: '17/08/2026', to: '25/10/2026',
  })
  const out = await JSZip.loadAsync(await res.blob.arrayBuffer())
  const text = async (p: string) => (await out.file(p)!.async('string'))
  return { res, out, text, template }
}

const rowsOf = (xml: string) => [...xml.matchAll(/<row r="(\d+)"/g)].map((m) => Number(m[1]))

describe('fillTemplate', () => {
  test('file name, mime and untouched files', async () => {
    const { res, out, template } = await run()
    expect(res.fileName).toBe('DCBK_POS_Report_20260817_20261025.xlsm')
    expect(res.blob.type.toLowerCase()).toBe('application/vnd.ms-excel.sheet.macroenabled.12')
    const orig = await JSZip.loadAsync(template)
    for (const p of ['xl/charts/chart1.xml', 'xl/vbaProject.bin', 'xl/worksheets/sheet9.xml', 'xl/sharedStrings.xml', '_rels/.rels']) {
      expect(await out.file(p)!.async('uint8array')).toEqual(await orig.file(p)!.async('uint8array'))
    }
  })

  test('xlsx extension gives xlsx mime', async () => {
    const { res } = await run({ templateName: 'a.xlsx' })
    expect(res.fileName.endsWith('.xlsx')).toBe(true)
    expect(res.blob.type).toBe('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
  })

  test('table refs, autoFilter, sortState and dimension', async () => {
    const { text } = await run()
    const t7 = await text('xl/tables/table7.xml')
    expect(t7).toContain('ref="A1:N6"')
    expect(t7).toContain('<autoFilter ref="A1:N6"')
    expect(await text('xl/tables/table8.xml')).toContain('<autoFilter ref="A6:K9"')
    const t9 = await text('xl/tables/table9.xml')
    expect(t9).toMatch(/<table [^>]*ref="A5:AJ9"/)
    expect(t9).toContain('<autoFilter ref="A5:AJ9"')
    expect(t9).toContain('<sortState ref="A6:AJ9"')
    expect(await text('xl/worksheets/sheet1.xml')).toContain('<dimension ref="A1:S6"/>')
    expect(await text('xl/worksheets/sheet2.xml')).toContain('<dimension ref="A1:K9"/>')
    expect(await text('xl/worksheets/sheet3.xml')).toContain('<dimension ref="A1:AJ9"/>')
  })

  test('card rows: values, styles, preserved O..S cells, spans', async () => {
    const x = await (await run()).text('xl/worksheets/sheet1.xml')
    expect(rowsOf(x)).toEqual([1, 2, 3, 4, 5, 6])
    expect(x).toContain('<c r="A2" s="4" t="inlineStr"><is><t xml:space="preserve">S1</t></is></c>')
    expect(x).toContain('<c r="O2" s="7"/><c r="S2" s="7"/>')
    expect(x).toContain('<c r="O4" s="7"/><c r="P4" s="7"/>')
    expect(x).toContain('<c r="R6" s="7"/>')
    // spans are recomputed from the cells actually written
    expect(x).toContain('<row r="4" spans="1:16">')
    expect(x).toContain('<row r="5" spans="1:14">')
    expect(x).toContain('<row r="2" spans="1:19" x14ac:dyDescent="0.25">')
    expect(x).toContain('x14ac:dyDescent="0.25"')
    expect(x).toContain('<pageMargins left="1"/>')
    expect(x).not.toContain('t="s"><v>0</v></c></row><row r="6"')
    expect(x.match(/<c r="A5"/g)!.length).toBe(1)
  })

  test('top-up: header lines, numbers, row count', async () => {
    const x = await (await run()).text('xl/worksheets/sheet2.xml')
    expect(x).toContain('<c r="A2" s="2" t="inlineStr"><is><t xml:space="preserve">Org/Branch: DCI &amp; Co</t></is></c>')
    expect(x).toContain('Date Range: 17/08/2026 - 25/10/2026')
    expect(x).toMatch(/Exported At: \d{2}\/\d{2}\/\d{4}, \d{2}:\d{2}/)
    expect(x).toContain('<c r="A1" s="2" t="s">')
    expect(x).toContain('<c r="K9" s="9"><v>300</v></c>')
    expect(x).toContain('<c r="K7" s="9"><v>100</v></c>')
    expect(rowsOf(x)).toEqual([1, 2, 3, 4, 5, 6, 7, 8, 9])
  })

  test('items: escaping, empty string, calculated columns on every new row', async () => {
    const x = await (await run()).text('xl/worksheets/sheet3.xml')
    expect(x).toContain('องค์กร/สาขา: DCI &amp; Co')
    expect(x).toContain('A&amp;B &lt;&quot;x&quot;&gt; 1')
    expect(x).toContain('<c r="D6" t="inlineStr"><is><t xml:space="preserve"></t></is></c>')
    expect(x).toContain('<c r="J6" s="5"><v>1</v></c>')
    expect(x).toContain('<c r="J9" s="5"><v>4</v></c>')
    expect(rowsOf(x)).toEqual([1, 2, 5, 6, 7, 8, 9])
    for (const r of [6, 7, 8, 9]) {
      expect(x).toContain(`<c r="AG${r}" t="str" cm="1"><f t="array" ref="AG${r}">_xlfn.LET(x,1,x)</f><v></v></c>`)
      expect(x).toContain(`<c r="AH${r}" t="str" cm="1"><f t="array" ref="AH${r}">_xlfn.LET(y,2,y)</f><v></v></c>`)
    }
    expect(x).not.toContain('DCBK000147')
    expect(x).toContain('<row r="9" spans="1:34" ht="15">')
  })

  test('illegal XML chars are stripped', async () => {
    const bad: Row[] = [['a\u0001b\u0000c', ...Array(10).fill('')] as Row]
    const x = await (await run({ topups: bad })).text('xl/worksheets/sheet2.xml')
    expect(x).toContain('>abc<')
  })

  test('workbook calc flags and calcChain removal', async () => {
    const { out, text } = await run()
    expect(await text('xl/workbook.xml')).toContain('<calcPr calcId="191029" fullCalcOnLoad="1"/>')
    expect(out.file('xl/calcChain.xml')).toBeNull()
    expect(await text('xl/_rels/workbook.xml.rels')).not.toContain('calcChain')
    expect(await text('xl/_rels/workbook.xml.rels')).toContain('sharedStrings')
    expect(await text('[Content_Types].xml')).not.toContain('calcChain')
    expect(await text('[Content_Types].xml')).toContain('workbook.xml')
  })

  test('fewer rows than template removes old rows', async () => {
    const { text } = await run({ items: items.slice(0, 1), topups: [], cards: cards.slice(0, 1) })
    const t = await text('xl/worksheets/sheet3.xml')
    expect(rowsOf(t)).toEqual([1, 2, 5, 6])
    expect(await text('xl/tables/table9.xml')).toContain('ref="A5:AJ6"')
    const c = await text('xl/worksheets/sheet1.xml')
    // rows 3 and 4 had only styled extras: row 3 keeps O/S, row 4 keeps O/P, row 2 is replaced by data
    expect(rowsOf(c)).toEqual([1, 2, 3, 4, 5, 6])
    expect(c).toContain('<row r="3" spans="15:19" x14ac:dyDescent="0.25"><c r="O3" s="7"/><c r="S3" s="7"/></row>')
    expect(c).not.toContain('r="A3"')
    expect(c).toContain('<c r="A5" s="4" t="s">')
    const u = await text('xl/worksheets/sheet2.xml')
    expect(rowsOf(u)).toEqual([1, 2, 3, 4, 5, 6, 7])
    expect(await text('xl/tables/table8.xml')).toContain('ref="A6:K7"')
  })

  test('onStep reports 0..4', async () => {
    const steps: number[] = []
    await fillTemplate({ template: await buildTemplate(), templateName: 'a.xlsm', cards, topups, items, org: 'o', from: '01/01/2026', to: '02/01/2026', onStep: (s) => { if (!steps.includes(s)) steps.push(s) } })
    expect(steps).toEqual([0, 1, 2, 3, 4])
  })

  test('fill without table throws TemplateError', async () => {
    const template = await buildTemplate({ skipItemsTable: true })
    await expect(fillTemplate({ template, templateName: 'a.xlsm', cards, topups, items, org: 'o', from: '01/01/2026', to: '02/01/2026' })).rejects.toMatchObject({ issue: { code: 'TEMPLATE_TABLE_MISSING' } })
  })
})

describe('inspectTemplate', () => {
  test('ok', async () => {
    const buf = await buildTemplate()
    const info = await inspectTemplate('T.xlsm', buf)
    expect(info.errors).toEqual([])
    expect(info.sheetCount).toBe(4)
    expect(info.sizeBytes).toBe(buf.byteLength)
    expect(info.tables.sort()).toEqual(['WalletCards', 'WalletItems', 'WalletTopups'])
  })
  test('missing table', async () => {
    const info = await inspectTemplate('T.xlsm', await buildTemplate({ skipItemsTable: true }))
    expect(info.errors).toEqual([{ code: 'TEMPLATE_TABLE_MISSING', params: { tables: 'WalletItems' } }])
  })
  test('not xlsx', async () => {
    const info = await inspectTemplate('x.txt', new TextEncoder().encode('hello').buffer as ArrayBuffer)
    expect(info.errors).toEqual([{ code: 'NOT_XLSX', params: { file: 'x.txt' } }])
    const z = new JSZip(); z.file('a.txt', 'a')
    expect((await inspectTemplate('y.zip', await z.generateAsync({ type: 'arraybuffer' }))).errors[0]!.code).toBe('NOT_XLSX')
  })
})
