import { describe, expect, test } from 'bun:test'
import * as XLSX from 'xlsx'
import { crossCheck, readSlot } from '../src/lib/readExports'
import type { InputFile } from '../src/lib/readExports'
import { CARD_COLUMNS, ITEM_COLUMNS, TOPUP_COLUMNS } from '../src/types'
import type { ReadOptions } from '../src/types'

type V = string | number
const OPTS: ReadOptions = { replaceEmDash: true, dedupe: true }

function xlsx(name: string, aoa: V[][]): InputFile {
  const ws = XLSX.utils.aoa_to_sheet(aoa)
  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1')
  const out = XLSX.write(wb, { type: 'array', bookType: 'xlsx' }) as ArrayBuffer
  return { name, data: out }
}

const ITEM_HEADER_TH = [
  'ประเภท', 'วันที่', 'เวลา', 'เคาน์เตอร์', 'เลขที่ใบเสร็จ', 'ประเภท', 'ช่วงเวลา', 'รหัสสินค้า', 'ชื่อรายการ', 'จำนวน', 'ราคา',
  'ตัวเลือกเสริม', 'ราคาตัวเลือกเสริม', 'ยอดรวมสินค้า', 'ส่วนลด', 'ค่าบริการ', 'ภาษีมูลค่าเพิ่ม', 'ยอดสุทธิ', 'เลขอ้างอิง', 'เลขที่ธุรกรรม',
  'ช่องทางชำระเงิน', 'ชื่อการชำระเงิน', 'หมายเหตุการชำระเงิน', 'ชำระ', 'คงเหลือ', 'ยอดที่รับ', 'ชื่อลูกค้า', 'ประเภทผู้ใช้',
  'หมายเหตุออเดอร์', 'หมายเหตุรายการ', 'อ้างอิงถึง', 'สถานะ',
]

function itemRow(date: string, time: string, receipt: string, name = 'ข้าวมันไก่'): V[] {
  const r: V[] = ITEM_HEADER_TH.map((_, i) => `c${i}`)
  r[0] = 'ขาย'; r[1] = date; r[2] = time; r[4] = receipt; r[5] = 'อาหาร'; r[8] = name; r[9] = 1
  return r
}

function itemsFile(name: string, org: string, rows: V[][], header: V[] = ITEM_HEADER_TH): InputFile {
  return xlsx(name, [
    ['Sales Transactions'],
    [`สาขา/องค์กร: ${org}`],
    ['ช่วงวันที่: x'],
    ['เวลาที่ export: y'],
    [],
    header,
    ...rows,
  ])
}

function topupFile(name: string, org: string, rows: V[][], header: V[] = [...TOPUP_COLUMNS]): InputFile {
  return xlsx(name, [['Top-up History'], [`Org/Branch: ${org}`], ['Date Range: x'], ['Exported At: y'], [], header, ...rows])
}

const tRow = (dt: string, rec = 'R1'): V[] => [dt, rec, 'C1', 'Top-up', 'POS', 'Mom', 'S1', 'Kid', 'Meal', 'Cash', 100]

const CARD_H = [...CARD_COLUMNS]
const cRow = (id: string, th1: string, th2: string, en1: string, en2: string): V[] =>
  [id, th1, th2, en1, en2, '', '', '', '', '', 'นักเรียน', 'ใช้งาน', '123', '01/01/2030']

describe('items', () => {
  test('finds header by name under title rows; maps both ประเภท; reads org', () => {
    const s = readSlot('items', [itemsFile('a.xlsx', 'DCI', [itemRow('19/08/2026', '12:00', 'A1')])], OPTS)
    expect(s.errors).toEqual([])
    expect(s.org).toBe('DCI')
    expect(s.rows.length).toBe(1)
    const r = s.rows[0]!
    expect(r.length).toBe(ITEM_COLUMNS.length)
    expect(r[0]).toBe('ขาย')
    expect(r[5]).toBe('อาหาร')
    expect(r[9]).toBe(1) // number stays number
    expect(r[1]).toBe('19/08/2026') // text stays text
    expect(s.from).toBe('19/08/2026')
    expect(s.to).toBe('19/08/2026')
  })

  test('Bills tab export -> BILLS_NOT_TRANSACTIONS', () => {
    const f = xlsx('bills.xlsx', [['Bills'], [], ['วันที่', 'เวลา', 'เลขที่ใบเสร็จ', 'ประเภท', 'ช่วงเวลา', 'รหัสสินค้า', 'ชื่อรายการ', 'จำนวน'], ['19/08/2026', '12:00', 'A1', 'x', 'y', 'z', 'n', 1]])
    const s = readSlot('items', [f], OPTS)
    expect(s.errors[0]?.code).toBe('BILLS_NOT_TRANSACTIONS')
    expect(s.rows.length).toBe(0)
  })

  test('English sales export -> SALES_NOT_THAI', () => {
    const f = xlsx('en.xlsx', [['t'], ['Type', 'Date', 'Time', 'Receipt No.', 'Category', 'Period', 'Code', 'Name', 'Qty']])
    expect(readSlot('items', [f], OPTS).errors[0]?.code).toBe('SALES_NOT_THAI')
  })

  test('not an xlsx -> NOT_XLSX', () => {
    const s = readSlot('items', [{ name: 'bad.xlsx', data: new TextEncoder().encode('hello').buffer as ArrayBuffer }], OPTS)
    expect(s.errors[0]?.code).toBe('NOT_XLSX')
  })

  test('Group column dropped', () => {
    const header = [...ITEM_HEADER_TH, 'กลุ่ม']
    const row = [...itemRow('19/08/2026', '12:00', 'A1'), 'G1']
    const s = readSlot('items', [itemsFile('g.xlsx', 'DCI', [row], header)], OPTS)
    expect(s.errors).toEqual([])
    expect(s.rows[0]!.length).toBe(32)
    expect(s.rows[0]).not.toContain('G1')
  })

  test('missing optional columns -> warning and "-"', () => {
    const drop = new Set(['เคาน์เตอร์', 'หมายเหตุการชำระเงิน'])
    const keep = ITEM_HEADER_TH.map((h, i) => ({ h, i })).filter((x) => !drop.has(x.h))
    const row = itemRow('19/08/2026', '12:00', 'A1')
    const s = readSlot('items', [itemsFile('o.xlsx', 'DCI', [keep.map((k) => row[k.i]!)], keep.map((k) => k.h))], OPTS)
    expect(s.errors).toEqual([])
    const w = s.warnings.find((x) => x.code === 'MISSING_OPTIONAL_COLUMNS')
    expect(w?.params?.columns).toBe('เคาน์เตอร์, หมายเหตุการชำระเงิน')
    expect(s.rows[0]![3]).toBe('-')
    expect(s.rows[0]![22]).toBe('-')
    expect(s.rows[0]![4]).toBe('A1')
  })

  test('other missing column -> HEADER_NOT_FOUND', () => {
    const idx = ITEM_HEADER_TH.indexOf('ชื่อรายการ')
    const header = ITEM_HEADER_TH.filter((_, i) => i !== idx)
    const row = itemRow('19/08/2026', '12:00', 'A1').filter((_, i) => i !== idx)
    const s = readSlot('items', [itemsFile('m.xlsx', 'DCI', [row], header)], OPTS)
    expect(s.errors[0]?.code).toBe('HEADER_NOT_FOUND')
    expect(s.errors[0]?.params?.missing).toBe('ชื่อรายการ')
  })

  test('merge + sort + dedupe', () => {
    const a = itemsFile('a.xlsx', 'DCI', [itemRow('20/08/2026', '09:00', 'A2'), itemRow('19/08/2026', '10:00', 'A1')])
    const b = itemsFile('b.xlsx', 'DCI', [itemRow('20/08/2026', '09:00', 'A2'), itemRow('19/08/2026', '08:00', 'A0')])
    const s = readSlot('items', [a, b], OPTS)
    expect(s.rows.map((r) => r[4])).toEqual(['A0', 'A1', 'A2'])
    expect(s.duplicatesRemoved).toBe(1)
    const s2 = readSlot('items', [a, b], { ...OPTS, dedupe: false })
    expect(s2.rows.length).toBe(4)
    expect(s2.duplicatesRemoved).toBe(0)
  })

  test('em dash replaced (or kept when option off)', () => {
    const f = itemsFile('e.xlsx', 'DCI', [itemRow('19/08/2026', '12:00', 'A1', 'a—b')])
    expect(readSlot('items', [f], OPTS).rows[0]![8]).toBe('a-b')
    expect(readSlot('items', [f], { ...OPTS, replaceEmDash: false }).rows[0]![8]).toBe('a—b')
  })

  test('date gap and overlap warnings', () => {
    const a = itemsFile('a.xlsx', 'DCI', [itemRow('01/09/2026', '10:00', 'A'), itemRow('10/09/2026', '10:00', 'B')])
    const gap = itemsFile('g.xlsx', 'DCI', [itemRow('15/09/2026', '10:00', 'C'), itemRow('20/09/2026', '10:00', 'D')])
    const ov = itemsFile('o.xlsx', 'DCI', [itemRow('08/09/2026', '10:00', 'E'), itemRow('12/09/2026', '10:00', 'F')])
    const g = readSlot('items', [gap, a], OPTS)
    const gw = g.warnings.find((w) => w.code === 'DATE_GAP')
    expect(gw?.params).toEqual({ from: '11/09/2026', to: '14/09/2026' })
    expect(g.from).toBe('01/09/2026')
    expect(g.to).toBe('20/09/2026')
    const o = readSlot('items', [a, ov], OPTS)
    expect(o.warnings.find((w) => w.code === 'DATE_OVERLAP')?.params).toEqual({ from: '08/09/2026', to: '10/09/2026' })
    // adjacent days: neither
    const adj = itemsFile('adj.xlsx', 'DCI', [itemRow('11/09/2026', '10:00', 'G')])
    const n = readSlot('items', [a, adj], OPTS)
    expect(n.warnings.filter((w) => w.code === 'DATE_GAP' || w.code === 'DATE_OVERLAP')).toEqual([])
  })

  test('slot org mismatch', () => {
    const a = itemsFile('a.xlsx', 'DCI', [itemRow('19/08/2026', '10:00', 'A')])
    const b = itemsFile('b.xlsx', 'HIT', [itemRow('20/08/2026', '10:00', 'B')])
    const s = readSlot('items', [a, b], OPTS)
    expect(s.errors.find((e) => e.code === 'ORG_MISMATCH')?.params?.orgs).toBe('DCI, HIT')
  })
})

describe('topups', () => {
  test('reads, parses org, sorts by date/time', () => {
    const f = topupFile('t.xlsx', 'DCI', [tRow('26/10/2026, 08:00', 'R2'), tRow('25/10/2026, 20:28', 'R1')])
    const s = readSlot('topups', [f], OPTS)
    expect(s.errors).toEqual([])
    expect(s.org).toBe('DCI')
    expect(s.rows.map((r) => r[1])).toEqual(['R1', 'R2'])
    expect(s.rows[0]!.length).toBe(TOPUP_COLUMNS.length)
    expect(s.from).toBe('25/10/2026')
    expect(s.to).toBe('26/10/2026')
  })

  test('Thai top-up -> TOPUP_NOT_ENGLISH', () => {
    const f = xlsx('th.xlsx', [['x'], ['วันเวลาที่ชำระ', 'เลขที่ใบเสร็จ', 'เลขบัตร', 'ประเภท', 'อุปกรณ์', 'ผู้เติมเงิน', 'รหัสนักเรียน', 'ผู้รับเงิน', 'ประเภทกระเป๋าเงิน', 'ช่องทางชำระเงิน', 'จำนวนเงิน']])
    expect(readSlot('topups', [f], OPTS).errors[0]?.code).toBe('TOPUP_NOT_ENGLISH')
  })
})

describe('cards', () => {
  test('14-col, header row 1, org null', () => {
    const f = xlsx('c.xlsx', [CARD_H, cRow('S1', 'สมชาย', 'ใจดี', 'Somchai', 'Jaidee')])
    const s = readSlot('cards', [f], OPTS)
    expect(s.errors).toEqual([])
    expect(s.org).toBeNull()
    expect(s.rows[0]!.length).toBe(14)
    expect(s.from).toBeNull()
  })

  test('9-col member file mapped (tolerates missing " *")', () => {
    const f = xlsx('m.xlsx', [
      ['รหัสสมาชิก *', 'ชื่อ (TH) *', 'นามสกุล (TH)', ' ชื่อ (EN) *', 'นามสกุล (EN) *', 'ประเภท', 'สถานะ', 'เลขบัตร', 'วันหมดอายุบัตร'],
      ['M1', 'ครู', 'เอ', 'Teacher', 'A', 'ครู', 'ใช้งาน', '999', '01/01/2030'],
    ])
    const s = readSlot('cards', [f], OPTS)
    expect(s.errors).toEqual([])
    expect(s.warnings.some((w) => w.code === 'MEMBER_FILE_MAPPED')).toBe(true)
    expect(s.rows[0]).toEqual(['M1', 'ครู', 'เอ', 'Teacher', 'A', '', '', '', '', '', 'ครู', 'ใช้งาน', '999', '01/01/2030'])
  })

  test('students + members merged', () => {
    const a = xlsx('s.xlsx', [CARD_H, cRow('S1', 'ก', 'ข', 'A', 'B')])
    const m = xlsx('m.xlsx', [['รหัสสมาชิก *', 'ชื่อ (TH) *', 'นามสกุล (TH) *', 'ชื่อ (EN) *', 'นามสกุล (EN) *', 'ประเภท', 'สถานะ', 'เลขบัตร', 'วันหมดอายุบัตร'], ['M1', 'ค', 'ง', 'C', 'D', 'ครู', 'x', '1', '2']])
    expect(readSlot('cards', [a, m], OPTS).rows.length).toBe(2)
  })

  test('duplicate names warning (TH and EN, NBSP tolerant)', () => {
    const f = xlsx('d.xlsx', [
      CARD_H,
      cRow('S1', 'สมชาย', 'ใจดี', 'Som', 'One'),
      cRow('S2', 'สมชาย ', 'ใจดี', 'Som', 'Two'),
      cRow('S3', 'อื่น', 'คน', 'Sam', 'Lee'),
      cRow('S4', 'อีก', 'คน', 'Sam', 'Lee'),
      cRow('S5', '', '', '', ''),
      cRow('S6', '', '', '', ''),
    ])
    const s = readSlot('cards', [f], OPTS)
    const w = s.warnings.find((x) => x.code === 'DUPLICATE_NAMES')
    expect(w?.params?.names).toBe('สมชาย ใจดี, Sam Lee')
    expect(w?.params?.count).toBe(2)
  })
})

describe('crossCheck', () => {
  test('org mismatch between items and topups', () => {
    const i = readSlot('items', [itemsFile('a.xlsx', 'DCI', [itemRow('17/08/2026', '10:00', 'A')])], OPTS)
    const t = readSlot('topups', [topupFile('t.xlsx', 'HIT', [tRow('17/08/2026, 10:00')])], OPTS)
    const r = crossCheck({ items: i, topups: t })
    expect(r.errors[0]?.code).toBe('ORG_MISMATCH')
    expect(r.errors[0]?.params?.orgs).toBe('DCI, HIT')
    expect(r.warnings).toEqual([])
  })

  test('start-date warnings', () => {
    const late = readSlot('items', [itemsFile('a.xlsx', 'DCI', [itemRow('19/08/2026', '10:00', 'A')])], OPTS)
    const early = readSlot('topups', [topupFile('t.xlsx', 'DCI', [tRow('01/08/2026, 10:00')])], OPTS)
    const r = crossCheck({ items: late, topups: early })
    expect(r.errors).toEqual([])
    expect(r.warnings).toEqual([
      { code: 'STARTS_AFTER_TEMPLATE_START', params: { kind: 'items', from: '19/08/2026' } },
      { code: 'BEFORE_TEMPLATE_START', params: { kind: 'topups', from: '01/08/2026' } },
    ])
    const ok = readSlot('items', [itemsFile('a.xlsx', 'DCI', [itemRow('17/08/2026', '10:00', 'A')])], OPTS)
    expect(crossCheck({ items: ok }).warnings).toEqual([])
  })
})
