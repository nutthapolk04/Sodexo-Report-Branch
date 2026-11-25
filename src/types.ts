// Shared contract between src/lib/readExports.ts, src/lib/fillTemplate.ts and the UI.
// Editing this file is a coordination event: every module depends on it.

export type Cell = string | number

/** One parsed export row, already ordered to the target table's column list. */
export type Row = Cell[]

export type TableKey = 'cards' | 'topups' | 'items'

/** Template table names (Excel displayName) per key. */
export const TABLE_NAMES: Record<TableKey, string> = {
  cards: 'WalletCards',
  topups: 'WalletTopups',
  items: 'WalletItems',
}

/** Columns the tool writes, in table order. WalletItems also has 4 calculated columns after these (AG:AJ) that the tool never writes. */
export const CARD_COLUMNS = [
  'รหัสนักเรียน/สมาชิก *', 'ชื่อ (TH) *', 'นามสกุล (TH) *', 'ชื่อ (EN) *', 'นามสกุล (EN) *',
  'รหัสชั้นปี', 'ห้องเรียน', 'อีเมล', 'เบอร์โทร', 'ข้อมูลแพ้อาหาร', 'ประเภทบัตร', 'สถานะ', 'เลขบัตร', 'วันหมดอายุบัตร',
] as const

export const TOPUP_COLUMNS = [
  'Date/Time', 'Receipt No.', 'Card ID', 'Type', 'Device', 'Topped Up By', 'Student ID', 'Recipient', 'Wallet Type', 'Payment Method', 'Amount',
] as const

/** 'ประเภท2' is the export's second "ประเภท" header (the category column right after เลขที่ใบเสร็จ). */
export const ITEM_COLUMNS = [
  'ประเภท', 'วันที่', 'เวลา', 'เคาน์เตอร์', 'เลขที่ใบเสร็จ', 'ประเภท2', 'ช่วงเวลา', 'รหัสสินค้า', 'ชื่อรายการ', 'จำนวน', 'ราคา',
  'ตัวเลือกเสริม', 'ราคาตัวเลือกเสริม', 'ยอดรวมสินค้า', 'ส่วนลด', 'ค่าบริการ', 'ภาษีมูลค่าเพิ่ม', 'ยอดสุทธิ', 'เลขอ้างอิง', 'เลขที่ธุรกรรม',
  'ช่องทางชำระเงิน', 'ชื่อการชำระเงิน', 'หมายเหตุการชำระเงิน', 'ชำระ', 'คงเหลือ', 'ยอดที่รับ', 'ชื่อลูกค้า', 'ประเภทผู้ใช้',
  'หมายเหตุออเดอร์', 'หมายเหตุรายการ', 'อ้างอิงถึง', 'สถานะ',
] as const

/** Template start date the customer's formulas hard-code (DATE(2026,8,17)). */
export const TEMPLATE_START = '2026-08-17'

export type ErrorCode =
  | 'NOT_XLSX'
  | 'HEADER_NOT_FOUND'
  | 'BILLS_NOT_TRANSACTIONS'
  | 'SALES_NOT_THAI'
  | 'TOPUP_NOT_ENGLISH'
  | 'ORG_MISMATCH'
  | 'TEMPLATE_TABLE_MISSING'
  | 'MEMBER_FILE_IN_STUDENTS' // 9-column member list dropped in the Students slot
  | 'STUDENT_FILE_IN_MEMBERS' // 14-column student list dropped in the Members slot

export type WarningCode =
  | 'DATE_GAP'          // gap between merged Sales files' date ranges
  | 'DATE_OVERLAP'      // overlap between merged Sales files (duplicates removed if option on)
  | 'STARTS_AFTER_TEMPLATE_START' // earliest Sales/Top-up row is after 17/08/2026 -> balances wrong
  | 'BEFORE_TEMPLATE_START'       // some row is before 17/08/2026
  | 'DUPLICATE_NAMES'   // same full TH or EN name on 2+ cards
  | 'MEMBER_FILE_MAPPED' // a 9-column member-page file was mapped into 14 columns
  | 'MISSING_OPTIONAL_COLUMNS' // older export without เคาน์เตอร์ / หมายเหตุการชำระเงิน; filled with "-"

export interface Issue<C extends string> {
  code: C
  /** Values for the i18n message, e.g. { file: 'x.xlsx', org: 'HIT', from: '19/08/2026', to: '20/10/2026', names: 'A, B' }. */
  params?: Record<string, string | number>
}

/** Result of reading one slot (one or more files of the same kind). */
export interface ParsedSlot {
  kind: TableKey
  files: string[]
  rows: Row[]
  /** Org read from the report header line "สาขา/องค์กร: X" / "องค์กร/สาขา: X" / "Org/Branch: X" / "Organization/Branch: X". null when absent (card files). */
  org: string | null
  /** dd/mm/yyyy of the earliest / latest row; null when no dated rows. */
  from: string | null
  to: string | null
  duplicatesRemoved: number
  errors: Issue<ErrorCode>[]
  warnings: Issue<WarningCode>[]
}

export interface ReadOptions {
  /** Replace "—" (U+2014) with "-" in every string cell. Default true. */
  replaceEmDash: boolean
  /** Drop exact duplicate rows across files. Default true. */
  dedupe: boolean
}

/** Summary of the template slot. */
export interface TemplateInfo {
  fileName: string
  sizeBytes: number
  sheetCount: number
  /** Table displayNames found. */
  tables: string[]
  errors: Issue<ErrorCode>[]
}

export interface FillInput {
  template: ArrayBuffer
  templateName: string
  cards: Row[]
  topups: Row[]
  items: Row[]
  org: string
  /** dd/mm/yyyy */
  from: string
  to: string
  /** Called after each build step (0..4) with optional progress text. */
  onStep?: (step: number, detail?: string) => void
}

export interface FillResult {
  blob: Blob
  fileName: string
}
