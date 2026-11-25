<script setup lang="ts">
import { computed, nextTick, onBeforeUnmount, ref, shallowRef, watch } from 'vue'
import type { ErrorCode, Issue, ParsedSlot, ReadOptions, TableKey, TemplateInfo, WarningCode } from './types'
import { crossCheck, readSlot, type InputFile } from './lib/readExports'
import { fillTemplate, inspectTemplate, TemplateError } from './lib/fillTemplate'
import { useI18n } from './composables/useI18n'
import type { MsgKey } from './i18n'
import DropZone, { type ZoneStatus } from './components/DropZone.vue'
import MessageList, { type UiMessage } from './components/MessageList.vue'
import BuildProgress from './components/BuildProgress.vue'
import DoneView, { type DoneSummary } from './components/DoneView.vue'

type View = 'form' | 'building' | 'done' | 'failed'
type AnyIssue = { level: 'error'; issue: Issue<ErrorCode> } | { level: 'warning'; issue: Issue<WarningCode> }

const { lang, setLang, t } = useI18n()

// ---------- state ----------
const view = ref<View>('form')
const templateFile = shallowRef<InputFile | null>(null)
const templateInfo = shallowRef<TemplateInfo | null>(null)
const inputs = shallowRef<Record<TableKey, InputFile[]>>({ cards: [], topups: [], items: [] })
const slots = shallowRef<Partial<Record<TableKey, ParsedSlot>>>({})
const replaceEmDash = ref(true)
const dedupe = ref(true)
const busy = ref<Record<TableKey | 'template', boolean>>({ template: false, cards: false, topups: false, items: false })

const buildStep = ref(0)
const buildDetail = ref('')
const failMessage = ref('')
const resultUrl = ref('')
const resultName = ref('')
const doneSummary = ref<DoneSummary>({ cards: 0, topups: 0, items: 0, range: '' })
const doneDuplicateIssue = ref<Issue<WarningCode> | null>(null)

// ---------- helpers ----------
const tick = (): Promise<void> => new Promise((r) => setTimeout(r, 0))
const options = (): ReadOptions => ({ replaceEmDash: replaceEmDash.value, dedupe: dedupe.value })

function failedSlot(kind: TableKey, files: InputFile[]): ParsedSlot {
  return {
    kind,
    files: files.map((f) => f.name),
    rows: [],
    org: null,
    from: null,
    to: null,
    duplicatesRemoved: 0,
    errors: [{ code: 'NOT_XLSX', params: { file: files[0]?.name ?? '' } }],
    warnings: [],
  }
}

function parseKind(kind: TableKey): void {
  const files = inputs.value[kind]
  const next = { ...slots.value }
  if (files.length === 0) {
    delete next[kind]
  } else {
    try {
      next[kind] = readSlot(kind, files, options())
    } catch {
      next[kind] = failedSlot(kind, files)
    }
  }
  slots.value = next
}

async function addFiles(kind: TableKey, picked: File[]): Promise<void> {
  busy.value[kind] = true
  try {
    const loaded: InputFile[] = []
    for (const f of picked) loaded.push({ name: f.name, data: await f.arrayBuffer() })
    const multiple = kind !== 'topups'
    const existing = multiple ? inputs.value[kind].filter((e) => !loaded.some((l) => l.name === e.name)) : []
    inputs.value = { ...inputs.value, [kind]: [...existing, ...loaded] }
    await nextTick()
    await tick()
    parseKind(kind)
  } finally {
    busy.value[kind] = false
  }
}

async function addTemplate(picked: File[]): Promise<void> {
  const f = picked[0]
  if (!f) return
  busy.value.template = true
  try {
    const data = await f.arrayBuffer()
    templateFile.value = { name: f.name, data }
    templateInfo.value = null
    await nextTick()
    await tick()
    try {
      templateInfo.value = await inspectTemplate(f.name, data)
    } catch {
      templateInfo.value = {
        fileName: f.name,
        sizeBytes: data.byteLength,
        sheetCount: 0,
        tables: [],
        errors: [{ code: 'NOT_XLSX', params: { file: f.name } }],
      }
    }
  } finally {
    busy.value.template = false
  }
}

function removeKind(kind: TableKey): void {
  inputs.value = { ...inputs.value, [kind]: [] }
  parseKind(kind)
}

function removeTemplate(): void {
  templateFile.value = null
  templateInfo.value = null
}

function revokeResult(): void {
  if (resultUrl.value) URL.revokeObjectURL(resultUrl.value)
  resultUrl.value = ''
}

function clearAll(): void {
  revokeResult()
  templateFile.value = null
  templateInfo.value = null
  inputs.value = { cards: [], topups: [], items: [] }
  slots.value = {}
}

function startOver(): void {
  clearAll()
  view.value = 'form'
}

watch([replaceEmDash, dedupe], async () => {
  const kinds: TableKey[] = ['cards', 'topups', 'items']
  const active = kinds.filter((k) => inputs.value[k].length > 0)
  if (active.length === 0) return
  for (const k of active) busy.value[k] = true
  await nextTick()
  await tick()
  for (const k of active) {
    parseKind(k)
    busy.value[k] = false
  }
})

onBeforeUnmount(revokeResult)

// ---------- derived ----------
const cross = computed(() => crossCheck(slots.value))

function paramText(key: string, value: string | number): string | number {
  if (key === 'kind' && (value === 'cards' || value === 'topups' || value === 'items')) return t(`kind.${value}` as MsgKey)
  return value
}

function issueText(prefix: 'err' | 'warn', issue: Issue<string>): string {
  const params: Record<string, string | number> = {}
  for (const [k, v] of Object.entries(issue.params ?? {})) params[k] = paramText(k, v)
  return t(`${prefix}.${issue.code}` as MsgKey, params)
}

const allIssues = computed<AnyIssue[]>(() => {
  const out: AnyIssue[] = []
  for (const e of templateInfo.value?.errors ?? []) out.push({ level: 'error', issue: e })
  for (const s of Object.values(slots.value)) {
    for (const e of s.errors) out.push({ level: 'error', issue: e })
  }
  for (const e of cross.value.errors) out.push({ level: 'error', issue: e })
  for (const s of Object.values(slots.value)) {
    for (const w of s.warnings) out.push({ level: 'warning', issue: w })
  }
  for (const w of cross.value.warnings) out.push({ level: 'warning', issue: w })
  const seen = new Set<string>()
  return out.filter((i) => {
    const key = `${i.level}:${i.issue.code}:${JSON.stringify(i.issue.params ?? {})}`
    if (seen.has(key)) return false
    seen.add(key)
    return true
  })
})

const messageItems = computed<UiMessage[]>(() => {
  void lang.value
  return allIssues.value.map((i) => ({
    level: i.level,
    text: issueText(i.level === 'error' ? 'err' : 'warn', i.issue),
  }))
})

const hasErrors = computed(() => allIssues.value.some((i) => i.level === 'error'))
const hasWarnings = computed(() => allIssues.value.some((i) => i.level === 'warning'))

const canBuild = computed(
  () =>
    templateInfo.value !== null &&
    Boolean(slots.value.cards && slots.value.topups && slots.value.items) &&
    !hasErrors.value,
)

const anyLoaded = computed(
  () => templateFile.value !== null || inputs.value.cards.length + inputs.value.topups.length + inputs.value.items.length > 0,
)

function joinList(items: string[]): string {
  if (items.length <= 1) return items.join('')
  return items.slice(0, -1).join(', ') + t('tpl.and') + items[items.length - 1]
}

function slotZone(kind: TableKey): { status: ZoneStatus; names: string[]; summary: string; message: string } {
  void lang.value
  const names = inputs.value[kind].map((f) => f.name)
  const s = slots.value[kind]
  if (!s) return { status: 'empty', names, summary: '', message: '' }
  const dash = (v: string | null): string => v ?? '-'
  let summary = ''
  if (kind === 'cards') summary = t('cards.summary', { n: s.rows.length.toLocaleString('en-US') })
  else if (kind === 'topups')
    summary = t('topups.summary', { n: s.rows.length.toLocaleString('en-US'), from: dash(s.from), to: dash(s.to), org: dash(s.org) })
  else
    summary = t('items.summary', {
      files: s.files.length,
      n: s.rows.length.toLocaleString('en-US'),
      from: dash(s.from),
      to: dash(s.to),
      org: dash(s.org),
    })
  const err = s.errors[0]
  const warn = s.warnings[0]
  if (err) return { status: 'error', names, summary, message: issueText('err', err) }
  if (warn) return { status: 'warning', names, summary, message: issueText('warn', warn) }
  return { status: 'ok', names, summary, message: '' }
}

const templateZone = computed(() => {
  void lang.value
  const f = templateFile.value
  const info = templateInfo.value
  const names = f ? [f.name] : []
  if (!f || !info) return { status: 'empty' as ZoneStatus, names, summary: '', message: '' }
  const summary = t('tpl.summary', { tables: joinList(info.tables), n: info.sheetCount })
  const err = info.errors[0]
  if (err) return { status: 'error' as ZoneStatus, names, summary, message: issueText('err', err) }
  return { status: 'ok' as ZoneStatus, names, summary, message: '' }
})

const cardsZone = computed(() => slotZone('cards'))
const topupsZone = computed(() => slotZone('topups'))
const itemsZone = computed(() => slotZone('items'))

// ---------- build ----------
async function build(): Promise<void> {
  const tpl = templateFile.value
  const cards = slots.value.cards
  const topups = slots.value.topups
  const items = slots.value.items
  if (!tpl || !cards || !topups || !items) return

  revokeResult()
  buildStep.value = 0
  buildDetail.value = ''
  view.value = 'building'
  await nextTick()
  await tick()
  await tick()

  const from = items.from ?? topups.from ?? ''
  const to = items.to ?? topups.to ?? ''
  try {
    const result = await fillTemplate({
      template: tpl.data,
      templateName: tpl.name,
      cards: cards.rows,
      topups: topups.rows,
      items: items.rows,
      org: items.org ?? topups.org ?? 'DCI',
      from,
      to,
      onStep: (step, detail) => {
        buildStep.value = step + 1
        buildDetail.value = detail ?? ''
      },
    })
    resultUrl.value = URL.createObjectURL(result.blob)
    resultName.value = result.fileName
    doneSummary.value = {
      cards: cards.rows.length,
      topups: topups.rows.length,
      items: items.rows.length,
      range: from && to ? `${from} – ${to}` : '-',
    }
    doneDuplicateIssue.value = null
    for (const i of allIssues.value) {
      if (i.level === 'warning' && i.issue.code === 'DUPLICATE_NAMES') doneDuplicateIssue.value = i.issue
    }
    view.value = 'done'
  } catch (e) {
    failMessage.value = e instanceof TemplateError ? issueText('err', e.issue) : e instanceof Error ? e.message : String(e)
    view.value = 'failed'
  }
}

const duplicateMessage = computed(() => {
  void lang.value
  return doneDuplicateIssue.value ? issueText('warn', doneDuplicateIssue.value) : ''
})
</script>

<template>
  <header class="topbar">
    <div class="wrap topbar__in">
      <span class="topbar__brand">{{ t('app.name') }}</span>
      <div class="seg" role="group" :aria-label="t('lang.label')">
        <button type="button" class="seg__btn" :aria-pressed="lang === 'th'" lang="th" @click="setLang('th')">TH</button>
        <button type="button" class="seg__btn" :aria-pressed="lang === 'en'" lang="en" @click="setLang('en')">EN</button>
      </div>
    </div>
  </header>

  <main class="wrap main">
    <template v-if="view === 'form'">
      <h1 class="main__title">{{ t('title') }}</h1>
      <p class="main__intro">{{ t('intro') }}</p>

      <div class="grid">
        <DropZone
          :title="t('tpl.title')"
          :hint="t('tpl.hint')"
          accept=".xlsm,.xlsx"
          :status="templateZone.status"
          :file-names="templateZone.names"
          :summary="templateZone.summary"
          :message="templateZone.message"
          :busy="busy.template"
          @files="addTemplate"
          @remove="removeTemplate"
        />
        <DropZone
          :title="t('cards.title')"
          :hint="t('cards.hint')"
          accept=".xlsx,.xls"
          multiple
          :status="cardsZone.status"
          :file-names="cardsZone.names"
          :summary="cardsZone.summary"
          :message="cardsZone.message"
          :busy="busy.cards"
          @files="(f) => addFiles('cards', f)"
          @remove="removeKind('cards')"
        />
        <DropZone
          :title="t('topups.title')"
          :hint="t('topups.hint')"
          accept=".xlsx,.xls"
          :status="topupsZone.status"
          :file-names="topupsZone.names"
          :summary="topupsZone.summary"
          :message="topupsZone.message"
          :busy="busy.topups"
          @files="(f) => addFiles('topups', f)"
          @remove="removeKind('topups')"
        />
        <DropZone
          :title="t('items.title')"
          :hint="t('items.hint')"
          accept=".xlsx,.xls"
          multiple
          :status="itemsZone.status"
          :file-names="itemsZone.names"
          :summary="itemsZone.summary"
          :message="itemsZone.message"
          :busy="busy.items"
          @files="(f) => addFiles('items', f)"
          @remove="removeKind('items')"
        />
      </div>

      <fieldset class="opts">
        <legend>{{ t('opt.title') }}</legend>
        <label class="opts__row"><input v-model="replaceEmDash" type="checkbox" /> <span>{{ t('opt.emdash') }}</span></label>
        <label class="opts__row"><input v-model="dedupe" type="checkbox" /> <span>{{ t('opt.dedupe') }}</span></label>
      </fieldset>

      <MessageList :items="messageItems" />

      <div class="actions">
        <button type="button" class="btn btn--ghost" :disabled="!anyLoaded" @click="clearAll">{{ t('btn.clear') }}</button>
        <button type="button" class="btn btn--primary" :disabled="!canBuild" @click="build">
          {{ hasWarnings ? t('btn.buildAnyway') : t('btn.build') }}
        </button>
      </div>
    </template>

    <BuildProgress v-else-if="view === 'building'" :current="buildStep" :detail="buildDetail" />

    <DoneView
      v-else-if="view === 'done'"
      :summary="doneSummary"
      :file-name="resultName"
      :download-url="resultUrl"
      :duplicate-names-message="duplicateMessage"
      @reset="startOver"
    />

    <section v-else class="card failed" role="alert">
      <h2 class="failed__title">{{ t('fail.title') }}</h2>
      <p class="failed__msg">{{ failMessage }}</p>
      <button type="button" class="btn btn--primary" @click="view = 'form'">{{ t('fail.back') }}</button>
    </section>
  </main>

  <footer class="wrap foot">{{ t('footer') }}</footer>
</template>

<style scoped>
.wrap { width: 100%; max-width: 960px; margin: 0 auto; padding-left: 16px; padding-right: 16px; }
.topbar { background: var(--surface); border-bottom: 1px solid var(--border); }
.topbar__in { display: flex; align-items: center; justify-content: space-between; gap: 12px; min-height: 56px; }
.topbar__brand { font-weight: 700; font-size: 1.0625rem; }
.seg { display: inline-flex; border: 1px solid var(--border-strong); border-radius: var(--radius-sm); overflow: hidden; }
.seg__btn { font: inherit; font-size: 0.875rem; font-weight: 600; min-width: 48px; min-height: 36px; padding: 4px 12px; color: var(--text); background: transparent; border: 0; cursor: pointer; }
.seg__btn + .seg__btn { border-left: 1px solid var(--border-strong); }
.seg__btn[aria-pressed='true'] { background: var(--accent); color: #fff; }
.main { padding-top: 28px; padding-bottom: 24px; display: flex; flex-direction: column; gap: 20px; }
.main__title { margin: 0; font-size: 1.75rem; line-height: 1.3; }
.main__intro { margin: -8px 0 0; color: var(--muted); max-width: 70ch; }
.grid { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; }
@media (max-width: 639px) { .grid { grid-template-columns: 1fr; } }
.opts { margin: 0; padding: 12px 16px; border: 1px solid var(--border); border-radius: var(--radius); background: var(--surface); display: flex; flex-direction: column; gap: 8px; }
.opts legend { padding: 0 6px; font-weight: 600; }
.opts__row { display: flex; align-items: flex-start; gap: 10px; cursor: pointer; }
.opts__row input { width: 20px; height: 20px; margin-top: 2px; flex: none; accent-color: var(--accent); }
.actions { display: flex; flex-wrap: wrap; justify-content: flex-end; gap: 12px; }
.failed__title { margin: 0 0 8px; font-size: 1.25rem; color: var(--err-text); }
.failed__msg { margin: 0 0 16px; overflow-wrap: anywhere; }
.foot { padding-top: 8px; padding-bottom: 32px; color: var(--muted); font-size: 0.8125rem; }
</style>
