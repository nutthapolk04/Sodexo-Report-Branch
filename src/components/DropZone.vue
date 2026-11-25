<script setup lang="ts">
import { computed, ref, useId } from 'vue'
import { useI18n } from '../composables/useI18n'

export type ZoneStatus = 'empty' | 'ok' | 'warning' | 'error'

const props = defineProps<{
  title: string
  hint: string
  accept: string
  multiple?: boolean
  status: ZoneStatus
  fileNames: string[]
  summary?: string
  message?: string
  busy?: boolean
}>()

const emit = defineEmits<{
  files: [files: File[]]
  remove: []
}>()

const { t, richParts } = useI18n()
const uid = useId()
const input = ref<HTMLInputElement | null>(null)
const dragging = ref(false)

const hintParts = computed(() => richParts(props.hint))
const hasFiles = computed(() => props.fileNames.length > 0)
const icon = computed(() => ({ empty: '+', ok: '✓', warning: '!', error: '✕' })[props.status])

function open(): void {
  input.value?.click()
}

function take(list: FileList | null): void {
  if (!list || list.length === 0) return
  const files = Array.from(list)
  emit('files', props.multiple ? files : files.slice(0, 1))
}

function onChange(e: Event): void {
  const el = e.target as HTMLInputElement
  take(el.files)
  el.value = ''
}

function onDrop(e: DragEvent): void {
  dragging.value = false
  take(e.dataTransfer?.files ?? null)
}
</script>

<template>
  <section
    class="zone"
    :class="[`zone--${status}`, { 'zone--drag': dragging }]"
    :aria-labelledby="`${uid}-t`"
    @dragenter.prevent="dragging = true"
    @dragover.prevent="dragging = true"
    @dragleave.prevent="dragging = false"
    @drop.prevent="onDrop"
  >
    <h2 :id="`${uid}-t`" class="zone__title">{{ title }}</h2>
    <p :id="`${uid}-h`" class="zone__hint">
      <template v-for="(p, i) in hintParts" :key="i">
        <strong v-if="p.bold">{{ p.text }}</strong>
        <template v-else>{{ p.text }}</template>
      </template>
    </p>

    <button type="button" class="zone__target" :aria-describedby="`${uid}-h`" @click="open">
      <span class="zone__icon" aria-hidden="true">{{ icon }}</span>
      <span class="zone__body">
        <span v-if="busy" class="zone__prompt">{{ t('zone.reading') }}</span>
        <span v-else-if="!hasFiles" class="zone__prompt">{{ multiple ? t('zone.dropMany') : t('zone.drop') }}</span>
        <template v-else>
          <span v-for="n in fileNames" :key="n" class="zone__file">{{ n }}</span>
          <span v-if="summary && status !== 'error'" class="zone__summary">{{ summary }}</span>
        </template>
      </span>
    </button>

    <p class="zone__msg" aria-live="polite">{{ message }}</p>

    <button
      v-if="hasFiles"
      type="button"
      class="zone__remove"
      :aria-label="t('zone.removeFor', { title })"
      @click="emit('remove')"
    >
      {{ t('zone.remove') }}
    </button>

    <input ref="input" class="zone__input" type="file" :accept="accept" :multiple="multiple" tabindex="-1" aria-hidden="true" @change="onChange" />
  </section>
</template>

<style scoped>
.zone {
  display: flex;
  flex-direction: column;
  gap: 8px;
  padding: 16px;
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: var(--radius);
  min-width: 0;
}
.zone--drag {
  border-color: var(--accent);
  box-shadow: 0 0 0 3px var(--accent-soft);
}
.zone--ok { border-color: var(--ok-border); }
.zone--warning { border-color: var(--warn-border); }
.zone--error { border-color: var(--err-border); }
.zone__title { margin: 0; font-size: 1rem; font-weight: 600; }
.zone__hint { margin: 0; color: var(--muted); font-size: 0.875rem; }
.zone__target {
  display: flex;
  align-items: center;
  gap: 12px;
  min-height: 72px;
  padding: 12px;
  text-align: left;
  font: inherit;
  color: var(--text);
  background: var(--bg);
  border: 2px dashed var(--border-strong);
  border-radius: var(--radius-sm);
  cursor: pointer;
  width: 100%;
}
.zone__target:hover { border-color: var(--accent); }
.zone--ok .zone__target { background: var(--ok-bg); border: 2px solid var(--ok-border); }
.zone--warning .zone__target { background: var(--warn-bg); border: 2px solid var(--warn-border); }
.zone--error .zone__target { background: var(--err-bg); border: 2px solid var(--err-border); }
.zone--drag .zone__target { border-color: var(--accent); background: var(--accent-soft); }
.zone__icon {
  flex: none;
  display: grid;
  place-items: center;
  width: 28px;
  height: 28px;
  border-radius: 50%;
  font-weight: 700;
  background: var(--border);
  color: var(--text);
}
.zone--ok .zone__icon { background: var(--ok-text); color: var(--surface); }
.zone--warning .zone__icon { background: var(--warn-text); color: var(--surface); }
.zone--error .zone__icon { background: var(--err-text); color: var(--surface); }
.zone__body { display: flex; flex-direction: column; gap: 2px; min-width: 0; }
.zone__prompt { color: var(--muted); }
.zone__file { font-weight: 600; overflow-wrap: anywhere; }
.zone__summary { font-size: 0.875rem; color: var(--muted); overflow-wrap: anywhere; }
.zone__msg { margin: 0; font-size: 0.875rem; min-height: 0; }
.zone__msg:empty { display: none; }
.zone--error .zone__msg { color: var(--err-text); }
.zone--warning .zone__msg { color: var(--warn-text); }
.zone__remove {
  align-self: flex-start;
  font: inherit;
  font-size: 0.875rem;
  padding: 4px 12px;
  color: var(--text);
  background: transparent;
  border: 1px solid var(--border-strong);
  border-radius: var(--radius-sm);
  cursor: pointer;
}
.zone__remove:hover { background: var(--bg); }
.zone__input { display: none; }
</style>
