<script setup lang="ts">
import { useI18n } from '../composables/useI18n'

export interface DoneSummary {
  cards: number
  topups: number
  items: number
  range: string
}

defineProps<{
  summary: DoneSummary
  fileName: string
  downloadUrl: string
  duplicateNamesMessage?: string
}>()
const emit = defineEmits<{ reset: [] }>()
const { t } = useI18n()
const fmt = (n: number): string => n.toLocaleString('en-US')
</script>

<template>
  <section class="card done" aria-live="polite">
    <h2 class="done__title">{{ t('done.title') }}</h2>

    <dl class="tiles">
      <div class="tile"><dt>{{ t('done.cards') }}</dt><dd>{{ fmt(summary.cards) }}</dd></div>
      <div class="tile"><dt>{{ t('done.topups') }}</dt><dd>{{ fmt(summary.topups) }}</dd></div>
      <div class="tile"><dt>{{ t('done.items') }}</dt><dd>{{ fmt(summary.items) }}</dd></div>
      <div class="tile"><dt>{{ t('done.range') }}</dt><dd class="tile__range">{{ summary.range }}</dd></div>
    </dl>

    <p v-if="duplicateNamesMessage" class="done__warn">
      <strong>{{ t('msg.warning') }}</strong> {{ duplicateNamesMessage }}
    </p>

    <div class="done__actions">
      <a class="btn btn--primary" :href="downloadUrl" :download="fileName">{{ t('done.download', { fileName }) }}</a>
      <button type="button" class="btn btn--ghost" @click="emit('reset')">{{ t('done.again') }}</button>
    </div>
    <p class="done__note">{{ t('done.note') }}</p>
  </section>
</template>

<style scoped>
.done__title { margin: 0 0 16px; font-size: 1.25rem; }
.tiles { display: grid; grid-template-columns: repeat(auto-fit, minmax(140px, 1fr)); gap: 12px; margin: 0 0 16px; }
.tile { margin: 0; padding: 12px; background: var(--bg); border: 1px solid var(--border); border-radius: var(--radius-sm); }
.tile dt { font-size: 0.8125rem; color: var(--muted); }
.tile dd { margin: 4px 0 0; font-size: 1.375rem; font-weight: 700; }
.tile__range { font-size: 1rem !important; overflow-wrap: anywhere; }
.done__warn { margin: 0 0 16px; padding: 10px 12px; background: var(--warn-bg); border: 1px solid var(--warn-border); color: var(--warn-text); border-radius: var(--radius-sm); }
.done__actions { display: flex; flex-wrap: wrap; gap: 12px; margin-bottom: 12px; }
.done__note { margin: 0; color: var(--muted); font-size: 0.875rem; }
</style>
