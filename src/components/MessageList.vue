<script setup lang="ts">
import { useI18n } from '../composables/useI18n'

export interface UiMessage {
  level: 'error' | 'warning'
  text: string
}

defineProps<{ items: UiMessage[] }>()
const { t } = useI18n()
</script>

<template>
  <div class="msgs" aria-live="polite">
    <ul v-if="items.length" class="msgs__list" :aria-label="t('msg.title')">
      <li v-for="(m, i) in items" :key="i" class="msgs__item" :class="`msgs__item--${m.level}`">
        <strong class="msgs__tag">{{ m.level === 'error' ? t('msg.error') : t('msg.warning') }}</strong>
        <span class="msgs__text">{{ m.text }}</span>
      </li>
    </ul>
  </div>
</template>

<style scoped>
.msgs__list { list-style: none; margin: 0; padding: 0; display: flex; flex-direction: column; gap: 8px; }
.msgs__item { display: flex; gap: 8px; padding: 10px 12px; border-radius: var(--radius-sm); border: 1px solid; font-size: 0.9375rem; }
.msgs__item--error { background: var(--err-bg); border-color: var(--err-border); color: var(--err-text); }
.msgs__item--warning { background: var(--warn-bg); border-color: var(--warn-border); color: var(--warn-text); }
.msgs__tag { flex: none; }
.msgs__text { overflow-wrap: anywhere; }
</style>
