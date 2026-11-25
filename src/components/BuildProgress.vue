<script setup lang="ts">
import { computed } from 'vue'
import { useI18n } from '../composables/useI18n'
import type { MsgKey } from '../i18n'

const STEP_COUNT = 5

const props = defineProps<{
  /** Number of finished steps (0..5). Step at this index is the current one. */
  current: number
  detail?: string
}>()

const { t } = useI18n()
const steps = computed(() =>
  Array.from({ length: STEP_COUNT }, (_v, i) => ({
    label: t(`build.step${i}` as MsgKey),
    state: i < props.current ? 'done' : i === props.current ? 'current' : 'pending',
  })),
)
const percent = computed(() => Math.round((Math.min(props.current, STEP_COUNT) / STEP_COUNT) * 100))
</script>

<template>
  <section class="card progress" aria-live="polite">
    <h2 class="progress__title">{{ t('build.title') }}</h2>
    <div
      class="progress__bar"
      role="progressbar"
      :aria-label="t('build.progress')"
      aria-valuemin="0"
      aria-valuemax="100"
      :aria-valuenow="percent"
    >
      <div class="progress__fill" :style="{ width: `${percent}%` }" />
    </div>
    <ol class="progress__steps">
      <li v-for="(s, i) in steps" :key="i" class="step" :class="`step--${s.state}`" :aria-current="s.state === 'current' ? 'step' : undefined">
        <span class="step__mark" aria-hidden="true">{{ s.state === 'done' ? '✓' : i + 1 }}</span>
        <span class="step__label">
          {{ s.label }}
          <span class="sr-only"> ({{ t(`build.${s.state}` as MsgKey) }})</span>
          <span v-if="s.state === 'current' && detail" class="step__detail">{{ detail }}</span>
        </span>
      </li>
    </ol>
  </section>
</template>

<style scoped>
.progress__title { margin: 0 0 16px; font-size: 1.25rem; }
.progress__bar { height: 8px; background: var(--border); border-radius: 4px; overflow: hidden; margin-bottom: 20px; }
.progress__fill { height: 100%; background: var(--accent); transition: width 0.3s ease; }
.progress__steps { list-style: none; margin: 0; padding: 0; display: flex; flex-direction: column; gap: 10px; }
.step { display: flex; gap: 12px; align-items: flex-start; color: var(--muted); }
.step__mark { flex: none; display: grid; place-items: center; width: 26px; height: 26px; border-radius: 50%; border: 2px solid var(--border-strong); font-size: 0.8125rem; font-weight: 700; }
.step__label { display: flex; flex-direction: column; padding-top: 2px; }
.step__detail { font-size: 0.875rem; color: var(--muted); }
.step--done { color: var(--text); }
.step--done .step__mark { background: var(--ok-text); border-color: var(--ok-text); color: var(--surface); }
.step--current { color: var(--text); font-weight: 600; }
.step--current .step__mark { border-color: var(--accent); background: var(--accent-soft); color: var(--text); }
@media (prefers-reduced-motion: no-preference) {
  .step--current .step__mark { animation: pulse 1.2s ease-in-out infinite; }
}
@keyframes pulse { 50% { box-shadow: 0 0 0 5px var(--accent-soft); } }
</style>
