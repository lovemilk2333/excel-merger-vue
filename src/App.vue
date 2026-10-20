<script setup lang="ts">
import { onMounted } from 'vue'

const SPLASH_MIN_MS = 400
const SPLASH_FADE_MS = 350

function hideSplash() {
  const el = document.getElementById('app-splash')
  if (!el) return

  el.classList.add('app-splash--hide')
  el.setAttribute('aria-busy', 'false')

  window.setTimeout(() => {
    el.remove()
  }, SPLASH_FADE_MS)
}

onMounted(() => {
  const started = performance.now()
  requestAnimationFrame(() => {
    const elapsed = performance.now() - started
    const wait = Math.max(0, SPLASH_MIN_MS - elapsed)
    window.setTimeout(hideSplash, wait)
  })
})
</script>

<template>
  <RouterView />
</template>
