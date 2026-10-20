import { defineConfig } from 'vite'
import vue from '@vitejs/plugin-vue'
import legacy from '@vitejs/plugin-legacy'
import tailwindcss from '@tailwindcss/vite'

// https://vite.dev/config/
export default defineConfig({
  base: './',
  plugins: [
    vue(),
    legacy({
      targets: ['chrome >= 96', 'defaults'],
      modernPolyfills: true,
      renderLegacyChunks: true,
    }),
    tailwindcss(),
  ],
  build: {
    target: 'chrome96',
    cssTarget: 'chrome96',
  },
})
