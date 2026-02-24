import './index.css'

import { useDebounceFn, useStorage } from '@vueuse/core'
import { createApp, watch } from 'vue'

import App from './App.vue'
import { i18n } from './i18n'
import router from './router'
import { localStorageKey } from './utils/enum'
import { detectOfficeHost } from './utils/hostDetection'

window.Office.onReady(() => {
  detectOfficeHost()
  const app = createApp(App)
  const _ResizeObserver = window.ResizeObserver
  window.ResizeObserver = class ResizeObserver extends _ResizeObserver {
    constructor(callback: ResizeObserverCallback) {
      super(useDebounceFn(callback, 16))
    }
  }

  const darkMode = useStorage(localStorageKey.darkMode, false)
  watch(darkMode, (value) => {
    document.documentElement.classList.toggle('dark', value)
  }, { immediate: true })

  app.config.errorHandler = (err, instance, info) => {
    console.error('Vue Global Error:', err, instance, info)
  }

  app.use(i18n)
  app.use(router)
  app.mount('#app')
})
