import './index.css'

import { createApp, watch, ref } from 'vue'

import App from './App.vue'
import { i18n } from './i18n'
import router from './router'
import { localStorageKey } from './utils/enum'
import { detectOfficeHost, markOfficeReady } from './utils/hostDetection'
import { setRememberCredentials, getRememberCredentials } from './utils/credentialStorage'

window.Office.onReady(async () => {
  markOfficeReady()
  detectOfficeHost()

  // BUGFIX: Initialize rememberCredentials carefully to avoid breaking existing credentials
  // Office Add-ins MUST persist credentials in localStorage (sessionStorage is wiped on restart)
  if (localStorage.getItem('rememberCredentials') === null) {
    // Check if user has existing credentials in sessionStorage before forcing localStorage
    const hasSessionCreds = sessionStorage.getItem('litellmUserKey') || sessionStorage.getItem('ko_cred_litellmUserKey')

    if (hasSessionCreds) {
      console.info('[KickOffice] Existing credentials found in sessionStorage, keeping rememberCredentials=false for now')
      // User will need to explicitly enable "remember credentials" to persist them
    } else {
      console.info('[KickOffice] First launch detected — enabling credential persistence')
      await setRememberCredentials(true)
    }
  }

  const app = createApp(App)

  // Use raw localStorage for dark mode
  const initialDarkMode = localStorage.getItem(localStorageKey.darkMode) === 'true'
  const darkMode = ref(initialDarkMode)
  
  window.addEventListener('storage', (e) => {
    if (e.key === localStorageKey.darkMode) {
      darkMode.value = e.newValue === 'true'
    }
  })

  watch(darkMode, (value) => {
    document.documentElement.classList.toggle('dark', value)
    localStorage.setItem(localStorageKey.darkMode, String(value))
  }, { immediate: true })

  app.config.errorHandler = (err, instance, info) => {
    console.error('Vue Global Error:', err, instance, info)
  }

  app.use(i18n)
  app.use(router)
  app.mount('#app')
})
