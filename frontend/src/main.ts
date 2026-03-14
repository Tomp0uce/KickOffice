import './index.css';

import { createApp, watch, ref } from 'vue';

import App from './App.vue';
import { i18n } from './i18n';
import router from './router';
import { localStorageKey } from './utils/enum';
import { detectOfficeHost, markOfficeReady } from './utils/hostDetection';
import { setRememberCredentials, migrateCredentialsOnStartup } from './utils/credentialStorage';
import { logService } from './utils/logger';

// Monkey patch console.error and console.warn
console.error = (...args) => {
  const msg = typeof args[0] === 'string' ? args[0] : 'Console Error';
  const err = args.find(a => a instanceof Error);
  logService.error(msg, err, args);
};
console.warn = (...args) => {
  const msg = typeof args[0] === 'string' ? args[0] : 'Console Warn';
  logService.warn(msg, args);
};

window.onerror = (message, _source, _lineno, _colno, error) => {
  logService.error(`Global Error: ${message}`, error);
};
window.onunhandledrejection = event => {
  logService.error('Unhandled Promise Rejection', event.reason);
};

window.Office.onReady(async () => {
  markOfficeReady();
  detectOfficeHost();

  // BUGFIX: Initialize rememberCredentials carefully to avoid breaking existing credentials
  // Office Add-ins MUST persist credentials in localStorage (sessionStorage is wiped on restart)
  if (localStorage.getItem('rememberCredentials') === null) {
    // Check if user has existing credentials in sessionStorage before forcing localStorage
    const hasSessionCreds =
      sessionStorage.getItem('litellmUserKey') || sessionStorage.getItem('ko_cred_litellmUserKey');

    if (hasSessionCreds) {
      console.info(
        '[KickOffice] Existing credentials found in sessionStorage, keeping rememberCredentials=false for now',
      );
      // User will need to explicitly enable "remember credentials" to persist them
    } else {
      console.info('[KickOffice] First launch detected — enabling credential persistence');
      await setRememberCredentials(true);
    }
  }

  // ARCH-M3: One-time credential migration at startup
  await migrateCredentialsOnStartup();

  const app = createApp(App);

  // Use raw localStorage for dark mode
  const initialDarkMode = localStorage.getItem(localStorageKey.darkMode) === 'true';
  const darkMode = ref(initialDarkMode);

  window.addEventListener('storage', e => {
    if (e.key === localStorageKey.darkMode) {
      darkMode.value = e.newValue === 'true';
    }
  });

  watch(
    darkMode,
    value => {
      document.documentElement.classList.toggle('dark', value);
      localStorage.setItem(localStorageKey.darkMode, String(value));
    },
    { immediate: true },
  );

  app.config.errorHandler = (err, _instance, info) => {
    logService.error(`Vue Global Error: ${info}`, err);
  };

  app.use(i18n);
  app.use(router);
  app.mount('#app');
});
