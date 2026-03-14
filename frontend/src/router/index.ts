import { createMemoryHistory, createRouter } from 'vue-router';

const router = createRouter({
  // Office Add-ins run inside iframes with URLs controlled by Office.
  // createMemoryHistory avoids manipulating the browser URL entirely,
  // which prevents redirect loops caused by catch-all routes or
  // history.pushState conflicts with the Office iframe host.
  history: createMemoryHistory(),
  routes: [
    {
      path: '/',
      name: 'Home',
      component: () => import('../pages/HomePage.vue'),
    },
    {
      path: '/settings',
      name: 'Settings',
      component: () => import('../pages/SettingsPage.vue'),
    },
    {
      path: '/:pathMatch(.*)*',
      redirect: '/',
    },
  ],
});

router.onError((error, to) => {
  const isChunkError =
    error.message.includes('Failed to fetch dynamically imported module') ||
    error.message.includes('Importing a module script failed');

  if (isChunkError) {
    // Prevent infinite reload loops
    if (to?.query?._refresh) {
      console.error('Failed to load chunk even after refresh.', error);
      return;
    }
    const targetPath = to?.fullPath || '/';
    const separator = targetPath.includes('?') ? '&' : '?';
    window.location.href = `${targetPath}${separator}_refresh=${Date.now()}`;
  }
});

export default router;
