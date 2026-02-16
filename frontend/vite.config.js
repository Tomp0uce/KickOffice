import { fileURLToPath, URL } from 'node:url'
import { execSync } from 'node:child_process'
import frontendPackage from './package.json'

import tailwindcss from '@tailwindcss/vite'
import vue from '@vitejs/plugin-vue'
import { defineConfig } from 'vite'

function getBuildVersion() {
  const packageVersion = frontendPackage.version
  const explicitBuildVersion = process.env.VITE_APP_VERSION?.trim()

  if (explicitBuildVersion) {
    return explicitBuildVersion
  }

  const commitFromCi = process.env.GITHUB_SHA
  if (commitFromCi) {
    return `${packageVersion}+${commitFromCi.slice(0, 7)}`
  }

  try {
    const shortCommit = execSync('git rev-parse --short HEAD', { stdio: ['ignore', 'pipe', 'ignore'] })
      .toString()
      .trim()
    return `${packageVersion}+${shortCommit}`
  } catch {
    return packageVersion
  }
}

const appVersion = getBuildVersion()

export default defineConfig({
  plugins: [tailwindcss(), vue()],
  define: {
    __APP_VERSION__: JSON.stringify(appVersion),
  },
  resolve: {
    alias: {
      '@': fileURLToPath(new URL('./src', import.meta.url)),
    },
  },
  server: {
    port: 3002,
    host: '0.0.0.0',
  },
})
