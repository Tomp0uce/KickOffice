import { defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    // Provide stub env vars so modules that validate them at import time
    // (e.g. config/models.js checking LLM_API_KEY) don't log warnings or
    // throw in CI where no .env file is present.
    env: {
      NODE_ENV: 'test',
      LLM_API_KEY: 'test-key-vitest',
    },
  },
});
