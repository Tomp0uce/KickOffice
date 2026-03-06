import { Router } from 'express'
import { createRequire } from 'module'

const require = createRequire(import.meta.url)
const { version } = require('../../package.json')

const healthRouter = Router()

healthRouter.get('/health', (_req, res) => {
  const hasKey = !!process.env.LLM_API_KEY
  res.json({
    status: 'ok',
    timestamp: new Date().toISOString(),
    version,
    llmConfigured: hasKey
  })
})

export {
  healthRouter,
}
