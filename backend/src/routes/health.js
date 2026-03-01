import { Router } from 'express'
import { createRequire } from 'module'

const require = createRequire(import.meta.url)
const { version } = require('../../package.json')

const healthRouter = Router()

healthRouter.get('/health', (_req, res) => {
  res.json({
    status: 'ok',
    timestamp: new Date().toISOString(),
    version,
  })
})

export {
  healthRouter,
}
