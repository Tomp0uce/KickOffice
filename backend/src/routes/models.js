import { Router } from 'express'

import { getPublicModels } from '../config/models.js'

const modelsRouter = Router()

modelsRouter.get('/api/models', (_req, res) => {
  res.json(getPublicModels())
})

export {
  modelsRouter,
}
