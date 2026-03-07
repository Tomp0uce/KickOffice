import { Router } from 'express'

import { models } from '../config/models.js'
import { ErrorCodes } from '../config/errorCodes.js'
import { validateImagePayload } from '../middleware/validate.js'
import { handleErrorResponse, imageGeneration } from '../services/llmClient.js'
import { logAndRespond } from '../utils/http.js'
import logger from '../utils/logger.js'

const imageRouter = Router()

imageRouter.post('/', async (req, res) => {
  req.logger.debug(` /api/image incoming request`)
  const parsedPayload = validateImagePayload(req.body)
  if (parsedPayload.error) {
    return logAndRespond(res, 400, { code: ErrorCodes.VALIDATION_ERROR, error: parsedPayload.error }, 'POST /api/image')
  }

  const imageModel = models.image

  try {
    const response = await imageGeneration({
      body: {
        model: imageModel.id,
        ...parsedPayload.value,
      },
      userCredentials: req.userCredentials,
    })

    req.logger.debug(` /api/image upstream payload`, { model: imageModel.id, promptLength: parsedPayload.value.prompt.length })

    if (!response.ok) {
      await handleErrorResponse(response, '/api/image')
      return logAndRespond(res, 502, {
        code: ErrorCodes.LLM_UPSTREAM_ERROR,
        error: 'The AI service returned an error. Please try again later.',
      }, 'POST /api/image')
    }

    const data = await response.json()
    res.json(data)
  } catch (error) {
    if (error.name === 'AbortError') {
      return logAndRespond(res, 504, { code: ErrorCodes.IMAGE_TIMEOUT, error: 'Image API request timeout' }, 'POST /api/image')
    }
    req.logger.error('Image proxy error', { error })
    return logAndRespond(res, 500, { code: ErrorCodes.INTERNAL_ERROR, error: 'Internal server error' }, 'POST /api/image')
  }
})

export {
  imageRouter,
}
