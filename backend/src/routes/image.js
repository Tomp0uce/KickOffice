import { Router } from 'express'

import { models } from '../config/models.js'
import { validateImagePayload } from '../middleware/validate.js'
import { handleErrorResponse, imageGeneration } from '../services/llmClient.js'
import { logAndRespond } from '../utils/http.js'

const imageRouter = Router()

imageRouter.post('/', async (req, res) => {
  const parsedPayload = validateImagePayload(req.body)
  if (parsedPayload.error) {
    return logAndRespond(res, 400, { error: parsedPayload.error }, 'POST /api/image')
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

    if (!response.ok) {
      await handleErrorResponse(response, '/api/image')
      return logAndRespond(res, 502, {
        error: 'The AI service returned an error. Please try again later.',
      }, 'POST /api/image')
    }

    const data = await response.json()
    res.json(data)
  } catch (error) {
    if (error.name === 'AbortError') {
      return logAndRespond(res, 504, { error: 'Image API request timeout' }, 'POST /api/image')
    }
    console.error('Image proxy error:', error)
    return logAndRespond(res, 500, { error: 'Internal server error' }, 'POST /api/image')
  }
})

export {
  imageRouter,
}
