import { Router } from 'express'

import { LLM_API_BASE_URL, LLM_API_KEY, models } from '../config/models.js'
import { validateImagePayload } from '../middleware/validate.js'
import { fetchWithTimeout, logAndRespond } from '../utils/http.js'

const imageRouter = Router()

function getImageTimeoutMs() {
  return 180_000
}

imageRouter.post('/', async (req, res) => {
  const parsedPayload = validateImagePayload(req.body)
  if (parsedPayload.error) {
    return logAndRespond(res, 400, { error: parsedPayload.error }, 'POST /api/image')
  }

  const imageModel = models.image
  if (!imageModel) {
    return logAndRespond(res, 500, { error: 'Image model not configured' }, 'POST /api/image')
  }

  try {
    const response = await fetchWithTimeout(`${LLM_API_BASE_URL}/images/generations`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${LLM_API_KEY}`,
      },
      body: JSON.stringify({
        model: imageModel.id,
        ...parsedPayload.value,
      }),
    }, getImageTimeoutMs())

    if (!response.ok) {
      const errorText = await response.text()
      console.error(`Image API error ${response.status}:`, errorText)
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
