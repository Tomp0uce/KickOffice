/**
 * /api/files — Proxy to the upstream LLM provider's /v1/files endpoint.
 *
 * Allows the frontend to upload extracted file content to the LLM provider
 * and receive a file_id that can be referenced in subsequent chat messages,
 * avoiding re-sending large file content inline on every request.
 */

import { Router } from 'express'
import multer from 'multer'
import { LLM_API_BASE_URL, LLM_API_KEY } from '../config/models.js'
import { sanitizeErrorText } from '../utils/http.js'
import logger from '../utils/logger.js'

const filesRouter = Router()

const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 50 * 1024 * 1024, // 50MB — larger limit for files API
  },
})

/**
 * POST /api/files
 * Upload a file to the LLM provider and return its file_id.
 * Body: multipart/form-data with a `file` field and optional `purpose` field.
 */
filesRouter.post('/', upload.single('file'), async (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: 'No file provided.' })
  }

  const purpose = req.body?.purpose || 'assistants'
  const filename = req.file.originalname || 'uploaded_file'

  req.logger.info('POST /api/files upload started', {
    filename,
    size: req.file.size,
    purpose,
  })

  try {
    // Re-create a FormData to forward to the upstream LLM provider
    const formData = new FormData()
    const blob = new Blob([req.file.buffer], { type: req.file.mimetype || 'application/octet-stream' })
    formData.append('file', blob, filename)
    formData.append('purpose', purpose)

    const response = await fetch(`${LLM_API_BASE_URL}/files`, {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${LLM_API_KEY}`,
        'X-User-Key': (req.headers['x-user-key'] || '').replace(/[\r\n\x00-\x1F\x7F]/g, ''),
        'X-OpenWebUi-User-Email': (req.headers['x-openwebui-user-email'] || req.headers['x-user-email'] || '').replace(/[\r\n\x00-\x1F\x7F]/g, ''),
      },
      body: formData,
    })

    if (!response.ok) {
      const errorText = await response.text()
      const sanitized = sanitizeErrorText(errorText)
      logger.error('POST /api/files upstream error', { status: response.status, errorText: sanitized })
      return res.status(response.status).json({ error: `File upload to LLM provider failed: ${sanitized}` })
    }

    const data = await response.json()
    const fileId = data.id || data.file_id

    if (!fileId) {
      logger.error('POST /api/files upstream returned no file id', { data })
      return res.status(502).json({ error: 'LLM provider returned no file id.' })
    }

    req.logger.info('POST /api/files upload completed', { filename, fileId })
    return res.json({ fileId })
  } catch (err) {
    req.logger.error('POST /api/files error', { error: err })
    return res.status(500).json({ error: `File upload failed: ${err.message}` })
  }
})

export { filesRouter }
