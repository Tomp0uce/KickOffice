import express from 'express'
import fs from 'fs'
import path from 'path'
import { fileURLToPath } from 'url'
import { ErrorCodes } from '../config/errorCodes.js'
import { logAndRespond } from '../utils/http.js'

const __filename = fileURLToPath(import.meta.url)
const __dirname = path.dirname(__filename)

export const logsRouter = express.Router()

const FRONTEND_LOGS_DIR = path.join(__dirname, '../../logs/frontend')

if (!fs.existsSync(FRONTEND_LOGS_DIR)) {
  fs.mkdirSync(FRONTEND_LOGS_DIR, { recursive: true })
}

const ALLOWED_LEVELS = new Set(['warn', 'error', 'fatal'])
const MAX_ENTRIES = 200

logsRouter.post('/', express.json({ limit: '10mb' }), async (req, res) => {
  try {
    const { entries } = req.body

    if (!Array.isArray(entries) || entries.length === 0) {
      return logAndRespond(res, 400, { code: ErrorCodes.LOGS_INVALID_ENTRIES, error: 'entries must be a non-empty array' }, 'POST /api/logs')
    }

    if (entries.length > MAX_ENTRIES) {
      return logAndRespond(res, 400, { code: ErrorCodes.LOGS_TOO_MANY_ENTRIES, error: `entries array exceeds maximum of ${MAX_ENTRIES} items` }, 'POST /api/logs')
    }

    // Validate required fields and filter by level
    const filtered = []
    for (const entry of entries) {
      if (!entry.timestamp || !entry.level || !entry.message || !entry.source) {
        continue
      }
      if (!ALLOWED_LEVELS.has(entry.level)) {
        continue
      }
      filtered.push(entry)
    }

    if (filtered.length > 0) {
      const timestamp = new Date().getTime()
      const filename = `frontend_logs_${timestamp}.json`
      const filePath = path.join(FRONTEND_LOGS_DIR, filename)

      await fs.promises.writeFile(filePath, JSON.stringify(filtered, null, 2))
      req.logger.info('Frontend logs saved', { traffic: 'system', saved: filtered.length })
    }

    res.json({ success: true, saved: filtered.length })
  } catch (error) {
    req.logger.error('Error handling frontend logs', { error })
    logAndRespond(res, 500, { code: ErrorCodes.INTERNAL_ERROR, error: 'Internal server error processing logs' }, 'POST /api/logs')
  }
})
