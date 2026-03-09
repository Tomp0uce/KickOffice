/**
 * Iconify API proxy — avoids CORS issues when fetching icons from the browser.
 */
import { Router } from 'express'
import logger from '../utils/logger.js'
import { ErrorCodes } from '../config/errorCodes.js'
import { logAndRespond } from '../utils/http.js'

const router = Router()

// Search icons
router.get('/search', async (req, res) => {
  try {
    const { query, limit = '10', prefix } = req.query
    if (!query) return logAndRespond(res, 400, { code: ErrorCodes.ICON_QUERY_REQUIRED, error: 'query is required' }, 'GET /api/icons/search')

    const params = new URLSearchParams({ query: String(query), limit: String(limit) })
    if (prefix) params.set('prefix', String(prefix))

    const response = await fetch(`https://api.iconify.design/search?${params}`, {
      headers: { 'User-Agent': 'KickOffice/1.0' }
    })
    const data = await response.json()
    res.json(data)
  } catch (err) {
    logger.error('[icons] search failed', { error: err.message, traffic: 'user' })
    logAndRespond(res, 502, { code: ErrorCodes.ICON_FETCH_FAILED, error: 'Icon search failed', details: err.message }, 'GET /api/icons/search')
  }
})

// Get SVG for a specific icon
router.get('/svg/:prefix/:name', async (req, res) => {
  try {
    const { prefix, name } = req.params
    const { color } = req.query

    let url = `https://api.iconify.design/${prefix}/${name}.svg`
    if (color) url += `?color=${encodeURIComponent(String(color))}`

    const response = await fetch(url, {
      headers: { 'User-Agent': 'KickOffice/1.0' }
    })
    if (!response.ok) return logAndRespond(res, 404, { code: ErrorCodes.ICON_NOT_FOUND, error: 'Icon not found' }, 'GET /api/icons/svg')

    const svg = await response.text()
    res.type('image/svg+xml').send(svg)
  } catch (err) {
    logger.error('[icons] svg fetch failed', { error: err.message, traffic: 'user' })
    logAndRespond(res, 502, { code: ErrorCodes.ICON_FETCH_FAILED, error: 'SVG fetch failed', details: err.message }, 'GET /api/icons/svg')
  }
})

export const iconsRouter = router
