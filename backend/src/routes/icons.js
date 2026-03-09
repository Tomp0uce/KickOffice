/**
 * Iconify API proxy — avoids CORS issues when fetching icons from the browser.
 */
import { Router } from 'express'
import logger from '../utils/logger.js'

const router = Router()

// Search icons
router.get('/search', async (req, res) => {
  try {
    const { query, limit = '10', prefix } = req.query
    if (!query) return res.status(400).json({ error: 'query is required' })

    const params = new URLSearchParams({ query: String(query), limit: String(limit) })
    if (prefix) params.set('prefix', String(prefix))

    const response = await fetch(`https://api.iconify.design/search?${params}`, {
      headers: { 'User-Agent': 'KickOffice/1.0' }
    })
    const data = await response.json()
    res.json(data)
  } catch (err) {
    logger.error('[icons] search failed', { error: err.message, traffic: 'user' })
    res.status(502).json({ error: 'Icon search failed', details: err.message })
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
    if (!response.ok) return res.status(404).json({ error: 'Icon not found' })

    const svg = await response.text()
    res.type('image/svg+xml').send(svg)
  } catch (err) {
    logger.error('[icons] svg fetch failed', { error: err.message, traffic: 'user' })
    res.status(502).json({ error: 'SVG fetch failed', details: err.message })
  }
})

export const iconsRouter = router
