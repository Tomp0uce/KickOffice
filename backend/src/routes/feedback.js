import express from 'express'
import fs from 'fs'
import path from 'path'
import { fileURLToPath } from 'url'

const __filename = fileURLToPath(import.meta.url)
const __dirname = path.dirname(__filename)

export const feedbackRouter = express.Router()

const FEEDBACK_DIR = path.join(__dirname, '../../logs/feedback')

if (!fs.existsSync(FEEDBACK_DIR)) {
  fs.mkdirSync(FEEDBACK_DIR, { recursive: true })
}

feedbackRouter.post('/:sessionId', express.json({ limit: '20mb' }), async (req, res) => {
  try {
    const { sessionId } = req.params
    const { comment, category, logs, chatHistory, systemContext } = req.body

    if (!comment || !category) {
      return res.status(400).json({ error: 'Comment and category are required' })
    }

    req.logger.info('Feedback received from user', { traffic: 'system', category })

    const feedbackEntry = {
      timestamp: new Date().toISOString(),
      sessionId,
      userId: req.logger.defaultMeta?.userId || 'anonymous',
      host: req.logger.defaultMeta?.host || 'unknown',
      category,
      comment,
      systemContext: systemContext || null,
      logs: logs || [],
      chatHistory: chatHistory || [],
    }

    const filename = `feedback_${category}_${new Date().getTime()}.json`
    const filePath = path.join(FEEDBACK_DIR, filename)

    await fs.promises.writeFile(filePath, JSON.stringify(feedbackEntry, null, 2))

    req.logger.info('Feedback saved', {
      traffic: 'system',
      category,
      logCount: feedbackEntry.logs.length,
      chatMessageCount: feedbackEntry.chatHistory.length,
      hasSystemContext: !!systemContext,
    })

    res.json({ success: true, message: 'Feedback submitted successfully' })
  } catch (error) {
    req.logger.error('Error handling feedback', { error })
    res.status(500).json({ error: 'Internal server error processing feedback' })
  }
})
