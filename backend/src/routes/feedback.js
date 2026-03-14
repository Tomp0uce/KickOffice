import express from 'express'
import fs from 'fs'
import path from 'path'
import { fileURLToPath } from 'url'
import { ErrorCodes } from '../config/errorCodes.js'
import { logAndRespond } from '../utils/http.js'
import { getRecentRequests, getRecentToolUsage, logFeedbackSubmission } from '../utils/toolUsageLogger.js'

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
      return logAndRespond(res, 400, { code: ErrorCodes.FEEDBACK_MISSING_FIELDS, error: 'Comment and category are required' }, 'POST /api/feedback')
    }

    req.logger.info('Feedback received from user', { traffic: 'system', category })

    const userId = req.logger.defaultMeta?.userId || 'anonymous'
    const host = req.logger.defaultMeta?.host || 'unknown'

    // FB-M1: Include recent requests and tool usage
    const recentRequests = getRecentRequests(userId, 4)
    const toolUsageSnapshot = getRecentToolUsage(userId, 50)

    const feedbackEntry = {
      timestamp: new Date().toISOString(),
      sessionId,
      userId,
      host,
      category,
      comment,
      systemContext: systemContext || null,
      logs: logs || [],
      chatHistory: chatHistory || [],
      recentRequests, // FB-M1: Last 4 backend requests
      toolUsageSnapshot, // FB-M1: Recent tool usage at feedback time
    }

    const filename = `feedback_${category}_${new Date().getTime()}.json`
    const filePath = path.join(FEEDBACK_DIR, filename)

    await fs.promises.writeFile(filePath, JSON.stringify(feedbackEntry, null, 2))

    // FB-M1: Log feedback submission to index
    try {
      logFeedbackSubmission(userId, host, category, sessionId, filename)
    } catch (indexError) {
      req.logger.warn('Failed to log feedback to index', { error: indexError, traffic: 'system' })
    }

    req.logger.info('Feedback saved', {
      traffic: 'system',
      category,
      logCount: feedbackEntry.logs.length,
      chatMessageCount: feedbackEntry.chatHistory.length,
      hasSystemContext: !!systemContext,
      recentRequestsCount: feedbackEntry.recentRequests.length,
      toolUsageSnapshotCount: feedbackEntry.toolUsageSnapshot.length,
    })

    res.json({ success: true, message: 'Feedback submitted successfully' })
  } catch (error) {
    req.logger.error('Error handling feedback', { error })
    logAndRespond(res, 500, { code: ErrorCodes.INTERNAL_ERROR, error: 'Internal server error processing feedback' }, 'POST /api/feedback')
  }
})
