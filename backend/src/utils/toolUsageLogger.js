import fs from 'fs'
import path from 'path'
import { fileURLToPath } from 'url'

const __filename = fileURLToPath(import.meta.url)
const __dirname = path.dirname(__filename)

const TOOL_USAGE_LOG = path.join(__dirname, '../../logs/tool-usage.jsonl')
const REQUEST_HISTORY_LOG = path.join(__dirname, '../../logs/request-history.jsonl')
const FEEDBACK_INDEX_LOG = path.join(__dirname, '../../logs/feedback-index.jsonl')

/**
 * LOG-H1: Log tool usage to JSONL file for analytics
 *
 * Format: {"ts":"2026-03-14T10:00:00Z","user":"john","host":"PowerPoint","tool":"screenshotSlide","count":1}
 *
 * @param {string} userId - User identifier
 * @param {string} host - Office host (Word, Excel, PowerPoint, Outlook)
 * @param {Array} toolCalls - Array of tool call objects with {name: string}
 */
export function logToolUsage(userId, host, toolCalls) {
  if (!Array.isArray(toolCalls) || toolCalls.length === 0) {
    return
  }

  const timestamp = new Date().toISOString()

  // Count tool usage by name
  const toolCounts = {}
  for (const toolCall of toolCalls) {
    const toolName = toolCall?.name || toolCall?.function?.name
    if (toolName) {
      toolCounts[toolName] = (toolCounts[toolName] || 0) + 1
    }
  }

  // Write one JSONL entry per tool type
  const entries = Object.entries(toolCounts).map(([tool, count]) =>
    JSON.stringify({
      ts: timestamp,
      user: userId || 'anonymous',
      host: host || 'unknown',
      tool,
      count,
    })
  ).join('\n')

  if (entries) {
    // Append to JSONL file (create if doesn't exist)
    fs.appendFileSync(TOOL_USAGE_LOG, entries + '\n', 'utf8')
  }
}

/**
 * LOG-H1: Read recent tool usage for a specific user
 *
 * @param {string} userId - User identifier
 * @param {number} limitLines - Maximum number of lines to return (default 100)
 * @returns {Array} Array of tool usage entries
 */
export function getRecentToolUsage(userId, limitLines = 100) {
  if (!fs.existsSync(TOOL_USAGE_LOG)) {
    return []
  }

  const content = fs.readFileSync(TOOL_USAGE_LOG, 'utf8')
  const lines = content.trim().split('\n').filter(Boolean)

  // Filter by user and return most recent entries
  const userEntries = lines
    .map(line => {
      try {
        return JSON.parse(line)
      } catch {
        return null
      }
    })
    .filter(entry => entry && entry.user === userId)
    .slice(-limitLines)

  return userEntries
}

/**
 * FB-M1: Log chat request for user history tracking
 *
 * Format: {"ts":"2026-03-14T10:00:00Z","user":"john","host":"PowerPoint","endpoint":"/api/chat","messageCount":3}
 *
 * @param {string} userId - User identifier
 * @param {string} host - Office host (Word, Excel, PowerPoint, Outlook)
 * @param {string} endpoint - Request endpoint (/api/chat or /api/chat/sync)
 * @param {number} messageCount - Number of messages in the request
 */
export function logChatRequest(userId, host, endpoint, messageCount) {
  const entry = JSON.stringify({
    ts: new Date().toISOString(),
    user: userId || 'anonymous',
    host: host || 'unknown',
    endpoint,
    messageCount: messageCount || 0,
  })

  fs.appendFileSync(REQUEST_HISTORY_LOG, entry + '\n', 'utf8')
}

/**
 * FB-M1: Get last N chat requests for a user
 *
 * @param {string} userId - User identifier
 * @param {number} limit - Number of requests to return (default 4)
 * @returns {Array} Array of request entries
 */
export function getRecentRequests(userId, limit = 4) {
  if (!fs.existsSync(REQUEST_HISTORY_LOG)) {
    return []
  }

  const content = fs.readFileSync(REQUEST_HISTORY_LOG, 'utf8')
  const lines = content.trim().split('\n').filter(Boolean)

  // Filter by user and return most recent entries
  const userEntries = lines
    .map(line => {
      try {
        return JSON.parse(line)
      } catch {
        return null
      }
    })
    .filter(entry => entry && entry.user === userId)
    .slice(-limit)

  return userEntries
}

/**
 * FB-M1: Log feedback submission to index
 *
 * Format: {"ts":"2026-03-14T10:00:00Z","user":"john","host":"PowerPoint","category":"bug","sessionId":"abc123","filename":"feedback_bug_1234567890.json"}
 *
 * @param {string} userId - User identifier
 * @param {string} host - Office host
 * @param {string} category - Feedback category
 * @param {string} sessionId - Session identifier
 * @param {string} filename - Feedback file name
 */
export function logFeedbackSubmission(userId, host, category, sessionId, filename) {
  const entry = JSON.stringify({
    ts: new Date().toISOString(),
    user: userId || 'anonymous',
    host: host || 'unknown',
    category,
    sessionId,
    filename,
  })

  fs.appendFileSync(FEEDBACK_INDEX_LOG, entry + '\n', 'utf8')
}
