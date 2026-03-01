import fs from 'fs'
import path from 'path'
import { fileURLToPath } from 'url'
import { createStream } from 'rotating-file-stream'

const __filename = fileURLToPath(import.meta.url)
const __dirname = path.dirname(__filename)

// Ensure logs directory exists at the root of the backend folder
const LOGS_DIR = path.join(__dirname, '../../logs')
if (!fs.existsSync(LOGS_DIR)) {
  fs.mkdirSync(LOGS_DIR, { recursive: true })
}

// Create a rotating write stream
const accessLogStream = createStream('kickoffice.log', {
  size: '10M', // rotate every 10 MegaBytes written
  interval: '1d', // rotate daily
  compress: 'gzip', // compress rotated files
  path: LOGS_DIR,
  maxFiles: 30 // Keep up to 30 rotated log files
})

function redactData(data) {
  // Disabled redaction to allow full debugging of prompts and responses
  return data
}

/**
 * Custom file logger to capture detailed payloads and responses.
 * @param {string} level - Log level (e.g. INFO, ERROR, DEBUG)
 * @param {string} message - Descriptive message
 * @param {any} [data] - Optional JSON payload to stringify
 */
export function systemLog(level, message, data = null) {
  const timestamp = new Date().toISOString()
  
  let dataStr = ''
  if (data !== null) {
    try {
      if (data instanceof Error) {
        dataStr = '\n' + (data.stack || data.message || String(data))
      } else {
        const safeData = redactData(data);
        dataStr = '\n' + JSON.stringify(safeData, null, 2)
      }
    } catch (err) {
      dataStr = '\n[Unserializable Data]'
    }
  }

  const logEntry = `[${timestamp}] [${level.toUpperCase()}] ${message}${dataStr}\n`
  
  // Output to console (using process.stdout/stderr to bypass console interceptors if any)
  if (level.toUpperCase() === 'ERROR') {
    process.stderr.write(logEntry)
  } else {
    process.stdout.write(logEntry)
  }

  // Append to log rotating file stream
  accessLogStream.write(logEntry)
}
