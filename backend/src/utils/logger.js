import fs from 'fs'
import path from 'path'
import { fileURLToPath } from 'url'

const __filename = fileURLToPath(import.meta.url)
const __dirname = path.dirname(__filename)

// Ensure logs directory exists at the root of the backend folder
const LOGS_DIR = path.join(__dirname, '../../logs')
if (!fs.existsSync(LOGS_DIR)) {
  fs.mkdirSync(LOGS_DIR, { recursive: true })
}

const LOG_FILE = path.join(LOGS_DIR, 'kickoffice.log')

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
      dataStr = '\n' + JSON.stringify(data, null, 2)
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

  // Append to log file
  fs.appendFile(LOG_FILE, logEntry, (err) => {
    if (err) console.error('Failed to write to system log file:', err)
  })
}
