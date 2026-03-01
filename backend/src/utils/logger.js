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
  if (!data) return data;
  if (typeof data !== 'object') return data;
  
  // Clone the object to avoid mutating the original
  let redacted;
  try {
    redacted = JSON.parse(JSON.stringify(data));
  } catch (e) {
    return '[Unserializable Data]';
  }

  const redactMessages = (obj) => {
    if (obj.messages && Array.isArray(obj.messages)) {
      obj.messages = obj.messages.map(msg => ({
        ...msg,
        content: '[REDACTED_FOR_PRIVACY]'
      }));
    }
    if (obj.body && obj.body.messages && Array.isArray(obj.body.messages)) {
      obj.body.messages = obj.body.messages.map(msg => ({
        ...msg,
        content: '[REDACTED_FOR_PRIVACY]'
      }));
    }
  };

  const redactChoices = (obj) => {
    if (obj.choices && Array.isArray(obj.choices)) {
      obj.choices = obj.choices.map(choice => {
        const c = { ...choice };
        if (c.message && c.message.content) {
          c.message.content = '[REDACTED_FOR_PRIVACY]';
        }
        return c;
      });
    }
  };

  redactMessages(redacted);
  redactChoices(redacted);

  return redacted;
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
      const safeData = redactData(data);
      dataStr = '\n' + JSON.stringify(safeData, null, 2)
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
