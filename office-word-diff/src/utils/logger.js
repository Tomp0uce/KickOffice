/**
 * Logger utility for office-word-diff
 * Provides configurable log levels and log capture
 * 
 * @module office-word-diff/utils/logger
 */

const LOG_LEVELS = {
  silent: 0,
  error: 1,
  warn: 2,
  info: 3,
  debug: 4
};

/**
 * Creates a logger instance with configurable level and output handler
 * 
 * @param {('silent'|'error'|'warn'|'info'|'debug')} [level='info'] - Minimum log level to capture
 * @param {function|null} [onLog=null] - Custom log handler: (message, level) => void
 * @returns {object} Logger instance
 */
export function createLogger(level = 'info', onLog = null) {
  let currentLevel = LOG_LEVELS[level] || LOG_LEVELS.info;
  const logs = [];

  const logger = {
    /**
     * Log a message at a specified level
     * @param {string} message - Message to log
     * @param {('error'|'warn'|'info'|'debug')} [msgLevel='info'] - Log level
     */
    log(message, msgLevel = 'info') {
      const levelValue = LOG_LEVELS[msgLevel] || LOG_LEVELS.info;
      
      if (levelValue <= currentLevel) {
        const entry = {
          timestamp: Date.now(),
          level: msgLevel,
          message
        };
        
        logs.push(entry);
        
        if (onLog) {
          onLog(message, msgLevel);
        } else if (currentLevel > LOG_LEVELS.silent) {
          const prefix = `[OfficeWordDiff:${msgLevel.toUpperCase()}]`;
          if (msgLevel === 'error') {
            console.error(prefix, message);
          } else if (msgLevel === 'warn') {
            console.warn(prefix, message);
          } else if (msgLevel === 'debug') {
            console.debug(prefix, message);
          } else {
            console.log(prefix, message);
          }
        }
      }
    },

    /**
     * Log an error message
     * @param {string} message
     */
    error(message) {
      this.log(message, 'error');
    },

    /**
     * Log a warning message
     * @param {string} message
     */
    warn(message) {
      this.log(message, 'warn');
    },

    /**
     * Log an info message
     * @param {string} message
     */
    info(message) {
      this.log(message, 'info');
    },

    /**
     * Log a debug message
     * @param {string} message
     */
    debug(message) {
      this.log(message, 'debug');
    },

    /**
     * Get all captured log entries
     * @returns {Array<{timestamp: number, level: string, message: string}>}
     */
    getLogs() {
      return [...logs];
    },

    /**
     * Clear all captured logs
     */
    clearLogs() {
      logs.length = 0;
    },

    /**
     * Set the minimum log level
     * @param {('silent'|'error'|'warn'|'info'|'debug')} level
     */
    setLevel(level) {
      currentLevel = LOG_LEVELS[level] || LOG_LEVELS.info;
    },

    /**
     * Get current log level
     * @returns {string}
     */
    getLevel() {
      return Object.keys(LOG_LEVELS).find(key => LOG_LEVELS[key] === currentLevel) || 'info';
    }
  };

  return logger;
}

/**
 * Creates a simple log function compatible with the strategy callbacks
 * @param {object} logger - Logger instance from createLogger()
 * @returns {function} Log callback function
 */
export function createLogCallback(logger) {
  return (message) => {
    // Parse message to determine level
    if (message.startsWith('❌') || message.toLowerCase().includes('error')) {
      logger.error(message);
    } else if (message.startsWith('⚠️') || message.toLowerCase().includes('warn')) {
      logger.warn(message);
    } else if (message.startsWith('DEBUG:')) {
      logger.debug(message);
    } else if (message.startsWith('✅')) {
      logger.info(message);
    } else {
      logger.debug(message);
    }
  };
}
