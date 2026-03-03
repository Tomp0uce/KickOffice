/**
 * Creates a logger instance with configurable level and output handler
 *
 * @param {('silent'|'error'|'warn'|'info'|'debug')} [level='info'] - Minimum log level to capture
 * @param {function|null} [onLog=null] - Custom log handler: (message, level) => void
 * @returns {object} Logger instance
 */
export function createLogger(level?: ("silent" | "error" | "warn" | "info" | "debug"), onLog?: Function | null): object;
/**
 * Creates a simple log function compatible with the strategy callbacks
 * @param {object} logger - Logger instance from createLogger()
 * @returns {function} Log callback function
 */
export function createLogCallback(logger: object): Function;
//# sourceMappingURL=logger.d.ts.map