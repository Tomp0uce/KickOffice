/**
 * Convenience function to apply a word-level diff.
 * Wraps OfficeWordDiff for simple one-off usage.
 *
 * @param {Word.RequestContext} context - The Word request context
 * @param {Word.Range} range - The target range to update
 * @param {string} originalText - The original text of the range
 * @param {string} newText - The new text to apply
 * @param {OfficeWordDiffOptions} [options={}] - Configuration options
 * @returns {Promise<DiffResult>} Result object with operation details
 *
 * @example
 * import { applyWordDiff } from 'office-word-diff';
 *
 * await Word.run(async (context) => {
 *   const range = context.document.getSelection();
 *   range.load('text');
 *   await context.sync();
 *
 *   const result = await applyWordDiff(context, range, range.text, newText);
 * });
 */
export function applyWordDiff(context: Word.RequestContext, range: Word.Range, originalText: string, newText: string, options?: OfficeWordDiffOptions): Promise<DiffResult>;
/**
 * Compute a word-level diff between two strings.
 * Convenience function that doesn't require instantiating OfficeWordDiff.
 *
 * @param {string} text1 - Original text
 * @param {string} text2 - New text
 * @returns {Array<[number, string]>} Array of diff operations
 */
export function computeDiff(text1: string, text2: string): Array<[number, string]>;
/**
 * Get diff statistics between two strings.
 * Convenience function that doesn't require instantiating OfficeWordDiff.
 *
 * @param {string} text1 - Original text
 * @param {string} text2 - New text
 * @returns {DiffStats} Statistics about the diff
 */
export function getDiffStats(text1: string, text2: string): DiffStats;
/**
 * @typedef {Object} DiffResult
 * @property {boolean} success - Whether the operation completed successfully
 * @property {('token'|'sentence'|'block')} strategyUsed - Which strategy was used
 * @property {number} insertions - Number of insertions applied
 * @property {number} deletions - Number of deletions applied
 * @property {number} duration - Time taken in milliseconds
 * @property {Array<{timestamp: number, level: string, message: string}>} logs - Operation logs
 */
/**
 * @typedef {Object} DiffStats
 * @property {number} insertions - Number of insertions in the diff
 * @property {number} deletions - Number of deletions in the diff
 * @property {number} unchanged - Number of unchanged segments
 * @property {number} totalChanges - Total number of changes
 */
/**
 * @typedef {Object} OfficeWordDiffOptions
 * @property {boolean} [enableTracking=true] - Enable Word track changes
 * @property {('silent'|'error'|'warn'|'info'|'debug')} [logLevel='info'] - Log level
 * @property {function|null} [onLog=null] - Custom log handler: (message, level) => void
 */
/**
 * Main class for applying word-level diffs to Microsoft Word documents.
 *
 * Uses a cascading fallback strategy:
 * 1. Token Map Strategy (word-level precision)
 * 2. Sentence Diff Strategy (sentence-level)
 * 3. Block Replace Strategy (full replacement)
 *
 * @example
 * import { OfficeWordDiff } from 'office-word-diff';
 *
 * await Word.run(async (context) => {
 *   const range = context.document.getSelection();
 *   range.load('text');
 *   await context.sync();
 *
 *   const differ = new OfficeWordDiff({ enableTracking: true });
 *   const result = await differ.applyDiff(context, range, range.text, newText);
 *   console.log(`Applied ${result.insertions} insertions, ${result.deletions} deletions`);
 * });
 */
export class OfficeWordDiff {
    /**
     * Creates a new OfficeWordDiff instance
     * @param {OfficeWordDiffOptions} [options={}] - Configuration options
     */
    constructor(options?: OfficeWordDiffOptions);
    enableTracking: boolean;
    logLevel: "silent" | "error" | "warn" | "info" | "debug";
    onLog: Function;
    _logger: any;
    _dmp: any;
    /**
     * Apply a diff to a Word range with cascading fallback.
     *
     * Strategy cascade:
     * 1. Token Map - Attempts word-level mapping for precise changes
     * 2. Sentence Diff - Falls back to sentence-level if token mapping fails
     * 3. Block Replace - Final fallback, replaces entire range
     *
     * @param {Word.RequestContext} context - The Word request context (from Word.run)
     * @param {Word.Range} range - The target range to update
     * @param {string} originalText - The original text of the range
     * @param {string} newText - The new text to apply
     * @returns {Promise<DiffResult>} Result object with operation details
     */
    applyDiff(context: Word.RequestContext, range: Word.Range, originalText: string, newText: string): Promise<DiffResult>;
    /**
     * Compute a word-level diff between two strings.
     * Does NOT require Office.js context - can be used for preview.
     *
     * @param {string} text1 - Original text
     * @param {string} text2 - New text
     * @returns {Array<[number, string]>} Array of diff operations where:
     *   - [0, text] = unchanged
     *   - [-1, text] = deletion
     *   - [1, text] = insertion
     *
     * @example
     * const diffs = differ.computeDiff('Hello world', 'Hello there');
     * // Returns: [[0, 'Hello '], [-1, 'world'], [1, 'there']]
     */
    computeDiff(text1: string, text2: string): Array<[number, string]>;
    /**
     * Get statistics about the diff between two strings.
     * Does NOT require Office.js context - can be used for preview.
     *
     * @param {string} text1 - Original text
     * @param {string} text2 - New text
     * @returns {DiffStats} Statistics about the diff
     *
     * @example
     * const stats = differ.getDiffStats('Hello world', 'Hello there');
     * console.log(`${stats.insertions} insertions, ${stats.deletions} deletions`);
     */
    getDiffStats(text1: string, text2: string): DiffStats;
    /**
     * Get all logs from the last operation
     * @returns {Array<{timestamp: number, level: string, message: string}>}
     */
    getLogs(): Array<{
        timestamp: number;
        level: string;
        message: string;
    }>;
    /**
     * Clear stored logs
     */
    clearLogs(): void;
    /**
     * Set the log level
     * @param {('silent'|'error'|'warn'|'info'|'debug')} level
     */
    setLogLevel(level: ("silent" | "error" | "warn" | "info" | "debug")): void;
}
export { applyTokenMapStrategy } from "./strategies/tokenMap.js";
export { applySentenceDiffStrategy } from "./strategies/sentenceDiff.js";
export { applyBlockReplaceStrategy } from "./strategies/blockReplace.js";
export default OfficeWordDiff;
export type DiffResult = {
    /**
     * - Whether the operation completed successfully
     */
    success: boolean;
    /**
     * - Which strategy was used
     */
    strategyUsed: ("token" | "sentence" | "block");
    /**
     * - Number of insertions applied
     */
    insertions: number;
    /**
     * - Number of deletions applied
     */
    deletions: number;
    /**
     * - Time taken in milliseconds
     */
    duration: number;
    /**
     * - Operation logs
     */
    logs: Array<{
        timestamp: number;
        level: string;
        message: string;
    }>;
};
export type DiffStats = {
    /**
     * - Number of insertions in the diff
     */
    insertions: number;
    /**
     * - Number of deletions in the diff
     */
    deletions: number;
    /**
     * - Number of unchanged segments
     */
    unchanged: number;
    /**
     * - Total number of changes
     */
    totalChanges: number;
};
export type OfficeWordDiffOptions = {
    /**
     * - Enable Word track changes
     */
    enableTracking?: boolean;
    /**
     * - Log level
     */
    logLevel?: ("silent" | "error" | "warn" | "info" | "debug");
    /**
     * - Custom log handler: (message, level) => void
     */
    onLog?: Function | null;
};
export { createLogger, createLogCallback } from "./utils/logger.js";
//# sourceMappingURL=index.d.ts.map