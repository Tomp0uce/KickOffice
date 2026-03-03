/**
 * office-word-diff
 * 
 * Apply word-level text diffs to Microsoft Word documents using the Office.js API,
 * preserving formatting and enabling granular track changes.
 * 
 * @module office-word-diff
 * @license Apache-2.0
 */

import DiffMatchPatch from '../lib/diff-wordmode.js';
import { applyTokenMapStrategy } from './strategies/tokenMap.js';
import { applySentenceDiffStrategy } from './strategies/sentenceDiff.js';
import { applyBlockReplaceStrategy } from './strategies/blockReplace.js';
import { createLogger, createLogCallback } from './utils/logger.js';

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
    constructor(options = {}) {
        this.enableTracking = options.enableTracking !== false;
        this.logLevel = options.logLevel || 'info';
        this.onLog = options.onLog || null;
        
        this._logger = createLogger(this.logLevel, this.onLog);
        this._dmp = new DiffMatchPatch();
    }

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
    async applyDiff(context, range, originalText, newText) {
        const startTime = Date.now();
        this._logger.clearLogs();
        
        this._logger.info(`Starting diff operation (${originalText.length} -> ${newText.length} chars)`);

        // Quick check for identical text
        if (originalText === newText) {
            this._logger.info('Text is identical, no changes needed');
            return {
                success: true,
                strategyUsed: 'token',
                insertions: 0,
                deletions: 0,
                duration: Date.now() - startTime,
                logs: this._logger.getLogs()
            };
        }

        const logCallback = createLogCallback(this._logger);

        try {
            const result = await applyTokenMapStrategy(
                context,
                range,
                originalText,
                newText,
                logCallback
            );

            return {
                success: true,
                strategyUsed: result.strategy,
                insertions: result.insertions,
                deletions: result.deletions,
                duration: Date.now() - startTime,
                logs: this._logger.getLogs()
            };

        } catch (error) {
            this._logger.error(`All strategies failed: ${error.message}`);
            
            return {
                success: false,
                strategyUsed: 'block',
                insertions: 0,
                deletions: 0,
                duration: Date.now() - startTime,
                logs: this._logger.getLogs()
            };
        }
    }

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
    computeDiff(text1, text2) {
        return this._dmp.diff_wordMode(text1, text2);
    }

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
    getDiffStats(text1, text2) {
        const diffs = this._dmp.diff_wordMode(text1, text2);
        
        let insertions = 0;
        let deletions = 0;
        let unchanged = 0;

        for (const [op, text] of diffs) {
            const tokens = text.match(/(\w+|[^\w\s]+|\s+)/g) || [];
            const count = tokens.length;
            
            if (op === 0) {
                unchanged += count;
            } else if (op === -1) {
                deletions += count;
            } else if (op === 1) {
                insertions += count;
            }
        }

        return {
            insertions,
            deletions,
            unchanged,
            totalChanges: insertions + deletions
        };
    }

    /**
     * Get all logs from the last operation
     * @returns {Array<{timestamp: number, level: string, message: string}>}
     */
    getLogs() {
        return this._logger.getLogs();
    }

    /**
     * Clear stored logs
     */
    clearLogs() {
        this._logger.clearLogs();
    }

    /**
     * Set the log level
     * @param {('silent'|'error'|'warn'|'info'|'debug')} level
     */
    setLogLevel(level) {
        this._logger.setLevel(level);
        this.logLevel = level;
    }
}

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
export async function applyWordDiff(context, range, originalText, newText, options = {}) {
    const differ = new OfficeWordDiff(options);
    return differ.applyDiff(context, range, originalText, newText);
}

/**
 * Compute a word-level diff between two strings.
 * Convenience function that doesn't require instantiating OfficeWordDiff.
 * 
 * @param {string} text1 - Original text
 * @param {string} text2 - New text
 * @returns {Array<[number, string]>} Array of diff operations
 */
export function computeDiff(text1, text2) {
    const dmp = new DiffMatchPatch();
    return dmp.diff_wordMode(text1, text2);
}

/**
 * Get diff statistics between two strings.
 * Convenience function that doesn't require instantiating OfficeWordDiff.
 * 
 * @param {string} text1 - Original text
 * @param {string} text2 - New text
 * @returns {DiffStats} Statistics about the diff
 */
export function getDiffStats(text1, text2) {
    const differ = new OfficeWordDiff({ logLevel: 'silent' });
    return differ.getDiffStats(text1, text2);
}

// Re-export strategies for advanced usage
export { applyTokenMapStrategy } from './strategies/tokenMap.js';
export { applySentenceDiffStrategy } from './strategies/sentenceDiff.js';
export { applyBlockReplaceStrategy } from './strategies/blockReplace.js';

// Re-export logger utilities
export { createLogger, createLogCallback } from './utils/logger.js';

// Default export
export default OfficeWordDiff;
