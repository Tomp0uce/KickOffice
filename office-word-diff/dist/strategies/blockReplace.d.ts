/**
 * Block Replace Strategy - Last resort fallback
 *
 * This strategy deletes the entire range content and inserts new text.
 * Used when token and sentence strategies fail.
 *
 * @module office-word-diff/strategies/blockReplace
 */
/**
 * Applies the "Block Replace" strategy.
 * Deletes the entire range and inserts new text as tracked changes.
 * This is the final fallback when more granular strategies fail.
 *
 * @param {Word.RequestContext} context - The Word request context
 * @param {Word.Range} range - The target range to update
 * @param {string} newText - The new text to apply
 * @param {function} log - Callback for logging messages
 * @returns {Promise<{strategy: string, insertions: number, deletions: number}>}
 */
export function applyBlockReplaceStrategy(context: Word.RequestContext, range: Word.Range, newText: string, log: Function): Promise<{
    strategy: string;
    insertions: number;
    deletions: number;
}>;
//# sourceMappingURL=blockReplace.d.ts.map