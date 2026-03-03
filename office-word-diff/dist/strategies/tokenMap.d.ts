/**
 * Applies the "Refined Token Map" strategy to update a range with new text.
 *
 * This strategy attempts to map words 1:1 to preserve formatting and track changes granularly.
 * If it fails (e.g., due to complex structural changes), it falls back to the Sentence Diff strategy.
 *
 * @param {Word.RequestContext} context - The Word request context
 * @param {Word.Range} range - The target range to update
 * @param {string} originalText - The original text of the range (for fallback)
 * @param {string} newText - The new text to apply
 * @param {function} log - Callback for logging messages
 * @returns {Promise<{strategy: string, insertions: number, deletions: number}>}
 * @throws {Error} If all strategies fail
 */
export function applyTokenMapStrategy(context: Word.RequestContext, range: Word.Range, originalText: string, newText: string, log: Function): Promise<{
    strategy: string;
    insertions: number;
    deletions: number;
}>;
//# sourceMappingURL=tokenMap.d.ts.map