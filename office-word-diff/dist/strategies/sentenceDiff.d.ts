/**
 * Applies the "Sentence Diff" strategy.
 * Tokenizes by sentence boundaries to handle larger structural changes.
 * Falls back to block replacement if it fails.
 *
 * @param {Word.RequestContext} context - The Word request context
 * @param {Word.Range} range - The target range to update
 * @param {string} text1 - Original text
 * @param {string} text2 - New text
 * @param {function} log - Callback for logging messages
 * @returns {Promise<{strategy: string, insertions: number, deletions: number}>}
 * @throws {Error} If all strategies fail
 */
export function applySentenceDiffStrategy(context: Word.RequestContext, range: Word.Range, text1: string, text2: string, log: Function): Promise<{
    strategy: string;
    insertions: number;
    deletions: number;
}>;
//# sourceMappingURL=sentenceDiff.d.ts.map