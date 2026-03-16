/**
 * Factory for mutation detection in Office.js sandboxed eval tools.
 *
 * DUP-H1: Deduplicates looksLikeMutation* pattern shared by wordTools,
 * excelTools, and powerpointTools. Each tool file passes its own regex array;
 * the logic (patterns.some(p => p.test(code))) is defined once here.
 */
export function createMutationDetector(patterns: RegExp[]): (code: string) => boolean {
  return (code: string): boolean => patterns.some(p => p.test(code));
}
