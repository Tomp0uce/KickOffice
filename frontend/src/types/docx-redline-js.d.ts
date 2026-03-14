/**
 * Type declarations for @ansonlai/docx-redline-js
 * https://github.com/AnsonLai/docx-redline-js
 */

declare module '@ansonlai/docx-redline-js' {
  export interface RedlineOptions {
    author?: string
    generateRedlines?: boolean
  }

  export interface RedlineResult {
    oxml: string
  }

  /**
   * Apply redline (track changes) markup to OOXML paragraph content.
   * Generates native Word w:ins and w:del revision markup.
   *
   * @param paragraphOoxml - OOXML string of the paragraph to modify
   * @param originalText - Original text content
   * @param revisedText - Modified text content
   * @param options - Optional configuration for redline generation
   * @returns Promise resolving to RedlineResult with modified OOXML
   */
  export function applyRedlineToOxml(
    paragraphOoxml: string,
    originalText: string,
    revisedText: string,
    options?: RedlineOptions
  ): Promise<RedlineResult>

  /**
   * Set the default author name for track changes revisions.
   *
   * @param authorName - Author name to use for revisions
   */
  export function setDefaultAuthor(authorName: string): void
}
