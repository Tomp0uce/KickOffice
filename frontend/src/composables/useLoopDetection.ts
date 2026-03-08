export function useLoopDetection(windowSize: number = 5, repeatThreshold: number = 2) {
  const recentSignatures: string[] = []

  /**
   * Adds the tool call signature (toolName + serialized args) to a sliding window
   * and returns true if the same signature has been seen repeatThreshold or more times.
   *
   * This correctly handles both cases:
   * - Same tool, same args (e.g. getAllSlidesOverview{} × 2) → triggers
   * - Same tool, different args (e.g. setCellRange(A1) then setCellRange(B1)) → does NOT trigger
   */
  function addSignatureAndCheckLoop(signature: string): boolean {
    if (!signature) return false
    recentSignatures.push(signature)
    if (recentSignatures.length > windowSize) {
      recentSignatures.shift()
    }
    const sigCount = recentSignatures.filter(s => s === signature).length
    return sigCount >= repeatThreshold
  }

  function clearSignatures() {
    recentSignatures.length = 0
  }

  return {
    addSignatureAndCheckLoop,
    clearSignatures
  }
}
