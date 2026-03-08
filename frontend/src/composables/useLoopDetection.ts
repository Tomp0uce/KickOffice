export function useLoopDetection(windowSize: number = 5, repeatThreshold: number = 2) {
  const recentSignatures: string[] = []

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
