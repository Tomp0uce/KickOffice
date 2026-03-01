export type OfficeHostType = 'Word' | 'Excel' | 'PowerPoint' | 'Outlook' | 'Unknown'

let detectedHost: OfficeHostType = 'Unknown'
// Cache is only trusted after Office.onReady has fired to avoid stale detection
let officeReady = false

/** Called from main.ts inside Office.onReady callback to enable result caching. */
export function markOfficeReady(): void {
  officeReady = true
  detectedHost = 'Unknown' // reset so next call re-detects with full Office context
}

export function detectOfficeHost(): OfficeHostType {
  // Return cached value only once Office is confirmed ready
  if (officeReady && detectedHost !== 'Unknown') return detectedHost

  let result: OfficeHostType = 'Unknown'

  try {
    // Office.context.host is the official way
    const host = (window as any).Office?.context?.host
    if (host === 'Word' || host === 'Document') {
      result = 'Word'
    } else if (host === 'Excel' || host === 'Workbook') {
      result = 'Excel'
    } else if (host === 'PowerPoint' || host === 'Presentation') {
      result = 'PowerPoint'
    } else if (host === 'Outlook' || host === 'Mailbox') {
      result = 'Outlook'
    }
  } catch {
    // Fallback: check global objects
  }

  if (result === 'Unknown') {
    if (typeof (window as any).Word !== 'undefined') {
      result = 'Word'
    } else if (typeof (window as any).Excel !== 'undefined') {
      result = 'Excel'
    } else if (typeof (window as any).PowerPoint !== 'undefined') {
      result = 'PowerPoint'
    } else if (typeof (window as any).Office?.context?.mailbox !== 'undefined') {
      result = 'Outlook'
    }
  }

  // Only persist to cache after Office is ready
  if (officeReady) {
    detectedHost = result
  }

  return result
}

export function isExcel(): boolean {
  return detectOfficeHost() === 'Excel'
}

export function isWord(): boolean {
  return detectOfficeHost() === 'Word'
}

export function isPowerPoint(): boolean {
  return detectOfficeHost() === 'PowerPoint'
}

export function isOutlook(): boolean {
  return detectOfficeHost() === 'Outlook'
}

export function forHost<T>(options: {
  word?: T
  excel?: T
  powerpoint?: T
  outlook?: T
  default?: T
}): T | undefined {
  const host = detectOfficeHost()
  switch (host) {
    case 'Word':
      return options.word !== undefined ? options.word : options.default
    case 'Excel':
      return options.excel !== undefined ? options.excel : options.default
    case 'PowerPoint':
      return options.powerpoint !== undefined ? options.powerpoint : options.default
    case 'Outlook':
      return options.outlook !== undefined ? options.outlook : options.default
    default:
      return options.default
  }
}
