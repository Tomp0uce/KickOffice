export type OfficeHostType = 'Word' | 'Excel' | 'PowerPoint' | 'Outlook' | 'Unknown'

let detectedHost: OfficeHostType = 'Unknown'

export function detectOfficeHost(): OfficeHostType {
  if (detectedHost !== 'Unknown') return detectedHost

  try {
    // Office.context.host is the official way
    const host = (window as any).Office?.context?.host
    if (host === 'Word' || host === 'Document') {
      detectedHost = 'Word'
    } else if (host === 'Excel' || host === 'Workbook') {
      detectedHost = 'Excel'
    } else if (host === 'PowerPoint' || host === 'Presentation') {
      detectedHost = 'PowerPoint'
    } else if (host === 'Outlook' || host === 'Mailbox') {
      detectedHost = 'Outlook'
    }
  } catch {
    // Fallback: check global objects
  }

  if (detectedHost === 'Unknown') {
    if (typeof (window as any).Word !== 'undefined') {
      detectedHost = 'Word'
    } else if (typeof (window as any).Excel !== 'undefined') {
      detectedHost = 'Excel'
    } else if (typeof (window as any).PowerPoint !== 'undefined') {
      detectedHost = 'PowerPoint'
    } else if (typeof (window as any).Office?.context?.mailbox !== 'undefined') {
      detectedHost = 'Outlook'
    }
  }

  return detectedHost
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
