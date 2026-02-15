export type OfficeHostType = 'Word' | 'Excel' | 'Unknown'

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
    }
  } catch {
    // Fallback: check global objects
  }

  if (detectedHost === 'Unknown') {
    if (typeof (window as any).Word !== 'undefined') {
      detectedHost = 'Word'
    } else if (typeof (window as any).Excel !== 'undefined') {
      detectedHost = 'Excel'
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

export function getHostName(): string {
  return detectOfficeHost()
}
