/**
 * PowerPoint interaction utilities.
 *
 * Unlike Word (Word.run) or Excel (Excel.run), the PowerPoint web text
 * manipulation API relies on the Common API (Office.context.document).
 * These helpers wrap the async callbacks in Promises.
 */

declare const Office: any

type ParsedListLine = {
  indentLevel: number
  text: string
}

function parseMarkdownListLine(line: string): ParsedListLine | null {
  const bulletMatch = line.match(/^(\s*)[-*+]\s+(.+)$/)
  if (bulletMatch) {
    const [, indent, itemText] = bulletMatch
    return {
      indentLevel: Math.floor(indent.replace(/\t/g, '  ').length / 2),
      text: itemText.trim(),
    }
  }

  const numberedMatch = line.match(/^(\s*)\d+[.)]\s+(.+)$/)
  if (numberedMatch) {
    const [, indent, itemText] = numberedMatch
    return {
      indentLevel: Math.floor(indent.replace(/\t/g, '  ').length / 2),
      text: itemText.trim(),
    }
  }

  return null
}

/**
 * PowerPoint keeps existing paragraph bullet formatting when replacing text
 * inside a bulleted shape. If we insert markdown markers (-, *, 1.) directly,
 * users can end up with duplicated bullets (native bullet + literal marker).
 *
 * This converter removes list markers while preserving hierarchy via tabs.
 */
export function normalizePowerPointListText(text: string): string {
  const lines = text.split(/\r?\n/)
  let hasListSyntax = false

  const normalizedLines = lines.map((line) => {
    const parsedLine = parseMarkdownListLine(line)
    if (!parsedLine) {
      return line
    }

    hasListSyntax = true
    return `${'\t'.repeat(parsedLine.indentLevel)}${parsedLine.text}`
  })

  return hasListSyntax ? normalizedLines.join('\n') : text
}

/**
 * Read the currently selected text inside a PowerPoint shape / text box.
 * Returns an empty string when nothing is selected or the selection is
 * not a text range (e.g. an entire slide is selected).
 */
export function getPowerPointSelection(): Promise<string> {
  return new Promise((resolve) => {
    try {
      Office.context.document.getSelectedDataAsync(
        Office.CoercionType.Text,
        { valueFormat: Office.ValueFormat.Unformatted },
        (result: any) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            resolve((result.value as string) || '')
          } else {
            console.warn('PowerPoint selection error:', result.error?.message)
            resolve('')
          }
        },
      )
    } catch (err) {
      console.warn('PowerPoint getSelectedDataAsync unavailable:', err)
      resolve('')
    }
  })
}

/**
 * Replace the current text selection inside the active PowerPoint shape
 * with the provided text.
 */
export function insertIntoPowerPoint(text: string): Promise<void> {
  return new Promise((resolve, reject) => {
    const normalizedText = normalizePowerPointListText(text)

    try {
      Office.context.document.setSelectedDataAsync(
        normalizedText,
        { coercionType: Office.CoercionType.Text },
        (result: any) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            resolve()
          } else {
            reject(new Error(result.error?.message || 'setSelectedDataAsync failed'))
          }
        },
      )
    } catch (err: any) {
      reject(new Error(err?.message || 'setSelectedDataAsync unavailable'))
    }
  })
}
