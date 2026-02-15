/**
 * PowerPoint interaction utilities.
 *
 * Unlike Word (Word.run) or Excel (Excel.run), the PowerPoint web text
 * manipulation API relies on the Common API (Office.context.document).
 * These helpers wrap the async callbacks in Promises.
 */

declare const Office: any

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
    try {
      Office.context.document.setSelectedDataAsync(
        text,
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
