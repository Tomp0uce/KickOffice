import { logService } from './logger'

const OFFICE_ACTION_TIMEOUT_MS = 10_000
const OFFICE_BUSY_TIMEOUT_MESSAGE = 'Office app is busy. Please exit cell editing or close dialogs.'

export async function executeOfficeAction<T>(action: () => Promise<T>, actionName: string = 'unknown_action'): Promise<T> {
  let timeoutId: ReturnType<typeof setTimeout> | undefined

  const timeoutPromise = new Promise<never>((_, reject) => {
    timeoutId = setTimeout(() => {
      logService.warn(`[OfficeAction] Timeout executing ${actionName}`, { actionName })
      reject(new Error(OFFICE_BUSY_TIMEOUT_MESSAGE))
    }, OFFICE_ACTION_TIMEOUT_MS)
  })

  try {
    return await Promise.race([action(), timeoutPromise])
  } catch (err) {
    if (err instanceof Error && err.message !== OFFICE_BUSY_TIMEOUT_MESSAGE) {
      logService.error(`[OfficeAction] Error executing ${actionName}`, err, { actionName })
    }
    throw err
  } finally {
    if (timeoutId) {
      clearTimeout(timeoutId)
    }
  }
}

