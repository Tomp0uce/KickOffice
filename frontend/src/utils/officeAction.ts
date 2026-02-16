const OFFICE_ACTION_TIMEOUT_MS = 10_000
const OFFICE_BUSY_TIMEOUT_MESSAGE = 'Office app is busy. Please exit cell editing or close dialogs.'

export async function executeOfficeAction<T>(action: () => Promise<T>): Promise<T> {
  let timeoutId: ReturnType<typeof setTimeout> | undefined

  const timeoutPromise = new Promise<never>((_, reject) => {
    timeoutId = setTimeout(() => {
      reject(new Error(OFFICE_BUSY_TIMEOUT_MESSAGE))
    }, OFFICE_ACTION_TIMEOUT_MS)
  })

  try {
    return await Promise.race([action(), timeoutPromise])
  } finally {
    if (timeoutId) {
      clearTimeout(timeoutId)
    }
  }
}

export { OFFICE_ACTION_TIMEOUT_MS, OFFICE_BUSY_TIMEOUT_MESSAGE }
