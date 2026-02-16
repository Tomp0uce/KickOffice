interface OfficeAsyncResult<TValue = unknown> {
  status: string
  value?: TValue
  error?: {
    message?: string
  }
}

interface OutlookBody {
  getAsync(coercionType: string, callback: (result: OfficeAsyncResult<string>) => void): void
  setAsync(
    data: string,
    options: { coercionType: string },
    callback: (result: OfficeAsyncResult) => void,
  ): void
}

interface OutlookItem {
  body: OutlookBody
  getSelectedDataAsync?: (
    coercionType: string,
    callback: (result: OfficeAsyncResult<{ data?: string }>) => void,
  ) => void
}

interface OutlookMailbox {
  item?: OutlookItem
}

interface OfficeRuntime {
  context?: {
    mailbox?: OutlookMailbox
  }
  CoercionType: {
    Text: string
  }
  AsyncResultStatus: {
    Succeeded: string
  }
}

function getOfficeRuntime(): OfficeRuntime | null {
  return (window as unknown as { Office?: OfficeRuntime }).Office ?? null
}

function getOutlookMailbox(): OutlookMailbox | null {
  return getOfficeRuntime()?.context?.mailbox ?? null
}

function getOfficeTextCoercionType(): string {
  return getOfficeRuntime()?.CoercionType.Text ?? 'text'
}

function isOfficeAsyncSucceeded(status: string): boolean {
  return status === getOfficeRuntime()?.AsyncResultStatus.Succeeded
}

export {
  getOfficeTextCoercionType,
  getOutlookMailbox,
  isOfficeAsyncSucceeded,
  type OfficeAsyncResult,
}
