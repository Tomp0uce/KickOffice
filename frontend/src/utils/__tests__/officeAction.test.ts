import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest'
import { executeOfficeAction } from '../officeAction'
import { logService } from '../logger'

vi.mock('../logger', () => ({
  logService: {
    warn: vi.fn(),
    error: vi.fn(),
  }
}))

describe('officeAction', () => {
  beforeEach(() => {
    vi.useFakeTimers()
    vi.clearAllMocks()
  })

  afterEach(() => {
    vi.useRealTimers()
  })

  it('should resolve successfully when action completes quickly', async () => {
    const action = vi.fn().mockResolvedValue('success')
    const promise = executeOfficeAction(action, 'test_action', 5000)
    
    vi.advanceTimersByTime(10)
    const result = await promise
    
    expect(result).toBe('success')
    expect(action).toHaveBeenCalledTimes(1)
    expect(logService.error).not.toHaveBeenCalled()
  })

  it('should timeout and throw busy error if action is too slow', async () => {
    const action = vi.fn().mockImplementation(() => new Promise(resolve => setTimeout(() => resolve('too_late'), 10000)))
    
    const promise = executeOfficeAction(action, 'slow_action', 5000)
    const expectPromise = expect(promise).rejects.toThrow('Office app is busy. Please exit cell editing or close dialogs.')
    
    await vi.advanceTimersByTimeAsync(20000)
    
    await expectPromise
    expect(logService.warn).toHaveBeenCalledWith(
      expect.stringContaining('Timeout executing slow_action'),
      expect.anything()
    )
  })

  it('should retry on GeneralException up to maxRetries', async () => {
    let callCount = 0
    const action = vi.fn().mockImplementation(() => {
      callCount++
      if (callCount < 3) {
        return Promise.reject(new Error('Some GeneralException occurred'))
      }
      return Promise.resolve('success_on_retry')
    })

    const promise = executeOfficeAction(action, 'retry_action', 5000)
    
    // Attempt 1 fails, waits 1000ms
    await vi.advanceTimersByTimeAsync(1)
    expect(logService.warn).toHaveBeenCalledWith(
      expect.stringContaining('Retrying retry_action after error'),
      expect.anything()
    )
    
    await vi.advanceTimersByTimeAsync(1000)
    
    // Attempt 2 fails, waits 2000ms
    await vi.advanceTimersByTimeAsync(1)
    
    await vi.advanceTimersByTimeAsync(2000)
    
    // Attempt 3 succeeds
    const result = await promise
    expect(result).toBe('success_on_retry')
    expect(callCount).toBe(3)
  })

  it('should throw immediately on non-retriable error', async () => {
    const action = vi.fn().mockRejectedValue(new Error('SyntaxError'))
    
    const promise = executeOfficeAction(action, 'syntax_action', 5000)
    
    await expect(promise).rejects.toThrow('SyntaxError')
    expect(action).toHaveBeenCalledTimes(1)
    expect(logService.error).toHaveBeenCalled()
  })

  it('should abort when abortSignal is triggered before action', async () => {
    const controller = new AbortController()
    controller.abort() // abort before even calling
    
    const action = vi.fn().mockResolvedValue('success')
    
    const promise = executeOfficeAction(action, 'abort_action', 5000, controller.signal)
    
    await expect(promise).rejects.toThrow('Operation aborted by user')
    expect(action).not.toHaveBeenCalled()
  })

  it('should abort when abortSignal is triggered during action', async () => {
    const controller = new AbortController()
    
    let resolveAction: (v: string) => void = () => {}
    const action = vi.fn().mockImplementation(() => new Promise(resolve => {
      resolveAction = resolve
    }))
    
    const promise = executeOfficeAction(action, 'abort_action', 15000, controller.signal)
    // Attach error handler immediately to avoid unhandled rejection warnings
    const expectPromise = expect(promise).rejects.toThrow('Operation aborted by user')
    
    // Wait a little, then abort
    await vi.advanceTimersByTimeAsync(100)
    controller.abort()
    
    await expectPromise
    expect(logService.warn).toHaveBeenCalledWith(
      expect.stringContaining('Aborted abort_action'),
      expect.anything()
    )
    
    // Resolve action to clean up pending promises
    resolveAction('too_late')
  })
})
