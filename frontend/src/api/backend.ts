import type { ModelTier, ModelInfo } from '@/types'
const BACKEND_URL = import.meta.env.VITE_BACKEND_URL

if (!BACKEND_URL) {
  throw new Error('VITE_BACKEND_URL is required. Please define it in frontend/.env')
}

// Timeouts by model tier — reasoning models need more time (up to 6 min LLM + overhead)
const BASE_TIMEOUT_MS = Number(import.meta.env.VITE_REQUEST_TIMEOUT_MS) || 180_000
const TIMEOUT_BY_TIER: Record<string, number> = {
  reasoning: 360_000,
  standard: 180_000,
  fast: 120_000,
}

function getTimeoutForTier(modelTier?: string): number {
  if (modelTier && TIMEOUT_BY_TIER[modelTier]) return TIMEOUT_BY_TIER[modelTier]
  return BASE_TIMEOUT_MS
}

const RETRY_DELAYS_MS = [1_500, 4_000] as const

function wait(ms: number): Promise<void> {
  return new Promise((resolve) => {
    setTimeout(resolve, ms)
  })
}

function isRetryableError(error: unknown): boolean {
  return error instanceof TypeError || (error instanceof DOMException && error.name === 'TimeoutError')
}

// ────────────────────────────────────────────────────────────────────────────
// Error categorisation — exposes structured info for user-facing messages
// ────────────────────────────────────────────────────────────────────────────
export type ErrorType = 'timeout' | 'network' | 'rate_limit' | 'auth' | 'server' | 'unknown'

export interface CategorizedError {
  type: ErrorType
  /** i18n key to use in the UI */
  i18nKey: string
}

export function categorizeError(error: unknown): CategorizedError {
  if (error instanceof DOMException && error.name === 'AbortError') {
    return { type: 'unknown', i18nKey: 'generationStop' }
  }
  if (error instanceof DOMException && error.name === 'TimeoutError') {
    return { type: 'timeout', i18nKey: 'errorTimeout' }
  }
  if (error instanceof TypeError) {
    // Fetch TypeError typically means network unreachable
    return { type: 'network', i18nKey: 'errorNetwork' }
  }
  const msg = (error instanceof Error ? error.message : String(error)).toLowerCase()
  if (msg.includes('401') || msg.includes('credentials') || msg.includes('x-user-key') || msg.includes('x-user-email')) {
    return { type: 'auth', i18nKey: 'credentialsRequired' }
  }
  if (msg.includes('429') || msg.includes('rate limit') || msg.includes('too many')) {
    return { type: 'rate_limit', i18nKey: 'errorRateLimit' }
  }
  if (msg.includes('500') || msg.includes('502') || msg.includes('503') || msg.includes('internal server')) {
    return { type: 'server', i18nKey: 'errorServer' }
  }
  if (msg.includes('timeout') || msg.includes('timed out')) {
    return { type: 'timeout', i18nKey: 'errorTimeout' }
  }
  return { type: 'unknown', i18nKey: 'failedToResponse' }
}

function createTimeoutSignal(timeoutMs: number, externalSignal?: AbortSignal): { signal: AbortSignal; cleanup: () => void } {
  const timeoutController = new AbortController()

  const timeoutId = setTimeout(() => {
    timeoutController.abort(new DOMException('Request timed out', 'TimeoutError'))
  }, timeoutMs)

  const abortFromExternal = () => {
    timeoutController.abort(externalSignal?.reason)
  }

  if (externalSignal) {
    if (externalSignal.aborted) {
      abortFromExternal()
    } else {
      externalSignal.addEventListener('abort', abortFromExternal, { once: true })
    }
  }

  return {
    signal: timeoutController.signal,
    cleanup: () => {
      clearTimeout(timeoutId)
      externalSignal?.removeEventListener('abort', abortFromExternal)
    },
  }
}

async function fetchWithTimeoutAndRetry(url: string, init: RequestInit = {}, modelTier?: string): Promise<Response> {
  let attempt = 0
  const timeoutMs = getTimeoutForTier(modelTier)

  while (true) {
    const { signal, cleanup } = createTimeoutSignal(timeoutMs, init.signal ?? undefined)

    try {
      return await fetch(url, {
        ...init,
        credentials: 'include',
        signal,
      })
    } catch (error) {
      if (init.signal?.aborted) {
        throw error
      }

      const isPost = init.method?.toUpperCase() === 'POST'
      // Allow 1 retry on POST for timeout/network errors (transient failures)
      const maxPostRetries = 1
      const shouldRetry =
        attempt < RETRY_DELAYS_MS.length &&
        isRetryableError(error) &&
        (!isPost || attempt < maxPostRetries)
      if (!shouldRetry) {
        throw error
      }

      await wait(RETRY_DELAYS_MS[attempt])
      attempt += 1
    } finally {
      cleanup()
    }
  }
}


import { getUserKey, getUserEmail } from '@/utils/credentialStorage'

function getCsrfToken(): string {
  const match = document.cookie.match(/(?:^| )csrf_token=([^;]+)/)
  if (match) return match[1]
  return ''
}

function getUserCredentialHeaders(): Record<string, string> {
  const userKey = getUserKey()
  const userEmail = getUserEmail()
  const headers: Record<string, string> = {}
  if (userKey) headers['X-User-Key'] = userKey
  if (userEmail) headers['X-User-Email'] = userEmail
  
  const csrf = getCsrfToken()
  if (csrf) headers['x-csrf-token'] = csrf
  
  return headers
}

export async function fetchModels(): Promise<Record<string, ModelInfo>> {
  const res = await fetchWithTimeoutAndRetry(`${BACKEND_URL}/api/models`, {
    headers: { ...getUserCredentialHeaders() }
  })
  if (!res.ok) throw new Error(`Failed to fetch models: ${res.status}`)
  return res.json()
}

export async function healthCheck(): Promise<boolean> {
  try {
    const res = await fetchWithTimeoutAndRetry(`${BACKEND_URL}/health`, {
      headers: { ...getUserCredentialHeaders() }
    })
    return res.ok
  } catch {
    return false
  }
}

export interface ChatMessage {
  role: 'system' | 'user' | 'assistant'
  content: string
  tool_calls?: Array<{
    id: string
    type: 'function'
    function: {
      name: string
      arguments: string
    }
  }>
}

export interface ToolChatMessage {
  role: 'tool'
  tool_call_id: string
  content: string
}

export type ChatRequestMessage = ChatMessage | ToolChatMessage

export interface TokenUsage {
  promptTokens: number
  completionTokens: number
  totalTokens: number
}

export interface ChatStreamOptions {
  messages: ChatRequestMessage[]
  modelTier: ModelTier
  tools?: ApiToolDefinition[]
  onStream: (text: string) => void
  onToolCallDelta?: (toolCallDeltas: any[]) => void
  onFinishReason?: (finishReason: string | null) => void
  onUsage?: (usage: TokenUsage) => void
  abortSignal?: AbortSignal
}

export async function chatStream(options: ChatStreamOptions): Promise<void> {
  const { messages, modelTier, tools, onStream, onToolCallDelta, onFinishReason, onUsage, abortSignal } = options

  const res = await fetchWithTimeoutAndRetry(`${BACKEND_URL}/api/chat`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', ...getUserCredentialHeaders() },
    body: JSON.stringify({ messages, modelTier, tools, stream_options: { include_usage: true } }),
    signal: abortSignal,
  }, modelTier)

  if (!res.ok) {
    const err = await res.text()
    throw new Error(`Chat API error ${res.status}: ${err}`)
  }

  if (!res.body) throw new Error('Empty response body')
  const reader = res.body.getReader()
  const decoder = new TextDecoder()
  let fullContent = ''
  let buffer = ''

  while (true) {
    const { done, value } = await reader.read()
    if (done) break

    buffer += decoder.decode(value, { stream: true })
    const lines = buffer.split('\n')
    buffer = lines.pop() || ''

    for (const line of lines) {
      const trimmedLine = line.trim()
      if (!trimmedLine || !trimmedLine.startsWith('data: ')) continue
      const data = trimmedLine.slice(6)
      if (data === '[DONE]') return

      try {
        const parsed = JSON.parse(data)

        // Detect error objects embedded in the SSE stream
        if (parsed.error) {
          const errMsg = parsed.error.message || JSON.stringify(parsed.error)
          throw new Error(`Stream error: ${errMsg}`)
        }

        const finishReason = parsed.choices?.[0]?.finish_reason ?? null
        if (finishReason !== null) {
          onFinishReason?.(finishReason)
        }
        const delta = parsed.choices?.[0]?.delta
        if (delta?.content) {
          fullContent += delta.content
          onStream(fullContent)
        }
        if (delta?.tool_calls?.length && onToolCallDelta) {
          onToolCallDelta(delta.tool_calls)
        }
        // Capture token usage from final SSE chunk
        if (parsed.usage && onUsage) {
          onUsage({
            promptTokens: parsed.usage.prompt_tokens ?? 0,
            completionTokens: parsed.usage.completion_tokens ?? 0,
            totalTokens: parsed.usage.total_tokens ?? 0,
          })
        }
      } catch (parseError) {
        // Only re-throw if it was our own explicit stream error; skip malformed JSON
        if (parseError instanceof Error && parseError.message.startsWith('Stream error:')) {
          throw parseError
        }
        // Otherwise skip malformed JSON line silently
      }
    }
  }
}

export interface ApiToolDefinition {
  type: 'function'
  function: {
    name: string
    description?: string
    parameters: Record<string, any>
    strict?: boolean
  }
}

export interface ImageGenerateOptions {
  prompt: string
  size?: string
  quality?: string
}

export async function generateImage(options: ImageGenerateOptions): Promise<string> {
  const res = await fetchWithTimeoutAndRetry(`${BACKEND_URL}/api/image`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', ...getUserCredentialHeaders() },
    body: JSON.stringify(options),
  })

  if (!res.ok) {
    const err = await res.text()
    throw new Error(`Image API error ${res.status}: ${err}`)
  }

  const data = await res.json()
  const image = data.data?.[0]

  if (image?.b64_json) {
    return `data:image/png;base64,${image.b64_json}`
  }

  if (image?.url) {
    return image.url
  }

  return ''
}

export async function uploadFile(file: File): Promise<{ filename: string; extractedText: string; imageBase64?: string }> {
  const formData = new FormData()
  formData.append('file', file)

  const res = await fetchWithTimeoutAndRetry(`${BACKEND_URL}/api/upload`, {
    method: 'POST',
    headers: { ...getUserCredentialHeaders() },
    body: formData,
  })

  if (!res.ok) {
    const err = await res.text()
    throw new Error(`File upload error ${res.status}: ${err}`)
  }

  return res.json()
}
