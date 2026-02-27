const BACKEND_URL = import.meta.env.VITE_BACKEND_URL

if (!BACKEND_URL) {
  throw new Error('VITE_BACKEND_URL is required. Please define it in frontend/.env')
}

const REQUEST_TIMEOUT_MS = Number(import.meta.env.VITE_REQUEST_TIMEOUT_MS) || 45_000
const RETRY_DELAYS_MS = [1_000, 3_000, 5_000] as const

function wait(ms: number): Promise<void> {
  return new Promise((resolve) => {
    setTimeout(resolve, ms)
  })
}

function isRetryableError(error: unknown): boolean {
  return error instanceof TypeError || (error instanceof DOMException && error.name === 'TimeoutError')
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

async function fetchWithTimeoutAndRetry(url: string, init: RequestInit = {}): Promise<Response> {
  let attempt = 0

  while (true) {
    const { signal, cleanup } = createTimeoutSignal(REQUEST_TIMEOUT_MS, init.signal ?? undefined)

    try {
      return await fetch(url, {
        ...init,
        signal,
      })
    } catch (error) {
      if (init.signal?.aborted) {
        throw error
      }

      const isPost = init.method?.toUpperCase() === 'POST'
      const shouldRetry = attempt < RETRY_DELAYS_MS.length && isRetryableError(error) && !isPost
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

function getUserCredentialHeaders(): Record<string, string> {
  const userKey = getUserKey()
  const userEmail = getUserEmail()
  const headers: Record<string, string> = {}
  if (userKey) headers['X-User-Key'] = userKey
  if (userEmail) headers['X-User-Email'] = userEmail
  return headers
}

export async function fetchModels(): Promise<Record<string, ModelInfo>> {
  const res = await fetchWithTimeoutAndRetry(`${BACKEND_URL}/api/models`)
  if (!res.ok) throw new Error(`Failed to fetch models: ${res.status}`)
  return res.json()
}

export async function healthCheck(): Promise<boolean> {
  try {
    const res = await fetchWithTimeoutAndRetry(`${BACKEND_URL}/health`)
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

export interface ChatStreamOptions {
  messages: ChatRequestMessage[]
  modelTier: ModelTier
  tools?: ToolDefinition[]
  onStream: (text: string) => void
  onToolCallDelta?: (toolCallDeltas: any[]) => void
  onFinishReason?: (finishReason: string | null) => void
  abortSignal?: AbortSignal
}

export async function chatStream(options: ChatStreamOptions): Promise<void> {
  const { messages, modelTier, tools, onStream, onToolCallDelta, onFinishReason, abortSignal } = options

  const res = await fetchWithTimeoutAndRetry(`${BACKEND_URL}/api/chat`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', ...getUserCredentialHeaders() },
    body: JSON.stringify({ messages, modelTier, tools }),
    signal: abortSignal,
  })

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
      } catch {
        // Skip malformed JSON lines
      }
    }
  }
}

export interface ToolDefinition {
  type: 'function'
  function: {
    name: string
    description?: string
    parameters: Record<string, unknown>
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
