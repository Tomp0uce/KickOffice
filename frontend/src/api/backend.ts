import type { ModelTier, ModelInfo } from '@/types'
const BACKEND_URL = import.meta.env.VITE_BACKEND_URL

if (!BACKEND_URL) {
  throw new Error('VITE_BACKEND_URL is required. Please define it in frontend/.env')
}

// Timeouts by model tier — reasoning models need more time (up to 6 min LLM + overhead)
const BASE_TIMEOUT_MS = Number(import.meta.env.VITE_REQUEST_TIMEOUT_MS) || 180_000
const TIMEOUT_BY_TIER: Record<string, number> = {
  reasoning: 600_000,  // 10 min — GPT-5.2 up to 65k output tokens
  standard: 300_000,   // 5 min — GPT-5.2 up to 32k output tokens
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

/**
 * Point 1 Fix: Prevents massive Base64 data from saturating backend/terminal logs.
 * Truncates image_url data for logging purposes.
 */
function sanitizePayloadForLogs(payload: any) {
  try {
    const clone = JSON.parse(JSON.stringify(payload));
    if (clone.messages) {
      clone.messages.forEach((msg: any) => {
        if (Array.isArray(msg.content)) {
          msg.content.forEach((part: any) => {
            if (part.type === 'image_url' && part.image_url?.url) {
              part.image_url.url = '[BASE64_IMAGE_DATA_TRUNCATED_FOR_LOGS]';
            }
          });
        }
      });
    }
    return clone;
  } catch {
    return payload;
  }
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

/** Maps backend error codes to i18n keys. Falls back to message inspection if no code present. */
const ERROR_CODE_MAP: Record<string, CategorizedError> = {
  VALIDATION_ERROR: { type: 'unknown', i18nKey: 'failedToResponse' },
  AUTH_REQUIRED: { type: 'auth', i18nKey: 'credentialsRequired' },
  RATE_LIMITED: { type: 'rate_limit', i18nKey: 'errorRateLimit' },
  LLM_UPSTREAM_ERROR: { type: 'server', i18nKey: 'errorServer' },
  LLM_EMPTY_RESPONSE: { type: 'server', i18nKey: 'errorServer' },
  LLM_INVALID_JSON: { type: 'server', i18nKey: 'errorServer' },
  LLM_NO_CHOICES: { type: 'server', i18nKey: 'errorServer' },
  LLM_CONTENT_FILTERED: { type: 'server', i18nKey: 'errorServer' },
  LLM_TIMEOUT: { type: 'timeout', i18nKey: 'errorTimeout' },
  IMAGE_TIMEOUT: { type: 'timeout', i18nKey: 'errorTimeout' },
  INTERNAL_ERROR: { type: 'server', i18nKey: 'errorServer' },
  PDF_EXTRACTION_FAILED: { type: 'unknown', i18nKey: 'failedToResponse' },
  DOCX_EXTRACTION_FAILED: { type: 'unknown', i18nKey: 'failedToResponse' },
  NO_FILE_UPLOADED: { type: 'unknown', i18nKey: 'failedToResponse' },
  UNSUPPORTED_FILE_TYPE: { type: 'unknown', i18nKey: 'failedToResponse' },
  FILE_EMPTY: { type: 'unknown', i18nKey: 'failedToResponse' },
  CHART_IMAGE_NOT_FOUND: { type: 'unknown', i18nKey: 'failedToResponse' },
  CHART_EXTRACTION_FAILED: { type: 'unknown', i18nKey: 'failedToResponse' },
}

export function categorizeError(error: unknown): CategorizedError {
  if (error instanceof DOMException && error.name === 'AbortError') {
    return { type: 'unknown', i18nKey: 'generationStop' }
  }
  if (error instanceof DOMException && error.name === 'TimeoutError') {
    return { type: 'timeout', i18nKey: 'errorTimeout' }
  }
  if (error instanceof TypeError) {
    return { type: 'network', i18nKey: 'errorNetwork' }
  }

  // Try structured error code first (from backend ErrorCodes registry)
  if (error instanceof Error && 'code' in error) {
    const mapped = ERROR_CODE_MAP[(error as any).code]
    if (mapped) return mapped
  }

  // Fallback: inspect error message string
  const msg = (error instanceof Error ? error.message : String(error)).toLowerCase()
  if (msg.includes('401') || msg.includes('403') || msg.includes('credentials') || msg.includes('x-user-key') || msg.includes('x-user-email')) {
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
        logService.error(`Network request failed: ${url}`, error)
        throw error
      }
      
      logService.warn(`Network retry ${attempt + 1}/${RETRY_DELAYS_MS.length} for ${url}`, error)
      await wait(RETRY_DELAYS_MS[attempt])
      attempt += 1
    } finally {
      cleanup()
    }
  }
}


import { getUserKey, getUserEmail } from '@/utils/credentialStorage'
import { logService } from '@/utils/logger'

function getCsrfToken(): string {
  const match = document.cookie.match(/(?:^| )csrf_token=([^;]+)/)
  if (match) return match[1]
  return ''
}

// ─── QC-L2: Header cache ─────────────────────────────────────────────────────
// getGlobalHeaders() is called on every network request. The calls to
// getUserKey(), getUserEmail() and logService.getContext() touch async storage
// on each invocation. We cache the resolved headers and expose
// invalidateHeaderCache() so callers can bust the cache after credential changes.
let _headerCache: Promise<Record<string, string>> | null = null

/** Bust the cached credentials headers (call after saving new credentials). */
export function invalidateHeaderCache(): void {
  _headerCache = null
}

async function buildGlobalHeaders(): Promise<Record<string, string>> {
  const userKey = await getUserKey()
  const userEmail = await getUserEmail()
  const ctx = await logService.getContext()

  const headers: Record<string, string> = {}
  if (userKey) headers['X-User-Key'] = userKey
  if (userEmail) headers['X-User-Email'] = userEmail
  if (ctx.host) headers['X-Office-Host'] = ctx.host
  if (ctx.sessionId) headers['X-Session-Id'] = ctx.sessionId

  return headers
}

async function getGlobalHeaders(): Promise<Record<string, string>> {
  if (!_headerCache) {
    _headerCache = buildGlobalHeaders()
  }
  const cached = await _headerCache
  // CSRF token reads from document.cookie — always fresh (not cached)
  const csrf = getCsrfToken()
  if (csrf) return { ...cached, 'x-csrf-token': csrf }
  return cached
}

export async function fetchModels(): Promise<Record<string, ModelInfo>> {
  const res = await fetchWithTimeoutAndRetry(`${BACKEND_URL}/api/models`, {
    headers: { ...(await getGlobalHeaders()) }
  })
  if (!res.ok) throw new Error(`Failed to fetch models: ${res.status}`)
  return res.json()
}

export async function healthCheck(): Promise<boolean> {
  try {
    const res = await fetchWithTimeoutAndRetry(`${BACKEND_URL}/health`, {
      headers: { ...(await getGlobalHeaders()) }
    })
    return res.ok
  } catch {
    return false
  }
}

export interface ChatMessage {
  role: 'system' | 'user' | 'assistant'
  content: string | any[]
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
    headers: { 'Content-Type': 'application/json', ...(await getGlobalHeaders()) },
    body: JSON.stringify({ messages, modelTier, tools, stream_options: { include_usage: true } }),
    signal: abortSignal,
  }, modelTier)

  if (!res.ok) {
    const err = await res.text()
    const sanitizedBody = sanitizePayloadForLogs({ messages, modelTier, tools })
    logService.error('Chat API error', undefined, { status: res.status, error: err, body: sanitizedBody })
    throw new Error(`Chat API error ${res.status}: ${err}`)
  }

  const chatReqId = res.headers.get('x-request-id')
  if (chatReqId) logService.info(`Request correlated: ${chatReqId}`, 'system', { reqId: chatReqId })

  if (!res.body) throw new Error('Empty response body')
  const reader = res.body.getReader()
  const decoder = new TextDecoder()
  let fullContent = ''
  let buffer = ''

  while (true) {
    const { done, value } = await reader.read()
    if (done) break

    buffer += decoder.decode(value, { stream: true })
    
    // Safety check against unbounded memory growth (5MB limit)
    if (buffer.length > 5 * 1024 * 1024) {
      throw new Error('SSE stream buffer exceeded maximum allowed size')
    }
    
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
        // Re-throw explicit stream errors
        if (parseError instanceof Error && parseError.message.startsWith('Stream error:')) {
          throw parseError
        }
        // Log malformed JSON
        logService.warn('Malformed JSON in chatStream SSE', {
          data: data.length > 200 ? data.slice(0, 200) + '...' : data,
          error: parseError instanceof Error ? parseError.message : String(parseError)
        })
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
  const payload = { ...options, size: options.size || '1024x1024' }
  const res = await fetchWithTimeoutAndRetry(`${BACKEND_URL}/api/image`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', ...(await getGlobalHeaders()) },
    body: JSON.stringify(payload),
  })

  if (!res.ok) {
    const err = await res.text()
    logService.error('Image API error', undefined, { status: res.status, error: err })
    throw new Error(`Image API error ${res.status}: ${err}`)
  }

  const imageReqId = res.headers.get('x-request-id')
  if (imageReqId) logService.info(`Request correlated: ${imageReqId}`, 'system', { reqId: imageReqId })

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

export async function uploadFile(file: File): Promise<{ filename: string; extractedText: string; imageBase64?: string; imageId?: string }> {
  const formData = new FormData()
  formData.append('file', file)

  const res = await fetchWithTimeoutAndRetry(`${BACKEND_URL}/api/upload`, {
    method: 'POST',
    headers: { ...(await getGlobalHeaders()) },
    body: formData,
  })

  if (!res.ok) {
    const err = await res.text()
    logService.error('File upload error', undefined, { status: res.status, error: err })
    throw new Error(`File upload error ${res.status}: ${err}`)
  }

  const uploadReqId = res.headers.get('x-request-id')
  if (uploadReqId) logService.info(`Request correlated: ${uploadReqId}`, 'system', { reqId: uploadReqId })

  return res.json()
}

/**
 * Upload a file to the LLM provider via the backend proxy.
 * Returns a file_id that can be referenced in subsequent LLM messages.
 * May throw if the provider does not support the /v1/files API.
 */
export async function uploadFileToPlatform(file: File, purpose = 'assistants'): Promise<{ fileId: string }> {
  const formData = new FormData()
  formData.append('file', file)
  formData.append('purpose', purpose)

  const res = await fetchWithTimeoutAndRetry(`${BACKEND_URL}/api/files`, {
    method: 'POST',
    headers: { ...(await getGlobalHeaders()) },
    body: formData,
  })

  if (!res.ok) {
    const err = await res.text()
    throw new Error(`File platform upload error ${res.status}: ${err}`)
  }

  return res.json()
}

export async function submitLogs(entries: unknown[]): Promise<void> {
  try {
    const res = await fetchWithTimeoutAndRetry(`${BACKEND_URL}/api/logs`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', ...(await getGlobalHeaders()) },
      body: JSON.stringify({ entries }),
    })
    if (!res.ok) {
      logService.originalConsole.warn('[KO] Failed to submit logs:', res.status)
    }
  } catch {
    // Silent: log submission failure should never break the UI
  }
}

export interface PlotAreaBox {
  /** Left edge of the chart's plot area. Value in [0,1] = fraction of image width; value > 1 = raw pixels. */
  xMinPx: number
  /** Right edge of the chart's plot area. */
  xMaxPx: number
  /** Top edge of the chart's plot area (smaller pixel value = higher on screen). */
  yMinPx: number
  /** Bottom edge of the chart's plot area (larger pixel value = lower on screen, where X axis sits). */
  yMaxPx: number
}

export interface ChartExtractParams {
  imageId: string
  xAxisRange: [number, number]
  yAxisRange: [number, number]
  targetColor: string
  plotAreaBox: PlotAreaBox
  chartType?: string
  colorTolerance?: number
  numPoints?: number
}

export interface ChartExtractResult {
  points: Array<{ x: number; y: number }>
  pixelsMatched: number
  imageSize: { width: number; height: number }
  plotBounds?: { pxMin: number; pxMax: number; pyMin: number; pyMax: number }
  warning?: string
}

export async function extractChartData(params: ChartExtractParams): Promise<ChartExtractResult> {
  const res = await fetchWithTimeoutAndRetry(`${BACKEND_URL}/api/chart-extract`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', ...(await getGlobalHeaders()) },
    body: JSON.stringify(params),
  })

  if (!res.ok) {
    const err = await res.text()
    logService.error('Chart extraction error', undefined, { status: res.status, error: err })
    throw new Error(`Chart extraction error ${res.status}: ${err}`)
  }

  const reqId = res.headers.get('x-request-id')
  if (reqId) logService.info(`Request correlated: ${reqId}`, 'system', { reqId })

  return res.json()
}

export async function searchIconify(query: string, limit = 10, prefix?: string): Promise<any> {
  const params = new URLSearchParams({ query, limit: String(limit) })
  if (prefix) params.set('prefix', prefix)
  const res = await fetchWithTimeoutAndRetry(`${BACKEND_URL}/api/icons/search?${params}`, {
    headers: await getGlobalHeaders(),
  })
  if (!res.ok) throw new Error(`Icon search failed: ${res.status}`)
  return res.json()
}

export async function fetchIconSvg(prefix: string, name: string, color?: string): Promise<string> {
  const params = color ? new URLSearchParams({ color }) : undefined
  const url = `${BACKEND_URL}/api/icons/svg/${encodeURIComponent(prefix)}/${encodeURIComponent(name)}${params ? '?' + params : ''}`
  const res = await fetchWithTimeoutAndRetry(url, {
    headers: await getGlobalHeaders(),
  })
  if (!res.ok) throw new Error(`Icon SVG fetch failed: ${res.status}`)
  return res.text()
}

export interface FeedbackSystemContext {
  host: string
  appVersion: string
  modelTier: string
  userAgent: string
}

export async function submitFeedback(sessionId: string, payload: {
  category: string
  comment: string
  logs: unknown[]
  chatHistory?: unknown[]
  systemContext?: FeedbackSystemContext
}): Promise<{ success: boolean }> {
  const res = await fetchWithTimeoutAndRetry(`${BACKEND_URL}/api/feedback/${sessionId}`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      ...(await getGlobalHeaders()),
    },
    body: JSON.stringify(payload),
  })

  if (!res.ok) {
    const errText = await res.text()
    throw new Error(`Feedback submission failed: ${res.status} ${errText}`)
  }

  return res.json()
}
