import type { ModelInfo } from '@/types';
import { logService } from '@/utils/logger';
import { fetchWithTimeoutAndRetry, getGlobalHeaders, generateRequestId } from './httpClient';
import type {
  ChatRequestMessage,
  ChatStreamOptions,
  ApiToolDefinition,
  ImageGenerateOptions,
  ChartExtractParams,
  ChartExtractResult,
  FeedbackSystemContext,
} from './types';

// ─── Public API re-exports ────────────────────────────────────────────────────
export type { ErrorType, CategorizedError } from './errorCategorization';
export { categorizeError } from './errorCategorization';
export { invalidateHeaderCache } from './httpClient';
export type {
  ChatMessage,
  ToolChatMessage,
  ChatRequestMessage,
  TokenUsage,
  ChatStreamOptions,
  ApiToolDefinition,
  ImageGenerateOptions,
  PlotAreaBox,
  ChartExtractParams,
  ChartExtractResult,
  FeedbackSystemContext,
} from './types';

// ─────────────────────────────────────────────────────────────────────────────

const BACKEND_URL = import.meta.env.VITE_BACKEND_URL;

if (!BACKEND_URL) {
  throw new Error('VITE_BACKEND_URL is required. Please define it in frontend/.env');
}

interface LogPayload {
  messages?: ChatRequestMessage[];
  tools?: ApiToolDefinition[];
}

/**
 * Point 1 Fix: Prevents massive Base64 data from saturating backend/terminal logs.
 * Truncates image_url data for logging purposes.
 */
function sanitizePayloadForLogs(payload: LogPayload): LogPayload {
  try {
    const clone = JSON.parse(JSON.stringify(payload)) as LogPayload;
    if (clone.messages) {
      clone.messages.forEach((msg: ChatRequestMessage) => {
        if (Array.isArray(msg.content)) {
          (msg.content as Array<{ type: string; image_url?: { url: string } }>).forEach(part => {
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

export async function fetchModels(): Promise<Record<string, ModelInfo>> {
  const res = await fetchWithTimeoutAndRetry(`${BACKEND_URL}/api/models`, {
    headers: { ...(await getGlobalHeaders()) },
  });
  if (!res.ok) throw new Error(`Failed to fetch models: ${res.status}`);
  return res.json();
}

export async function healthCheck(): Promise<boolean> {
  try {
    const res = await fetchWithTimeoutAndRetry(`${BACKEND_URL}/health`, {
      headers: { ...(await getGlobalHeaders()) },
    });
    return res.ok;
  } catch {
    return false;
  }
}

export async function chatStream(options: ChatStreamOptions): Promise<void> {
  const {
    messages,
    modelTier,
    tools,
    onStream,
    onToolCallDelta,
    onFinishReason,
    onUsage,
    abortSignal,
  } = options;

  // ERR-L1: Per-request ID for frontend↔backend log correlation
  const requestId = generateRequestId();
  logService.debug(`[chatStream] requestId=${requestId}`, { traffic: 'llm' });

  const res = await fetchWithTimeoutAndRetry(
    `${BACKEND_URL}/api/chat`,
    {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'X-Request-Id': requestId,
        ...(await getGlobalHeaders()),
      },
      body: JSON.stringify({ messages, modelTier, tools, stream_options: { include_usage: true } }),
      signal: abortSignal,
    },
    modelTier,
  );

  if (!res.ok) {
    const errText = await res.text();
    const sanitizedBody = sanitizePayloadForLogs({ messages, tools });
    logService.error('Chat API error', undefined, {
      status: res.status,
      error: errText,
      body: sanitizedBody,
    });
    // Try to parse structured error response { code, error, detail? }
    let errCode: string | undefined;
    let errDetail: string | undefined;
    try {
      const parsed = JSON.parse(errText);
      errCode = parsed?.code;
      errDetail = parsed?.detail;
    } catch {
      // not JSON — keep defaults
    }
    const error = new Error(`Chat API error ${res.status}: ${errText}`) as Error & {
      code?: string;
      detail?: string;
    };
    if (errCode) error.code = errCode;
    if (errDetail) error.detail = errDetail;
    throw error;
  }

  const chatReqId = res.headers.get('x-request-id');
  if (chatReqId)
    logService.info(`Request correlated: ${chatReqId}`, 'system', { reqId: chatReqId });

  if (!res.body) throw new Error('Empty response body');
  const reader = res.body.getReader();
  const decoder = new TextDecoder();
  let fullContent = '';
  let buffer = '';

  while (true) {
    const { done, value } = await reader.read();
    if (done) break;

    buffer += decoder.decode(value, { stream: true });

    // Safety check against unbounded memory growth (5MB limit)
    if (buffer.length > 5 * 1024 * 1024) {
      throw new Error('SSE stream buffer exceeded maximum allowed size');
    }

    const lines = buffer.split('\n');
    buffer = lines.pop() || '';

    for (const line of lines) {
      const trimmedLine = line.trim();
      if (!trimmedLine || !trimmedLine.startsWith('data: ')) continue;
      const data = trimmedLine.slice(6);
      if (data === '[DONE]') return;

      try {
        const parsed = JSON.parse(data);

        // Detect error objects embedded in the SSE stream
        if (parsed.error) {
          const errMsg = parsed.error.message || JSON.stringify(parsed.error);
          throw new Error(`Stream error: ${errMsg}`);
        }

        const finishReason = parsed.choices?.[0]?.finish_reason ?? null;
        if (finishReason !== null) {
          onFinishReason?.(finishReason);
        }
        const delta = parsed.choices?.[0]?.delta;
        if (delta?.content) {
          fullContent += delta.content;
          onStream(fullContent);
        }
        if (delta?.tool_calls?.length && onToolCallDelta) {
          onToolCallDelta(delta.tool_calls);
        }
        // Capture token usage from final SSE chunk
        if (parsed.usage && onUsage) {
          onUsage({
            promptTokens: parsed.usage.prompt_tokens ?? 0,
            completionTokens: parsed.usage.completion_tokens ?? 0,
            totalTokens: parsed.usage.total_tokens ?? 0,
          });
        }
      } catch (parseError) {
        // Re-throw explicit stream errors
        if (parseError instanceof Error && parseError.message.startsWith('Stream error:')) {
          throw parseError;
        }
        // Log malformed JSON
        logService.warn('Malformed JSON in chatStream SSE', {
          data: data.length > 200 ? data.slice(0, 200) + '...' : data,
          error: parseError instanceof Error ? parseError.message : String(parseError),
        });
      }
    }
  }
}

export async function generateImage(options: ImageGenerateOptions): Promise<string> {
  // IMG-H1: default to landscape 1536x1024 to match PPT slide format and reduce cropping
  const { abortSignal, ...rest } = options;
  const payload = { ...rest, size: options.size || '1536x1024' };
  // IMG-H2: use 'reasoning' tier timeout (10 min) — image generation can take 2-5 min
  const res = await fetchWithTimeoutAndRetry(
    `${BACKEND_URL}/api/image`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', ...(await getGlobalHeaders()) },
      body: JSON.stringify(payload),
      signal: abortSignal,
    },
    'reasoning',
  );

  if (!res.ok) {
    const err = await res.text();
    logService.error('Image API error', undefined, { status: res.status, error: err });
    throw new Error(`Image API error ${res.status}: ${err}`);
  }

  const imageReqId = res.headers.get('x-request-id');
  if (imageReqId)
    logService.info(`Request correlated: ${imageReqId}`, 'system', { reqId: imageReqId });

  const data = await res.json();
  const image = data.data?.[0];

  if (image?.b64_json) {
    return `data:image/png;base64,${image.b64_json}`;
  }

  if (image?.url) {
    return image.url;
  }

  return '';
}

export async function uploadFile(
  file: File,
): Promise<{ filename: string; extractedText: string; imageBase64?: string; imageId?: string }> {
  const formData = new FormData();
  formData.append('file', file);

  const res = await fetchWithTimeoutAndRetry(`${BACKEND_URL}/api/upload`, {
    method: 'POST',
    headers: { ...(await getGlobalHeaders()) },
    body: formData,
  });

  if (!res.ok) {
    const err = await res.text();
    logService.error('File upload error', undefined, { status: res.status, error: err });
    throw new Error(`File upload error ${res.status}: ${err}`);
  }

  const uploadReqId = res.headers.get('x-request-id');
  if (uploadReqId)
    logService.info(`Request correlated: ${uploadReqId}`, 'system', { reqId: uploadReqId });

  return res.json();
}

/**
 * Upload a file to the LLM provider via the backend proxy.
 * Returns a file_id that can be referenced in subsequent LLM messages.
 * May throw if the provider does not support the /v1/files API.
 */
export async function uploadFileToPlatform(
  file: File,
  purpose = 'assistants',
): Promise<{ fileId: string }> {
  const formData = new FormData();
  formData.append('file', file);
  formData.append('purpose', purpose);

  const res = await fetchWithTimeoutAndRetry(`${BACKEND_URL}/api/files`, {
    method: 'POST',
    headers: { ...(await getGlobalHeaders()) },
    body: formData,
  });

  if (!res.ok) {
    const err = await res.text();
    throw new Error(`File platform upload error ${res.status}: ${err}`);
  }

  return res.json();
}

export async function submitLogs(entries: unknown[]): Promise<void> {
  try {
    const res = await fetchWithTimeoutAndRetry(`${BACKEND_URL}/api/logs`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', ...(await getGlobalHeaders()) },
      body: JSON.stringify({ entries }),
    });
    if (!res.ok) {
      logService.originalConsole.warn('[KO] Failed to submit logs:', res.status);
    }
  } catch {
    // Silent: log submission failure should never break the UI
  }
}

export async function extractChartData(params: ChartExtractParams): Promise<ChartExtractResult> {
  const res = await fetchWithTimeoutAndRetry(`${BACKEND_URL}/api/chart-extract`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', ...(await getGlobalHeaders()) },
    body: JSON.stringify(params),
  });

  if (!res.ok) {
    const err = await res.text();
    logService.error('Chart extraction error', undefined, { status: res.status, error: err });
    throw new Error(`Chart extraction error ${res.status}: ${err}`);
  }

  const reqId = res.headers.get('x-request-id');
  if (reqId) logService.info(`Request correlated: ${reqId}`, 'system', { reqId });

  return res.json();
}

export async function searchIconify(query: string, limit = 10, prefix?: string): Promise<any> {
  const params = new URLSearchParams({ query, limit: String(limit) });
  if (prefix) params.set('prefix', prefix);
  const res = await fetchWithTimeoutAndRetry(`${BACKEND_URL}/api/icons/search?${params}`, {
    headers: await getGlobalHeaders(),
  });
  if (!res.ok) throw new Error(`Icon search failed: ${res.status}`);
  return res.json();
}

export async function fetchIconSvg(prefix: string, name: string, color?: string): Promise<string> {
  const params = color ? new URLSearchParams({ color }) : undefined;
  const url = `${BACKEND_URL}/api/icons/svg/${encodeURIComponent(prefix)}/${encodeURIComponent(name)}${params ? '?' + params : ''}`;
  const res = await fetchWithTimeoutAndRetry(url, {
    headers: await getGlobalHeaders(),
  });
  if (!res.ok) throw new Error(`Icon SVG fetch failed: ${res.status}`);
  return res.text();
}

export async function submitFeedback(
  sessionId: string,
  payload: {
    category: string;
    comment: string;
    logs: unknown[];
    chatHistory?: unknown[];
    systemContext?: FeedbackSystemContext;
  },
): Promise<{ success: boolean }> {
  const res = await fetchWithTimeoutAndRetry(`${BACKEND_URL}/api/feedback/${sessionId}`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      ...(await getGlobalHeaders()),
    },
    body: JSON.stringify(payload),
  });

  if (!res.ok) {
    const errText = await res.text();
    throw new Error(`Feedback submission failed: ${res.status} ${errText}`);
  }

  return res.json();
}
