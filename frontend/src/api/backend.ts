const BACKEND_URL = import.meta.env.VITE_BACKEND_URL || 'http://192.168.50.10:3003'

export async function fetchModels(): Promise<Record<string, ModelInfo>> {
  const res = await fetch(`${BACKEND_URL}/api/models`)
  if (!res.ok) throw new Error(`Failed to fetch models: ${res.status}`)
  return res.json()
}

export async function healthCheck(): Promise<boolean> {
  try {
    const res = await fetch(`${BACKEND_URL}/health`)
    return res.ok
  } catch {
    return false
  }
}

export interface ChatMessage {
  role: 'system' | 'user' | 'assistant'
  content: string
}

export interface ChatStreamOptions {
  messages: ChatMessage[]
  modelTier: ModelTier
  onStream: (text: string) => void
  abortSignal?: AbortSignal
}

export async function chatStream(options: ChatStreamOptions): Promise<void> {
  const { messages, modelTier, onStream, abortSignal } = options

  const res = await fetch(`${BACKEND_URL}/api/chat`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ messages, modelTier }),
    signal: abortSignal,
  })

  if (!res.ok) {
    const err = await res.text()
    throw new Error(`Chat API error ${res.status}: ${err}`)
  }

  const reader = res.body!.getReader()
  const decoder = new TextDecoder()
  let fullContent = ''

  while (true) {
    const { done, value } = await reader.read()
    if (done) break

    const chunk = decoder.decode(value, { stream: true })
    const lines = chunk.split('\n')

    for (const line of lines) {
      if (!line.startsWith('data: ')) continue
      const data = line.slice(6)
      if (data === '[DONE]') return

      try {
        const parsed = JSON.parse(data)
        const delta = parsed.choices?.[0]?.delta?.content
        if (delta) {
          fullContent += delta
          onStream(fullContent)
        }
      } catch {
        // Skip malformed JSON lines
      }
    }
  }
}

export interface ChatSyncOptions {
  messages: ChatMessage[]
  modelTier: ModelTier
  tools?: any[]
}

export async function chatSync(options: ChatSyncOptions): Promise<any> {
  const { messages, modelTier, tools } = options

  const res = await fetch(`${BACKEND_URL}/api/chat/sync`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ messages, modelTier, tools }),
  })

  if (!res.ok) {
    const err = await res.text()
    throw new Error(`Chat sync API error ${res.status}: ${err}`)
  }

  return res.json()
}

export interface ImageGenerateOptions {
  prompt: string
  size?: string
  quality?: string
}

export async function generateImage(options: ImageGenerateOptions): Promise<string> {
  const res = await fetch(`${BACKEND_URL}/api/image`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
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
