<template>
  <div class="itemse-center relative flex h-full w-full flex-col justify-center bg-bg-secondary p-1">
    <div class="relative flex h-full w-full flex-col gap-1 rounded-md">
      <!-- Header -->
      <div class="flex justify-between rounded-sm border border-[#33ABC6]/20 bg-surface/90 p-1">
        <div class="flex flex-1 items-center gap-2 text-accent">
          <img src="/Logo.png" alt="KickAI logo" class="h-8 w-8 rounded-sm border border-black/10 bg-white object-contain p-0.5" />
          <div class="flex flex-col leading-none">
            <span class="text-sm font-semibold text-main">KickOffice</span>
            <span class="text-[10px] text-[#33ABC6]">AI Office Assistant</span>
          </div>
        </div>
        <div class="mr-1 flex items-center gap-1">
          <span class="h-2 w-2 rounded-full bg-[#33ABC6]" />
          <span class="h-2 w-2 rounded-full bg-black" />
          <span class="h-2 w-2 rounded-full border border-black/30 bg-white" />
        </div>
        <div class="flex items-center gap-1 rounded-md border border-accent/10">
          <CustomButton
            :title="t('newChat')"
            :icon="Plus"
            text=""
            type="secondary"
            class="border-none p-1!"
            :icon-size="18"
            @click="startNewChat"
          />
          <CustomButton
            :title="t('settings')"
            :icon="Settings"
            text=""
            type="secondary"
            class="border-none p-1!"
            :icon-size="18"
            @click="goToSettings"
          />
        </div>
      </div>

      <!-- Quick Actions Bar -->
      <div class="flex w-full items-center justify-center gap-2 overflow-hidden rounded-md">
        <CustomButton
          v-for="action in quickActions"
          :key="action.key"
          :title="action.label"
          text=""
          :icon="action.icon"
          type="secondary"
          :icon-size="16"
          class="shrink-0! bg-surface! p-1.5!"
          :disabled="loading"
          @click="applyQuickAction(action.key)"
        />
        <SingleSelect
          v-model="selectedPromptId"
          :key-list="savedPrompts.map(prompt => prompt.id)"
          :placeholder="t('selectPrompt')"
          title=""
          :fronticon="false"
          class="max-w-xs! flex-1! bg-surface! text-xs!"
          @change="loadSelectedPrompt"
        >
          <template #item="{ item }">
            {{ savedPrompts.find(prompt => prompt.id === item)?.name || item }}
          </template>
        </SingleSelect>
      </div>

      <!-- Chat Messages Container -->
      <div
        ref="messagesContainer"
        class="flex flex-1 flex-col gap-4 overflow-y-auto rounded-md border border-border-secondary bg-surface p-2 shadow-sm"
      >
        <div
          v-if="history.length === 0"
          class="flex h-full flex-col items-center justify-center gap-4 p-8 text-center text-accent"
        >
          <Sparkles :size="32" />
          <p class="font-semibold text-main">
            {{ $t('emptyTitle') }}
          </p>
          <p class="text-xs font-semibold text-secondary">
            {{ $t(hostIsExcel ? 'emptySubtitleExcel' : 'emptySubtitle') }}
          </p>
          <!-- Backend status -->
          <div
            class="flex items-center gap-1 rounded-md px-2 py-1 text-xs"
            :class="backendOnline ? 'bg-green-100 text-green-700' : 'bg-red-100 text-red-700'"
          >
            <div
              class="h-2 w-2 rounded-full"
              :class="backendOnline ? 'bg-green-500' : 'bg-red-500'"
            />
            {{ backendOnline ? t('backendOnline') : t('backendOffline') }}
          </div>
        </div>

        <div
          v-for="(msg, index) in history"
          :key="index"
          class="group flex items-end gap-4 [.user]:flex-row-reverse"
          :class="msg.role === 'assistant' ? 'assistant' : 'user'"
        >
          <div
            class="flex min-w-0 flex-1 flex-col gap-1 group-[.assistant]:items-start group-[.assistant]:text-left group-[.user]:items-end group-[.user]:text-left"
          >
            <div
              class="group max-w-[95%] rounded-md border border-border-secondary p-1 text-sm leading-[1.4] wrap-break-word whitespace-pre-wrap text-main/90 shadow-sm group-[.assistant]:bg-bg-tertiary group-[.assistant]:text-left group-[.user]:bg-accent/10"
            >
              <template v-for="(segment, idx) in renderSegments(msg.content)" :key="idx">
                <span v-if="segment.type === 'text'">{{ segment.text.trim() }}</span>
                <details v-else class="mb-1 rounded-sm border border-border-secondary bg-bg-secondary">
                  <summary class="cursor-pointer list-none p-1 text-sm font-semibold text-secondary">
                    Thought process
                  </summary>
                  <pre class="m-0 p-1 text-xs wrap-break-word whitespace-pre-wrap text-secondary">{{
                    segment.text.trim()
                  }}</pre>
                </details>
              </template>
              <!-- Image display -->
              <img
                v-if="msg.imageSrc"
                :src="msg.imageSrc"
                class="mt-2 max-w-full rounded-md"
                alt="Generated image"
              />
            </div>
            <div v-if="msg.role === 'assistant'" class="flex gap-1">
              <CustomButton
                :title="t('replaceSelectedText')"
                text=""
                :icon="FileText"
                type="secondary"
                class="bg-surface! p-1.5! text-secondary!"
                :icon-size="12"
                @click="insertToDocument(getMessageActionPayload(msg), 'replace')"
              />
              <CustomButton
                :title="t('appendToSelection')"
                text=""
                :icon="Plus"
                type="secondary"
                class="bg-surface! p-1.5! text-secondary!"
                :icon-size="12"
                @click="insertToDocument(getMessageActionPayload(msg), 'append')"
              />
              <CustomButton
                :title="t('copyToClipboard')"
                text=""
                :icon="Copy"
                type="secondary"
                class="bg-surface! p-1.5! text-secondary!"
                :icon-size="12"
                @click="copyToClipboard(getMessageActionPayload(msg))"
              />
            </div>
          </div>
        </div>
      </div>

      <!-- Input Area -->
      <div class="flex flex-col gap-1 rounded-md">
        <div class="flex items-center justify-between gap-2 overflow-hidden">
          <div class="flex shrink-0 gap-1 rounded-sm border border-border bg-surface p-0.5">
            <button
              class="cursor-po flex h-7 w-7 items-center justify-center rounded-md border-none text-secondary hover:bg-accent/30 hover:text-white! [.active]:text-accent"
              :class="{ active: mode === 'ask' }"
              :title="t('askMode')"
              @click="mode = 'ask'"
            >
              <MessageSquare :size="14" />
            </button>
            <button
              class="cursor-po flex h-7 w-7 items-center justify-center rounded-md border-none text-secondary hover:bg-accent/30 hover:text-white! [.active]:text-accent"
              :class="{ active: mode === 'agent' }"
              :title="t('agentMode')"
              @click="mode = 'agent'"
            >
              <BotMessageSquare :size="17" />
            </button>
            <button
              class="cursor-po flex h-7 w-7 items-center justify-center rounded-md border-none text-secondary hover:bg-accent/30 hover:text-white! [.active]:text-accent"
              :class="{ active: mode === 'image' }"
              :title="t('imageMode')"
              @click="mode = 'image'"
            >
              <ImageIcon :size="14" />
            </button>
          </div>
          <div class="flex min-w-0 flex-1 gap-1 overflow-hidden">
            <select
              v-if="mode !== 'image'"
              v-model="selectedModelTier"
              class="h-7 max-w-full min-w-0 cursor-pointer rounded-md border border-border bg-surface p-1 text-xs text-secondary hover:border-accent focus:outline-none"
            >
              <option v-for="(info, tier) in chatModels" :key="tier" :value="tier">
                {{ info.label }}
              </option>
            </select>
            <span
              v-else
              class="flex h-7 items-center rounded-md border border-border bg-surface px-2 text-xs text-secondary"
            >
              {{ availableModels.image?.id || 'image' }}
            </span>
          </div>
        </div>
        <div
          class="flex min-w-12 items-center gap-2 rounded-md border border-border bg-surface p-2 focus-within:border-accent"
        >
          <textarea
            ref="inputTextarea"
            v-model="userInput"
            class="placeholder::text-secondary block max-h-30 flex-1 resize-none overflow-y-auto border-none bg-transparent py-2 text-xs leading-normal text-main outline-none placeholder:text-xs"
            :placeholder="mode === 'image' ? $t('describeImage') : mode === 'agent' ? $t('directTheAgent') : $t('askAnything')"
            rows="1"
            @keydown.enter.exact.prevent="sendMessage"
            @input="adjustTextareaHeight"
          />
          <button
            v-if="loading"
            class="flex h-7 w-7 shrink-0 cursor-pointer items-center justify-center rounded-sm border-none bg-danger text-white"
            title="Stop"
            @click="stopGeneration"
          >
            <Square :size="18" />
          </button>
          <button
            v-else
            class="flex h-7 w-7 shrink-0 cursor-pointer items-center justify-center rounded-sm border-none bg-accent text-white disabled:cursor-not-allowed disabled:bg-accent/50"
            title="Send"
            :disabled="!userInput.trim() || !backendOnline"
            @click="sendMessage"
          >
            <Send :size="18" />
          </button>
        </div>
        <div class="flex justify-center gap-3 px-1">
          <label v-if="!hostIsExcel" class="flex h-3.5 w-3.5 flex-1 cursor-pointer items-center gap-1 text-xs text-secondary">
            <input v-model="useWordFormatting" type="checkbox" />
            <span>{{ $t('useWordFormattingLabel') }}</span>
          </label>
          <label class="flex h-3.5 w-3.5 flex-1 cursor-pointer items-center gap-1 text-xs text-secondary">
            <input v-model="useSelectedText" type="checkbox" />
            <span>{{ $t(hostIsExcel ? 'includeSelectionLabelExcel' : 'includeSelectionLabel') }}</span>
          </label>
        </div>
      </div>
    </div>
  </div>
</template>

<script lang="ts" setup>
import { useStorage } from '@vueuse/core'
import {
  BarChart3,
  BookOpen,
  BotMessageSquare,
  Calculator,
  CheckCircle,
  Copy,
  FileCheck,
  FileSpreadsheet,
  FileText,
  Globe,
  ImageIcon,
  LineChart,
  MessageSquare,
  Plus,
  Send,
  Settings,
  Sparkle,
  Sparkles,
  Square,
  TableProperties,
} from 'lucide-vue-next'
import { computed, nextTick, onBeforeMount, ref } from 'vue'
import { useI18n } from 'vue-i18n'
import { useRouter } from 'vue-router'

import { insertFormattedResult, insertResult } from '@/api/common'
import { type ChatMessage, chatStream, chatSync, fetchModels, generateImage, healthCheck } from '@/api/backend'
import CustomButton from '@/components/CustomButton.vue'
import SingleSelect from '@/components/SingleSelect.vue'
import { buildInPrompt, excelBuiltInPrompt, getBuiltInPrompt, getExcelBuiltInPrompt } from '@/utils/constant'
import { localStorageKey } from '@/utils/enum'
import { getExcelToolDefinitions } from '@/utils/excelTools'
import { getGeneralToolDefinitions } from '@/utils/generalTools'
import { isExcel, isWord } from '@/utils/hostDetection'
import { message as messageUtil } from '@/utils/message'
import { getWordToolDefinitions } from '@/utils/wordTools'

const router = useRouter()
const { t } = useI18n()

interface DisplayMessage {
  role: 'user' | 'assistant' | 'system'
  content: string
  imageSrc?: string
}

interface SavedPrompt {
  id: string
  name: string
  systemPrompt: string
  userPrompt: string
}

const savedPrompts = ref<SavedPrompt[]>([])
const selectedPromptId = ref<string>('')
const customSystemPrompt = ref<string>('')

// Backend state
const backendOnline = ref(false)
const availableModels = ref<Record<string, ModelInfo>>({})

// Chat state
const mode = useStorage(localStorageKey.chatMode, 'ask' as 'ask' | 'agent' | 'image')
const selectedModelTier = useStorage<ModelTier>(localStorageKey.modelTier, 'standard')
const history = ref<DisplayMessage[]>([])
const userInput = ref('')
const loading = ref(false)
const imageLoading = ref(false)

// Filter out image model from chat selector
const chatModels = ref<Record<string, ModelInfo>>({})
const messagesContainer = ref<HTMLElement>()
const inputTextarea = ref<HTMLTextAreaElement>()
const abortController = ref<AbortController | null>(null)

// Settings
const useWordFormatting = useStorage(localStorageKey.useWordFormatting, true)
const useSelectedText = useStorage(localStorageKey.useSelectedText, true)
const replyLanguage = useStorage(localStorageKey.replyLanguage, 'Fran\u00e7ais')
const agentMaxIterations = useStorage(localStorageKey.agentMaxIterations, 25)
const userGender = useStorage(localStorageKey.userGender, 'unspecified')
const userFirstName = useStorage(localStorageKey.userFirstName, '')
const userLastName = useStorage(localStorageKey.userLastName, '')
const excelFormulaLanguage = useStorage<'en' | 'fr'>(localStorageKey.excelFormulaLanguage, 'en')
const insertType = ref<insertTypes>('replace')

// Host detection
const hostIsExcel = isExcel()
const hostIsWord = isWord()

// Quick actions - different for Word vs Excel
const wordQuickActions: {
  key: string
  label: string
  icon: any
}[] = [
  { key: 'translate', label: t('translate'), icon: Globe },
  { key: 'polish', label: t('polish'), icon: Sparkle },
  { key: 'academic', label: t('academic'), icon: BookOpen },
  { key: 'summary', label: t('summary'), icon: FileCheck },
  { key: 'grammar', label: t('grammar'), icon: CheckCircle },
]

const excelQuickActions: {
  key: string
  label: string
  icon: any
}[] = [
  { key: 'analyze', label: t('excelAnalyze'), icon: BarChart3 },
  { key: 'chart', label: t('excelChart'), icon: LineChart },
  { key: 'formula', label: t('excelFormula'), icon: Calculator },
  { key: 'format', label: t('excelFormat'), icon: TableProperties },
  { key: 'explain', label: t('excelExplain'), icon: FileSpreadsheet },
]

const quickActions = computed(() => hostIsExcel ? excelQuickActions : wordQuickActions)

// Think tag parsing
const THINK_TAG = '<think>'
const THINK_TAG_END = '</think>'

interface RenderSegment {
  type: 'text' | 'think'
  text: string
}

function splitThinkSegments(text: string): RenderSegment[] {
  if (!text) return []
  const segments: RenderSegment[] = []
  let cursor = 0

  while (cursor < text.length) {
    const start = text.indexOf(THINK_TAG, cursor)
    if (start === -1) {
      segments.push({ type: 'text', text: text.slice(cursor) })
      break
    }
    if (start > cursor) {
      segments.push({ type: 'text', text: text.slice(cursor, start) })
    }
    const end = text.indexOf(THINK_TAG_END, start + THINK_TAG.length)
    if (end === -1) {
      segments.push({ type: 'think', text: text.slice(start + THINK_TAG.length) })
      break
    }
    segments.push({ type: 'think', text: text.slice(start + THINK_TAG.length, end) })
    cursor = end + THINK_TAG_END.length
  }
  return segments.filter(s => s.text)
}

function renderSegments(content: string): RenderSegment[] {
  return splitThinkSegments(content)
}

function cleanContent(content: string): string {
  const regex = new RegExp(`${THINK_TAG}[\\s\\S]*?${THINK_TAG_END}`, 'g')
  return content.replace(regex, '').trim()
}

function getMessageActionPayload(message: DisplayMessage): string {
  const cleanedText = cleanContent(message.content)
  if (cleanedText) {
    return cleanedText
  }
  return message.imageSrc || ''
}


function loadSavedPrompts() {
  const stored = localStorage.getItem('savedPrompts')
  if (stored) {
    try {
      savedPrompts.value = JSON.parse(stored)
    } catch {
      savedPrompts.value = []
    }
  }
}

function loadSelectedPrompt() {
  if (!selectedPromptId.value) {
    customSystemPrompt.value = ''
    return
  }
  const prompt = savedPrompts.value.find(p => p.id === selectedPromptId.value)
  if (prompt) {
    customSystemPrompt.value = prompt.systemPrompt
    userInput.value = prompt.userPrompt
    adjustTextareaHeight()
    if (inputTextarea.value) {
      inputTextarea.value.focus()
    }
  }
}

function goToSettings() {
  router.push('/settings')
}

function startNewChat() {
  if (loading.value) stopGeneration()
  userInput.value = ''
  history.value = []
  customSystemPrompt.value = ''
  selectedPromptId.value = ''
  adjustTextareaHeight()
}

function stopGeneration() {
  if (abortController.value) {
    abortController.value.abort()
    abortController.value = null
  }
  loading.value = false
}

function adjustTextareaHeight() {
  if (inputTextarea.value) {
    inputTextarea.value.style.height = 'auto'
    inputTextarea.value.style.height = Math.min(inputTextarea.value.scrollHeight, 120) + 'px'
  }
}

async function scrollToBottom() {
  await nextTick()
  if (messagesContainer.value) {
    messagesContainer.value.scrollTop = messagesContainer.value.scrollHeight
  }
}

function buildChatMessages(systemPrompt: string): ChatMessage[] {
  const messages: ChatMessage[] = [
    { role: 'system', content: systemPrompt },
  ]
  for (const msg of history.value) {
    if (msg.role === 'user' || msg.role === 'assistant') {
      messages.push({ role: msg.role, content: msg.content })
    }
  }
  return messages
}

const wordAgentPrompt = (lang: string) =>
  `
# Role
You are a highly skilled Microsoft Word Expert Agent. Your goal is to assist users in creating, editing, and formatting documents with professional precision.

# Capabilities
- You can interact with the document directly using provided tools (reading text, applying styles, inserting content, etc.).
- You understand document structure, typography, and professional writing standards.

# Guidelines
1. **Tool First**: If a request requires document modification or inspection, prioritize using the available tools.
2. **Direct Actions**: For Word formatting requests (bold, color, tags like <b_red>...</b_red>, etc.), execute the change directly with tools instead of giving manual steps.
3. **Accuracy**: Ensure formatting and content changes are precise and follow the user's intent.
4. **Conciseness**: Provide brief, helpful explanations of your actions.
5. **Language**: You must communicate entirely in ${lang}.

# Safety
Do not perform destructive actions (like clearing the whole document) unless explicitly instructed.
`.trim()

const excelFormulaLanguageInstruction = () =>
  excelFormulaLanguage.value === 'fr'
    ? 'Excel interface locale: French. Use localized French function names and separators when providing formulas, and prefer localized formula tool behavior.'
    : 'Excel interface locale: English. Use English function names and standard English formula syntax.'

const excelAgentPrompt = (lang: string) =>
  `
# Role
You are a highly skilled Microsoft Excel Expert Agent. Your goal is to assist users with data analysis, formulas, charts, formatting, and spreadsheet operations with professional precision.

# Capabilities
- You can interact with the spreadsheet directly using provided tools (reading cells, writing values, inserting formulas, creating charts, formatting ranges, etc.).
- You understand data analysis, statistical methods, Excel formulas, and data visualization best practices.
- You can perform operations like sorting, filtering, formatting, and creating charts.

# Guidelines
1. **Tool First**: If a request requires spreadsheet modification or data reading, prioritize using the available tools.
2. **Read First**: Before modifying data, read the current state to understand the structure.
3. **Accuracy**: Ensure formulas, formatting, and data operations are precise and follow the user's intent.
4. **Conciseness**: Provide brief, helpful explanations of your actions and results.
5. **Language**: You must communicate entirely in ${lang}.
6. **Formula locale**: ${excelFormulaLanguageInstruction()}

# Safety
Do not perform destructive actions (like clearing all data or deleting sheets) unless explicitly instructed.
`.trim()


function userProfilePromptBlock() {
  const firstName = userFirstName.value.trim()
  const lastName = userLastName.value.trim()
  const fullName = `${firstName} ${lastName}`.trim() || t('userProfileUnknownName')

  const genderMap: Record<string, string> = {
    female: t('userGenderFemale'),
    male: t('userGenderMale'),
    nonbinary: t('userGenderNonBinary'),
    unspecified: t('userGenderUnspecified'),
  }

  const genderLabel = genderMap[userGender.value] || t('userGenderUnspecified')

  return `

User profile context for communications (especially emails):
- First name: ${firstName || t('userProfileUnknownFirstName')}
- Last name: ${lastName || t('userProfileUnknownLastName')}
- Full name: ${fullName}
- Gender: ${genderLabel}
Use this profile when drafting salutations, signatures, and tone, unless the user asks otherwise.`
}

const agentPrompt = (lang: string) =>
  `${hostIsExcel ? excelAgentPrompt(lang) : wordAgentPrompt(lang)}${userProfilePromptBlock()}`

const wordStandardPrompt = (lang: string) =>
  `You are a helpful Microsoft Word specialist. Help users with drafting, brainstorming, and Word-related questions. Reply in ${lang}.`

const excelStandardPrompt = (lang: string) =>
  `You are a helpful Microsoft Excel specialist. Help users with data analysis, formulas, charts, formatting, and spreadsheet-related questions. Reply in ${lang}. ${excelFormulaLanguageInstruction()}`

const standardPrompt = (lang: string) =>
  `${hostIsExcel ? excelStandardPrompt(lang) : wordStandardPrompt(lang)}${userProfilePromptBlock()}`

async function sendMessage() {
  if (!userInput.value.trim() || loading.value) return
  if (!backendOnline.value) {
    messageUtil.error(t('backendOffline'))
    return
  }

  const userMessage = userInput.value.trim()
  userInput.value = ''
  adjustTextareaHeight()

  // Get selected content from the active Office app
  let selectedText = ''
  if (useSelectedText.value) {
    try {
      if (hostIsExcel) {
        selectedText = await Excel.run(async (ctx) => {
          const range = ctx.workbook.getSelectedRange()
          range.load('values, address')
          await ctx.sync()
          const values = range.values
          const formatted = values.map((row: any[]) => row.join('\t')).join('\n')
          return `[${range.address}]\n${formatted}`
        })
      } else {
        selectedText = await Word.run(async ctx => {
          const range = ctx.document.getSelection()
          range.load('text')
          await ctx.sync()
          return range.text
        })
      }
    } catch {
      // Not in Office context
    }
  }

  const selectionLabel = hostIsExcel ? 'Selected cells' : 'Selected text'
  const fullMessage = selectedText
    ? `${userMessage}\n\n[${selectionLabel}: "${selectedText}"]`
    : userMessage

  history.value.push({ role: 'user', content: fullMessage })
  scrollToBottom()

  loading.value = true
  abortController.value = new AbortController()

  try {
    await processChat(fullMessage)
  } catch (error: any) {
    if (error.name === 'AbortError') {
      messageUtil.info(t('generationStop'))
    } else {
      console.error(error)
      messageUtil.error(t('failedToResponse'))
      if (history.value.length > 0 && history.value[history.value.length - 1].role === 'assistant') {
        history.value.pop()
      }
    }
  } finally {
    loading.value = false
    abortController.value = null
  }
}

async function processChat(userMessage: string) {
  const lang = replyLanguage.value || 'Fran\u00e7ais'

  // Image generation mode
  if (mode.value === 'image') {
    history.value.push({ role: 'assistant', content: t('imageGenerating') })
    scrollToBottom()
    imageLoading.value = true
    try {
      const imageSrc = await generateImage({ prompt: userMessage })
      if (!imageSrc) {
        throw new Error('Image API returned no image payload (expected b64_json or url).')
      }
      const lastIndex = history.value.length - 1
      history.value[lastIndex] = {
        role: 'assistant',
        content: '',
        imageSrc,
      }
    } catch (err: any) {
      const lastIndex = history.value.length - 1
      history.value[lastIndex] = {
        role: 'assistant',
        content: `${t('imageError')}: ${err.message}`,
      }
    } finally {
      imageLoading.value = false
    }
    scrollToBottom()
    return
  }

  const isAgentMode = mode.value === 'agent'

  const systemPrompt =
    customSystemPrompt.value || (isAgentMode ? agentPrompt(lang) : standardPrompt(lang))

  const messages = buildChatMessages(systemPrompt)

  // Add placeholder for assistant response
  history.value.push({ role: 'assistant', content: '' })

  if (isAgentMode) {
    await runAgentLoop(messages, systemPrompt)
  } else {
    await chatStream({
      messages,
      modelTier: selectedModelTier.value,
      abortSignal: abortController.value?.signal,
      onStream: (text: string) => {
        const lastIndex = history.value.length - 1
        history.value[lastIndex] = { role: 'assistant', content: text }
        scrollToBottom()
      },
    })
  }

  scrollToBottom()
}

async function runAgentLoop(messages: ChatMessage[], _systemPrompt: string) {
  const appToolDefs = hostIsExcel ? getExcelToolDefinitions() : getWordToolDefinitions()
  const generalToolDefs = getGeneralToolDefinitions()

  // Build OpenAI-format tool definitions
  const tools = [...generalToolDefs, ...appToolDefs].map(def => ({
    type: 'function' as const,
    function: {
      name: def.name,
      description: def.description,
      parameters: def.inputSchema,
    },
  }))

  let iteration = 0
  const maxIter = Number(agentMaxIterations.value) || 25
  let currentMessages = [...messages]

  while (iteration < maxIter) {
    iteration++

    const response = await chatSync({
      messages: currentMessages,
      modelTier: selectedModelTier.value,
      tools,
    })

    const choice = response.choices?.[0]
    if (!choice) break

    const assistantMsg = choice.message
    currentMessages.push(assistantMsg)

    // If no tool calls, we're done
    if (!assistantMsg.tool_calls || assistantMsg.tool_calls.length === 0) {
      const lastIndex = history.value.length - 1
      history.value[lastIndex] = { role: 'assistant', content: assistantMsg.content || '' }
      scrollToBottom()
      break
    }

    // Process tool calls
    for (const toolCall of assistantMsg.tool_calls) {
      const toolName = toolCall.function.name
      let toolArgs: Record<string, any> = {}
      try {
        toolArgs = JSON.parse(toolCall.function.arguments)
      } catch {
        toolArgs = {}
      }

      // Show tool call in UI
      const lastIndex = history.value.length - 1
      const currentContent = history.value[lastIndex].content
      history.value[lastIndex] = {
        role: 'assistant',
        content: currentContent + `\n\nTool: ${toolName}...`,
      }
      scrollToBottom()

      // Execute the tool
      let result = ''
      const allTools = [...generalToolDefs, ...appToolDefs]
      const toolDef = allTools.find(t => t.name === toolName)
      if (toolDef) {
        try {
          result = await toolDef.execute(toolArgs)
        } catch (err: any) {
          result = `Error: ${err.message}`
        }
      } else {
        result = `Unknown tool: ${toolName}`
      }

      // Add tool result to messages
      currentMessages.push({
        role: 'tool' as any,
        tool_call_id: toolCall.id,
        content: result,
      } as any)

      // Update UI
      history.value[lastIndex] = {
        role: 'assistant',
        content: currentContent + `\nTool ${toolName} done.`,
      }
      scrollToBottom()
    }

    // Loop continues: next iteration sends tool results back to LLM
  }

  if (iteration >= maxIter) {
    messageUtil.warning(t('recursionLimitExceeded'))
  }
}

async function applyQuickAction(actionKey: string) {
  if (!backendOnline.value) {
    messageUtil.error(t('backendOffline'))
    return
  }

  let selectedText = ''
  try {
    if (hostIsExcel) {
      selectedText = await Excel.run(async (ctx) => {
        const range = ctx.workbook.getSelectedRange()
        range.load('values, address')
        await ctx.sync()
        const values = range.values
        const formatted = values.map((row: any[]) => row.join('\t')).join('\n')
        return `[${range.address}]\n${formatted}`
      })
    } else {
      selectedText = await Word.run(async ctx => {
        const range = ctx.document.getSelection()
        range.load('text')
        await ctx.sync()
        return range.text
      })
    }
  } catch {
    // Not in Office context
  }

  if (!selectedText) {
    messageUtil.error(t(hostIsExcel ? 'selectCellsPrompt' : 'selectTextPrompt'))
    return
  }

  // Get the right prompt set based on host
  let action: { system: (lang: string) => string; user: (text: string, lang: string) => string }
  if (hostIsExcel) {
    const excelPrompts = getExcelBuiltInPrompt()
    action = excelPrompts[actionKey as keyof typeof excelBuiltInPrompt]
  } else {
    const wordPrompts = getBuiltInPrompt()
    action = wordPrompts[actionKey as keyof typeof buildInPrompt]
  }

  if (!action) return

  const lang = replyLanguage.value || 'Français'
  const systemMsg = action.system(lang)
  const userMsg = action.user(selectedText, lang)

  const displayKey = hostIsExcel ? `excel${actionKey.charAt(0).toUpperCase() + actionKey.slice(1)}` : actionKey
  history.value.push({ role: 'user', content: `[${t(displayKey)}] ${selectedText.substring(0, 100)}...` })
  history.value.push({ role: 'assistant', content: '' })
  scrollToBottom()

  loading.value = true
  abortController.value = new AbortController()

  try {
    const messages: ChatMessage[] = [
      { role: 'system', content: systemMsg },
      { role: 'user', content: userMsg },
    ]

    await chatStream({
      messages,
      modelTier: selectedModelTier.value,
      abortSignal: abortController.value?.signal,
      onStream: (text: string) => {
        const lastIndex = history.value.length - 1
        history.value[lastIndex] = { role: 'assistant', content: text }
        scrollToBottom()
      },
    })
  } catch (error: any) {
    if (error.name === 'AbortError') {
      messageUtil.info(t('generationStop'))
    } else {
      console.error(error)
      messageUtil.error(t('failedToProcessAction'))
      history.value.pop()
    }
  } finally {
    loading.value = false
    abortController.value = null
  }
}

async function insertToDocument(content: string, type: insertTypes) {
  if (!content.trim()) {
    return
  }

  if (hostIsExcel) {
    // For Excel: write content to the selected cell
    try {
      await Excel.run(async (ctx) => {
        const range = ctx.workbook.getSelectedRange()
        range.values = [[content]]
        await ctx.sync()
      })
      messageUtil.success(t('insertedToCell'))
    } catch (err: any) {
      // Fallback to clipboard
      await copyToClipboard(content, true)
    }
    return
  }

  try {
    insertType.value = type
    if (useWordFormatting.value) {
      await insertFormattedResult(content, insertType)
    } else {
      await insertResult(content, insertType)
    }
    messageUtil.success(t('inserted'))
  } catch (err: any) {
    console.warn('Document insertion failed, falling back to clipboard:', err)
    await copyToClipboard(content, true)
  }
}

async function copyToClipboard(text: string, fallback = false) {
  if (!text.trim()) {
    return
  }

  const notifySuccess = () => messageUtil.success(t(fallback ? 'copiedFallback' : 'copied'))

  try {
    await navigator.clipboard.writeText(text)
    notifySuccess()
    return
  } catch (err: any) {
    console.warn('Clipboard API write failed, trying legacy copy fallback:', err)
  }

  try {
    const textarea = document.createElement('textarea')
    textarea.value = text
    textarea.setAttribute('readonly', '')
    textarea.style.position = 'fixed'
    textarea.style.opacity = '0'
    textarea.style.pointerEvents = 'none'
    document.body.appendChild(textarea)
    textarea.select()
    textarea.setSelectionRange(0, text.length)
    const copied = document.execCommand('copy')
    document.body.removeChild(textarea)

    if (copied) {
      notifySuccess()
    } else {
      messageUtil.error(t('failedToInsert'))
    }
  } catch (err: any) {
    console.error('Legacy clipboard copy failed:', err)
    messageUtil.error(t('failedToInsert'))
  }
}

async function checkBackend() {
  backendOnline.value = await healthCheck()
  if (backendOnline.value) {
    try {
      availableModels.value = await fetchModels()
      // Filter out image model for chat selector
      const filtered: Record<string, ModelInfo> = {}
      for (const [tier, info] of Object.entries(availableModels.value)) {
        if (info.type !== 'image') filtered[tier] = info
      }
      chatModels.value = filtered
    } catch {
      console.error('Failed to fetch models')
    }
  }
}

onBeforeMount(() => {
  insertType.value = (localStorage.getItem(localStorageKey.insertType) as insertTypes) || 'replace'
  loadSavedPrompts()
  checkBackend()
  // Re-check backend every 30 seconds
  setInterval(checkBackend, 30000)
})
</script>
