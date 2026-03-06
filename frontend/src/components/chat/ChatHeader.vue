<template>
  <div class="flex justify-between rounded-sm border border-[#33ABC6]/20 bg-surface/90 p-1">
    <div class="flex flex-1 items-center gap-2 text-accent">
      <img
        src="/Logo.png"
        alt="KickAI logo"
        class="h-8 w-8 rounded-sm border border-black/10 bg-white object-contain p-0.5"
      />
      <div class="flex flex-col leading-none">
        <span class="text-sm font-semibold text-main">{{ t('appTitle') }}</span>
        <span class="text-[10px] text-[#33ABC6]">{{ t('appSubtitle') }}</span>
      </div>
    </div>

    <!-- Session switcher dropdown -->
    <div class="relative flex items-center" ref="dropdownRef">
      <button
        type="button"
        class="flex flex-1 items-center gap-1 px-2 py-1 rounded text-xs text-secondary hover:bg-bg-tertiary transition-colors max-w-[200px]"
        :disabled="loading"
        :title="currentSessionName"
        @click="dropdownOpen = !dropdownOpen"
      >
        <MessageSquare :size="11" class="shrink-0 text-accent" />
        <span class="truncate">{{ currentSessionName }}</span>
        <ChevronDown
          :size="11"
          :class="dropdownOpen ? 'rotate-180' : ''"
          class="shrink-0 transition-transform"
        />
      </button>

      <!-- Dropdown panel -->
      <div
        v-if="dropdownOpen"
        class="absolute top-full left-0 mt-1 w-56 bg-surface border border-border-secondary rounded shadow-lg z-50 overflow-hidden"
        style="font-size: 11px"
      >
        <!-- New Chat -->
        <button
          type="button"
          class="w-full flex items-center gap-2 px-3 py-2 text-xs border-b border-border-secondary transition-colors"
          :class="
            loading ? 'text-secondary cursor-not-allowed' : 'text-accent hover:bg-bg-tertiary'
          "
          :disabled="loading"
          @click="handleNewChat"
        >
          <Plus :size="13" />
          {{ t('newChat', 'New Chat') }}
        </button>

        <!-- Session list -->
        <div class="max-h-48 overflow-y-auto">
          <button
            v-for="session in sessions"
            :key="session.id"
            type="button"
            class="flex items-center justify-between px-3 py-2 text-xs w-full text-left transition-colors"
            :class="[
              session.id === currentSessionId ? 'bg-bg-tertiary' : 'hover:bg-bg-tertiary',
              loading && session.id !== currentSessionId
                ? 'opacity-50 cursor-not-allowed'
                : 'cursor-pointer',
            ]"
            :disabled="loading && session.id !== currentSessionId"
            @click="handleSwitchSession(session.id)"
          >
            <div class="flex items-center gap-2 min-w-0 flex-1">
              <Check
                v-if="session.id === currentSessionId"
                :size="11"
                class="text-accent shrink-0"
              />
              <div v-else class="w-3 shrink-0" />
              <span class="truncate text-main">{{ session.name }}</span>
            </div>
            <span class="text-[10px] text-secondary shrink-0 ml-2">{{
              getSessionMessageCount(session)
            }}</span>
          </button>
        </div>

        <!-- Delete current -->
        <button
          v-if="sessions.length > 1"
          type="button"
          class="w-full flex items-center gap-2 px-3 py-2 text-xs border-t border-border-secondary transition-colors"
          :class="
            loading ? 'text-secondary cursor-not-allowed' : 'text-red-500 hover:bg-bg-tertiary'
          "
          :disabled="loading"
          @click="handleDeleteSession"
        >
          <Trash2 :size="13" />
          {{ t('deleteSession', 'Delete session') }}
        </button>
      </div>
    </div>

    <div class="flex items-center gap-1 rounded-md border border-accent/10">
      <CustomButton
        :title="settingsTitle"
        :icon="Settings"
        text=""
        type="secondary"
        class="border-none p-1!"
        :icon-size="18"
        @click="$emit('settings')"
      />
    </div>
  </div>
</template>

<script lang="ts" setup>
import { ref, computed, onMounted, onBeforeUnmount } from 'vue'
import { useI18n } from 'vue-i18n'
import { Check, ChevronDown, MessageSquare, Plus, Settings, Trash2 } from 'lucide-vue-next'

import CustomButton from '@/components/CustomButton.vue'
import type { ChatSession } from '@/composables/useSessionManager'
import { getSessionMessageCount } from '@/composables/useSessionManager'

const { t } = useI18n()

const props = defineProps<{
  settingsTitle: string
  loading: boolean
  sessions: ChatSession[]
  currentSessionId: string | null
}>()

const emit = defineEmits<{
  (e: 'new-chat'): void
  (e: 'settings'): void
  (e: 'switch-session', sessionId: string): void
  (e: 'delete-session'): void
}>()

const dropdownOpen = ref(false)
const dropdownRef = ref<HTMLElement>()

const currentSessionName = computed(() => {
  const session = props.sessions.find(s => s.id === props.currentSessionId)
  const name = session?.name ?? t('newChat')
  return name.length > 30 ? `${name.slice(0, 28)}…` : name
})

function handleNewChat() {
  dropdownOpen.value = false
  emit('new-chat')
}

function handleSwitchSession(sessionId: string) {
  dropdownOpen.value = false
  emit('switch-session', sessionId)
}

function handleDeleteSession() {
  dropdownOpen.value = false
  emit('delete-session')
}

function handleClickOutside(event: MouseEvent) {
  if (dropdownRef.value && !dropdownRef.value.contains(event.target as Node)) {
    dropdownOpen.value = false
  }
}

onMounted(() => {
  document.addEventListener('mousedown', handleClickOutside)
})

onBeforeUnmount(() => {
  document.removeEventListener('mousedown', handleClickOutside)
})
</script>
