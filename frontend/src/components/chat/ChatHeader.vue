<template>
  <div class="flex justify-between rounded-sm border border-accent/20 bg-surface/90 p-1">
    <div class="flex flex-1 items-center gap-2 text-accent">
      <img
        src="/Logo.png"
        alt="KickAI logo"
        class="h-8 w-8 rounded-sm border border-black/10 bg-white object-contain p-0.5"
      />
      <div class="flex flex-col leading-none">
        <span class="text-sm font-semibold text-main">{{ t('appTitle') }}</span>
        <span class="text-[10px] text-accent">{{ t('appSubtitle') }}</span>
      </div>
    </div>

    <!-- Session switcher dropdown -->
    <div class="relative flex items-center" ref="dropdownRef">
      <button
        type="button"
        class="flex flex-1 items-center gap-1 px-1.5 py-1 rounded text-xs text-secondary cursor-pointer hover:bg-bg-tertiary transition-colors duration-fast max-w-[140px]"
        :disabled="loading"
        :title="currentSessionName"
        :aria-label="t('sessionOptions', 'Session Options')"
        :aria-expanded="dropdownOpen"
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
      <Transition
        enter-active-class="transition-all duration-fast ease-apple"
        leave-active-class="transition-all duration-fast ease-apple"
        enter-from-class="opacity-0 -translate-y-1"
        leave-to-class="opacity-0 -translate-y-1"
      >
        <div
          v-if="dropdownOpen"
          class="absolute top-full right-0 mt-1 min-w-[220px] max-w-[calc(100vw-1rem)] bg-bg-tertiary border border-border-secondary rounded shadow-lg z-50 overflow-hidden"
          style="font-size: 11px"
        >
          <!-- New Chat -->
          <button
            type="button"
            class="w-full flex items-center gap-2 px-3 py-2 text-xs border-b border-border-secondary transition-colors duration-fast cursor-pointer"
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
              class="flex items-center justify-between px-3 py-2 text-xs w-full text-left transition-colors duration-fast cursor-pointer"
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
            class="w-full flex items-center gap-2 px-3 py-2 text-xs border-t border-border-secondary transition-colors duration-fast cursor-pointer"
            :class="
              loading ? 'text-secondary cursor-not-allowed' : 'text-danger hover:bg-bg-tertiary'
            "
            :disabled="loading"
            @click="handleDeleteSession"
          >
            <Trash2 :size="13" />
            {{ t('deleteSession', 'Delete session') }}
          </button>
        </div>
      </Transition>
    </div>

    <div class="flex items-center gap-1 rounded-md border border-accent/10">
      <CustomButton
        :title="settingsTitle"
        :aria-label="settingsTitle"
        :icon="Settings"
        text=""
        type="secondary"
        class="border-none p-1.5!"
        :icon-size="18"
        @click="$emit('settings')"
      />
    </div>
  </div>
</template>

<script lang="ts" setup>
import { ref, computed, onMounted, onBeforeUnmount } from 'vue';
import { useI18n } from 'vue-i18n';
import { Check, ChevronDown, MessageSquare, Plus, Settings, Trash2 } from 'lucide-vue-next';

import CustomButton from '@/components/CustomButton.vue';
import type { ChatSession } from '@/composables/useSessionManager';
import { getSessionMessageCount } from '@/composables/useSessionManager';

const { t } = useI18n();

const props = defineProps<{
  settingsTitle: string;
  loading: boolean;
  sessions: ChatSession[];
  currentSessionId: string | null;
}>();

const emit = defineEmits<{
  (e: 'new-chat'): void;
  (e: 'settings'): void;
  (e: 'switch-session', sessionId: string): void;
  (e: 'delete-session'): void;
}>();

const dropdownOpen = ref(false);
const dropdownRef = ref<HTMLElement>();

const currentSessionName = computed(() => {
  const session = props.sessions.find(s => s.id === props.currentSessionId);
  const name = session?.name ?? t('newChat');
  return name.length > 30 ? `${name.slice(0, 28)}…` : name;
});

function handleNewChat() {
  dropdownOpen.value = false;
  emit('new-chat');
}

function handleSwitchSession(sessionId: string) {
  dropdownOpen.value = false;
  emit('switch-session', sessionId);
}

function handleDeleteSession() {
  dropdownOpen.value = false;
  emit('delete-session');
}

function handleClickOutside(event: MouseEvent) {
  if (dropdownRef.value && !dropdownRef.value.contains(event.target as Node)) {
    dropdownOpen.value = false;
  }
}

onMounted(() => {
  document.addEventListener('mousedown', handleClickOutside);
});

onBeforeUnmount(() => {
  document.removeEventListener('mousedown', handleClickOutside);
});
</script>
