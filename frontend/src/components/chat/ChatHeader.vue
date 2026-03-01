<template>
  <div
    class="flex justify-between rounded-sm border border-[#33ABC6]/20 bg-surface/90 p-1"
  >
    <div class="flex flex-1 items-center gap-2 text-accent">
      <img
        src="/Logo.png"
        alt="KickAI logo"
        class="h-8 w-8 rounded-sm border border-black/10 bg-white object-contain p-0.5"
      />
      <div class="flex flex-col leading-none">
        <span class="text-sm font-semibold text-main">{{ t("appTitle") }}</span>
        <span class="text-[10px] text-[#33ABC6]">{{
          t("aiOfficeAssistant", "AI Office Assistant")
        }}</span>
      </div>
    </div>
    <div class="mr-1 flex items-center gap-1">
      <span class="h-2 w-2 rounded-full bg-[#33ABC6]" />
      <span class="h-2 w-2 rounded-full bg-black" />
      <span class="h-2 w-2 rounded-full border border-black/30 bg-white" />
    </div>
    <div class="flex items-center gap-1 rounded-md border border-accent/10">
      <CustomButton
        :title="clearChatTitle"
        :icon="Eraser"
        text=""
        type="secondary"
        class="border-none p-1!"
        :icon-size="18"
        @click="$emit('new-chat')"
      />
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
import { useI18n } from "vue-i18n";
import { Eraser, Settings } from "lucide-vue-next";

import CustomButton from "@/components/CustomButton.vue";

const { t } = useI18n();

defineProps<{ clearChatTitle: string; settingsTitle: string }>();
defineEmits<{ (e: "new-chat"): void; (e: "settings"): void }>();
</script>
