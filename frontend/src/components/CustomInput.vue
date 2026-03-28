<template>
  <div class="flex flex-col">
    <div class="mb-1 flex items-center gap-2">
      <label class="text-sm font-semibold text-secondary"
        >{{ title }}
        <span v-if="required" class="ml-1 text-danger">*</span>
      </label>
      <slot name="title-extra"></slot>
    </div>
    <div class="relative flex w-full gap-2">
      <input
        v-model="modelValue"
        :type="type"
        v-bind="$attrs"
        class="box-border w-full rounded-md border border-border bg-bg-tertiary p-2 pr-10 text-sm text-main transition-all duration-fast ease-apple focus:border-accent focus:outline-none focus:ring-2 focus:ring-accent/30 disabled:cursor-not-allowed disabled:opacity-50"
        :placeholder="placeholder"
      />
      <button
        v-if="isPassword"
        type="button"
        class="absolute top-1/2 right-3 flex -translate-y-1/2 items-center justify-center text-main"
        @click="type = type === 'password' ? 'text' : 'password'"
      >
        <EyeIcon v-if="type === 'password'" class="text-accent" :size="ICON_SIZE_MD" />
        <EyeClosedIcon v-else class="text-accent" :size="ICON_SIZE_MD" />
      </button>
      <slot name="input-extra"></slot>
    </div>
  </div>
</template>

<script setup lang="ts">
import { EyeClosedIcon, EyeIcon } from 'lucide-vue-next';
import { ref } from 'vue';
import { ICON_SIZE_MD } from '@/constants/limits';

const [modelValue, modifiers] = defineModel<string | number>({
  set(value) {
    let result = value;
    if (modifiers.trim && typeof result === 'string') {
      result = result.trim();
    }
    if (modifiers.number && typeof result === 'string') {
      const n = parseFloat(result);
      result = isNaN(n) ? result : n;
    }
    return result;
  },
});

const {
  isPassword = false,
  title,
  inputType = 'text',
  placeholder,
  required = false,
} = defineProps<{
  isPassword?: boolean;
  title: string;
  inputType?: string;
  placeholder: string;
  required?: boolean;
}>();

defineOptions({
  inheritAttrs: false,
});

const type = ref(isPassword ? 'password' : inputType);
</script>
