<template>
  <div
    :class="tight ? 'mb-0' : 'mb-3'"
    class="flex items-center gap-2 text-sm font-medium text-main"
  >
    <slot name="icon">
      <component :is="icon" v-if="icon" :size="iconSize" class="text-accent" />
    </slot>
    <span class="text-sm leading-[1.4] font-semibold text-secondary">{{
      title
    }}</span>
    <span v-if="required" class="ml-1 text-danger">*</span>
  </div>
  <div ref="dropdownRef" class="sort-dropdown relative">
    <button
      ref="triggerRef"
      class="flex h-7 w-full cursor-pointer items-center justify-between gap-1 rounded-md border border-border-secondary px-2 py-1.5 text-sm leading-[1.4] text-main transition-all duration-fast ease-apple hover:border-accent-hover focus:[.active]:border-accent-hover focus:[.active]:shadow-md"
      :class="{ active: dropDownOpen }"
      @click="toggleDropdown()"
    >
      <component
        :is="customFrontIcon || SortAscIcon"
        v-if="fronticon"
        :size="14"
      />
      <span class="text-center text-xs font-medium text-secondary">{{
        placeholder || modelValue
      }}</span>
      <ChevronDownIcon :size="14" />
    </button>
    <div
      v-show="dropDownOpen"
      ref="optionsRef"
      class="sort-options absolute z-10 mt-1 mb-1 max-h-50 min-w-37.5 overflow-hidden overflow-y-auto rounded-md border border-border-secondary bg-bg-tertiary shadow-lg"
    >
      <button
        v-for="key in keyList"
        :key="key"
        class="block min-h-[unset] w-full cursor-pointer border-none bg-bg-tertiary px-2 py-1 text-center text-sm leading-[1.4] text-main transition-all duration-fast ease-apple hover:bg-accent/50"
        @click="selectItem(key)"
      >
        <slot name="item" :item="key"> {{ key }} </slot>
      </button>
    </div>
  </div>
</template>

<script setup lang="ts">
import { onClickOutside } from "@vueuse/core";
import { ChevronDownIcon, SortAscIcon } from "lucide-vue-next";
import { nextTick, ref, type Component } from "vue";

const dropdownRef = ref(null);
const modelValue = defineModel<string>();
const triggerRef = ref<HTMLElement | null>(null);
const optionsRef = ref<HTMLElement | null>(null);

function selectItem(key: string) {
  modelValue.value = key;
  dropDownOpen.value = false;
}

const dropDownOpen = ref(false);

async function toggleDropdown() {
  dropDownOpen.value = !dropDownOpen.value;

  if (dropDownOpen.value) {
    await nextTick();
    updatePosition();
  }
}

function updatePosition() {
  const trigger = triggerRef.value;
  const dropdown = optionsRef.value;
  if (!trigger || !dropdown) return;

  const rect = trigger.getBoundingClientRect();
  const dropdownHeight = dropdown.offsetHeight;
  const viewportHeight = window.innerHeight;

  const spaceBelow = viewportHeight - rect.bottom;
  const canFitBelow = spaceBelow > dropdownHeight + 10;

  if (!canFitBelow && rect.top > dropdownHeight) {
    dropdown.style.top = "auto";
    dropdown.style.bottom = "100%";
  } else {
    dropdown.style.top = "100%";
    dropdown.style.bottom = "auto";
  }

  let dropdownWidth = Math.max(rect.width, 160);
  dropdown.style.left = "0px";
  dropdown.style.width = `${dropdownWidth}px`;
}

onClickOutside(dropdownRef, () => {
  if (dropDownOpen.value) {
    dropDownOpen.value = false;
  }
});

const {
  title,
  placeholder = "",
  fronticon = true,
  keyList,
  icon = null,
  tight = true,
  iconSize = 18,
  customFrontIcon = null,
  required = false,
} = defineProps<{
  title: string;
  icon?: Component | null;
  iconSize?: number;
  tight?: boolean;
  placeholder?: string;
  fronticon?: boolean;
  customFrontIcon?: Component | null;
  keyList: string[];
  required?: boolean;
}>();
</script>
