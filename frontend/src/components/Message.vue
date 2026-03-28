<template>
  <Teleport to="body">
    <Transition name="toast-slide">
      <div
        v-if="visible"
        class="toast-container"
        :class="`toast-${type}`"
        role="alert"
        aria-live="polite"
      >
        <div class="toast-content">
          <div class="toast-icon">
            <AlertCircle v-if="type === 'error'" />
            <CheckCircle v-if="type === 'success'" />
            <Info v-if="type === 'info'" />
            <AlertTriangle v-if="type === 'warning'" />
          </div>
          <span class="toast-text">{{ message }}</span>
        </div>
        <div class="toast-progress" :style="{ animationDuration: `${duration}ms` }"></div>
      </div>
    </Transition>
  </Teleport>
</template>

<script lang="ts" setup>
import { AlertCircle, AlertTriangle, CheckCircle, Info } from 'lucide-vue-next';
import { onMounted, onUnmounted, ref } from 'vue';

interface Props {
  message: string;
  type?: 'error' | 'success' | 'info' | 'warning';
  duration?: number;
}

const props = withDefaults(defineProps<Props>(), {
  type: 'info',
  duration: 3000,
});

const visible = ref(false);
const emit = defineEmits(['close']);

let hideTimeout: number;
let closeTimeout: number;

onMounted(() => {
  visible.value = true;
  if (props.duration > 0) {
    hideTimeout = window.setTimeout(() => {
      visible.value = false;
      closeTimeout = window.setTimeout(() => emit('close'), 300);
    }, props.duration);
  }
});

onUnmounted(() => {
  clearTimeout(hideTimeout);
  clearTimeout(closeTimeout);
});
</script>

<style scoped>
.toast-container {
  position: fixed;
  top: 16px;
  right: 16px;
  z-index: 9999;
  display: flex;
  align-items: stretch;
  overflow: hidden;
  border-radius: 8px;
  padding: 8px 10px;
  min-width: 20px;
  max-width: 360px;
  box-shadow:
    0 8px 24px rgb(0 0 0 / 12%),
    0 2px 8px rgb(0 0 0 / 8%);
  flex-direction: column;
  backdrop-filter: blur(12px);
}

.toast-content {
  display: flex;
  align-items: center;
  gap: 10px;
  width: 100%;
}

.toast-icon {
  display: flex;
  justify-content: center;
  align-items: center;
  width: 18px;
  height: 18px;
  flex-shrink: 0;
}

.toast-icon svg {
  width: 18px;
  height: 18px;
}

.toast-text {
  font-size: 13px;
  line-height: 1.4;
  flex: 1;
  font-weight: 500;
  word-break: break-word;
}

.toast-progress {
  position: absolute;
  bottom: 0;
  left: 0;
  width: 100%;
  height: 3px;
  transform-origin: left;
  animation: toast-progress linear forwards;
}

@keyframes toast-progress {
  from {
    transform: scaleX(1);
  }

  to {
    transform: scaleX(0);
  }
}

.toast-error {
  border: 1px solid color-mix(in srgb, var(--color-danger) 30%, transparent);
  color: var(--color-text-primary);
  background: color-mix(in srgb, var(--color-danger) 8%, var(--color-background-primary));
}

.toast-error .toast-icon svg {
  color: var(--color-danger);
}

.toast-error .toast-progress {
  background: var(--color-danger);
}

.toast-success {
  border: 1px solid color-mix(in srgb, var(--color-success) 30%, transparent);
  color: var(--color-text-primary);
  background: color-mix(in srgb, var(--color-success) 8%, var(--color-background-primary));
}

.toast-success .toast-icon svg {
  color: var(--color-success);
}

.toast-success .toast-progress {
  background: var(--color-success);
}

.toast-info {
  border: 1px solid color-mix(in srgb, var(--color-info) 30%, transparent);
  color: var(--color-text-primary);
  background: color-mix(in srgb, var(--color-info) 8%, var(--color-background-primary));
}

.toast-info .toast-icon svg {
  color: var(--color-info);
}

.toast-info .toast-progress {
  background: var(--color-info);
}

.toast-warning {
  border: 1px solid color-mix(in srgb, var(--color-warning) 30%, transparent);
  color: var(--color-text-primary);
  background: color-mix(in srgb, var(--color-warning) 8%, var(--color-background-primary));
}

.toast-warning .toast-icon svg {
  color: var(--color-warning);
}

.toast-warning .toast-progress {
  background: var(--color-warning);
}

.toast-slide-enter-active,
.toast-slide-leave-active {
  transition: all 0.3s cubic-bezier(0.3, 1, 0.3, 1);
}

.toast-slide-enter-from {
  opacity: 0;
  transform: translateX(100%) scale(0.8);
}

.toast-slide-leave-to {
  opacity: 0;
  transform: translateX(100%) scale(0.9);
}
</style>
