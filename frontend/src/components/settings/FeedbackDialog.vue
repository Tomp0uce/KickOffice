<template>
  <div class="fixed inset-0 z-50 flex items-center justify-center bg-black/50 p-4">
    <div class="card-base w-full max-w-md flex flex-col gap-4 bg-surface p-4">
      <h3 class="text-lg font-semibold text-main">
        {{ t('feedbackTitle') || 'Report Bug / Feedback' }}
      </h3>

      <div v-if="success" class="rounded-md bg-green-100 p-3 text-green-800 text-sm">
        {{ t('feedbackSuccess') || 'Thank you for your feedback!' }}
      </div>
      <div v-else class="flex flex-col gap-3">
        <label class="text-sm font-semibold text-main">{{
          t('feedbackCategory') || 'Category'
        }}</label>
        <select
          v-model="category"
          class="rounded-md border border-border p-2 text-sm bg-bg-secondary text-main"
        >
          <option value="bug">{{ t('feedbackBug') || 'Bug Report' }}</option>
          <option value="feature">{{ t('feedbackFeature') || 'Feature Request' }}</option>
          <option value="other">{{ t('feedbackOther') || 'Other' }}</option>
        </select>

        <label class="text-sm font-semibold text-main">{{
          t('feedbackComment') || 'Comment'
        }}</label>
        <textarea
          v-model="comment"
          rows="4"
          class="rounded-md border border-border p-2 text-sm bg-bg-secondary text-main focus:border-accent focus:outline-none"
          :placeholder="
            t('feedbackPlaceholder') || 'Please describe the issue or your suggestion here...'
          "
        ></textarea>

        <div class="flex items-center gap-2 mt-2">
          <input
            type="checkbox"
            id="includeLogs"
            v-model="includeLogs"
            class="h-4 w-4 cursor-pointer"
          />
          <label for="includeLogs" class="text-xs text-secondary cursor-pointer">
            {{ t('feedbackIncludeLogs') || 'Include recent session logs to help us debug' }}
          </label>
        </div>

        <div v-if="error" class="text-xs text-red-500 mt-1">{{ error }}</div>

        <div class="flex justify-end gap-2 mt-4">
          <CustomButton type="secondary" :text="t('cancel') || 'Cancel'" @click="$emit('close')" />
          <CustomButton
            type="primary"
            :text="submitting ? t('submitting') || 'Submitting...' : t('submit') || 'Submit'"
            :disabled="submitting || !comment.trim()"
            @click="handleSubmit"
          />
        </div>
      </div>
    </div>
  </div>
</template>

<script lang="ts" setup>
import { ref } from 'vue'
import { useI18n } from 'vue-i18n'
import CustomButton from '@/components/CustomButton.vue'
import { logService } from '@/utils/logger'
import { submitFeedback } from '@/api/backend'

const { t } = useI18n()
const emit = defineEmits(['close'])

const category = ref('bug')
const comment = ref('')
const includeLogs = ref(true)
const submitting = ref(false)
const success = ref(false)
const error = ref('')

async function handleSubmit() {
  if (!comment.value.trim()) return

  submitting.value = true
  error.value = ''

  try {
    const ctx = await logService.getContext()
    const sessionId = ctx.sessionId || 'unknown'
    const logs = includeLogs.value ? logService.getSessionLogs(sessionId) : []

    await submitFeedback(sessionId, {
      category: category.value,
      comment: comment.value,
      logs: logs,
    })

    success.value = true
    setTimeout(() => {
      emit('close')
    }, 2000)
  } catch (err: any) {
    error.value = err.message || 'Failed to submit feedback'
    logService.error('Feedback submission error', err)
  } finally {
    submitting.value = false
  }
}
</script>
