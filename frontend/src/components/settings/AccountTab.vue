<template>
  <div role="tabpanel" class="flex h-full w-full flex-col items-center gap-2 bg-bg-secondary p-1">
    <SettingCard>
      <CustomInput
        v-model="litellmUserKey"
        :title="t('litellmUserKeyLabel')"
        :placeholder="t('litellmUserKeyPlaceholder')"
        :is-password="true"
      />
    </SettingCard>

    <SettingCard>
      <CustomInput
        v-model="litellmUserEmail"
        :title="t('litellmUserEmailLabel')"
        :placeholder="t('litellmUserEmailPlaceholder')"
      />
    </SettingCard>

    <SettingCard>
      <label class="flex cursor-pointer items-center justify-between gap-2">
        <div class="flex flex-col">
          <span class="text-sm font-semibold text-main">{{ t('rememberCredentialsLabel') }}</span>
          <span class="text-xs text-secondary">{{ t('rememberCredentialsDescription') }}</span>
        </div>
        <input
          v-model="rememberCredentials"
          type="checkbox"
          class="h-4 w-4 cursor-pointer accent-accent"
          :aria-label="t('rememberCredentialsLabel')"
        />
      </label>
    </SettingCard>

    <SettingCard>
      <div class="flex items-center justify-between">
        <span class="text-sm font-semibold text-secondary">{{
          t('litellmCredentialsMissing')
        }}</span>
        <div
          class="flex items-center gap-1 rounded-md px-2 py-1 text-xs"
          :class="
            litellmConfigured ? 'bg-green-100 text-green-700' : 'bg-yellow-100 text-yellow-700'
          "
        >
          <div
            class="h-2 w-2 rounded-full"
            :class="litellmConfigured ? 'bg-green-500' : 'bg-yellow-500'"
          />
          {{
            litellmConfigured ? t('litellmCredentialsConfigured') : t('litellmCredentialsMissing')
          }}
        </div>
      </div>
    </SettingCard>

    <SettingCard>
      <div class="flex flex-col gap-1">
        <span class="text-xs text-secondary">{{ t('litellmCredentialsInfo') }}</span>
        <a
          href="https://getkey.ai.kickmaker.net/"
          target="_blank"
          rel="noopener noreferrer"
          class="text-xs text-accent underline"
          >{{ t('getApiKeyLink') }}</a
        >
      </div>
    </SettingCard>

    <!-- Crypto Warning -->
    <SettingCard v-if="!cryptoAvailable">
      <div class="flex items-start gap-2 rounded-md bg-yellow-50 p-2">
        <span class="text-2xl">⚠️</span>
        <div class="flex flex-col gap-1">
          <span class="text-sm font-semibold text-yellow-800">
            {{ t('cryptoNotAvailableTitle') }}
          </span>
          <span class="text-xs text-yellow-700">
            {{ t('cryptoNotAvailableMessage') }}
          </span>
        </div>
      </div>
    </SettingCard>
  </div>
</template>

<script setup lang="ts">
import { computed, onMounted, watch, ref } from 'vue';
import { useI18n } from 'vue-i18n';

import CustomInput from '@/components/CustomInput.vue';
import SettingCard from '@/components/SettingCard.vue';
import {
  getUserKey,
  setUserKey,
  getUserEmail,
  setUserEmail,
  getRememberCredentials,
  setRememberCredentials as setRememberCredentialsPersist,
} from '@/utils/credentialStorage';
import { isCryptoAvailable } from '@/utils/cryptoPolyfill';
import { invalidateHeaderCache } from '@/api/backend';

const { t } = useI18n();

const litellmUserKey = ref('');
const litellmUserEmail = ref('');
const rememberCredentials = ref(getRememberCredentials());
const cryptoAvailable = ref(isCryptoAvailable());

const litellmConfigured = computed(() => {
  return litellmUserKey.value.length > 0 && litellmUserEmail.value.length > 0;
});

watch(litellmUserKey, async value => {
  await setUserKey(value);
  invalidateHeaderCache();
});

watch(litellmUserEmail, async value => {
  await setUserEmail(value);
  invalidateHeaderCache();
});

watch(rememberCredentials, async value => {
  await setRememberCredentialsPersist(value);
  litellmUserKey.value = await getUserKey();
  litellmUserEmail.value = await getUserEmail();
});

onMounted(async () => {
  litellmUserKey.value = await getUserKey();
  litellmUserEmail.value = await getUserEmail();
});

// Expose so parent can override values after credential migration
defineExpose({
  setCredentials(key: string, email: string) {
    litellmUserKey.value = key;
    litellmUserEmail.value = email;
  },
});
</script>
