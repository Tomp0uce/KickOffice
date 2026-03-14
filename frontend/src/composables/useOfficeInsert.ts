import type { InsertType } from '@/types';
import { type Ref, ref } from 'vue';

import { logService } from '@/utils/logger';

import type { DisplayMessage } from '@/types/chat';
import { insertFormattedResult, insertResult } from '@/api/wordApi';
import { message as messageUtil } from '@/utils/message';
import {
  getOfficeHtmlCoercionType,
  getOutlookMailbox,
  isOfficeAsyncSucceeded,
  type OfficeAsyncResult,
} from '@/utils/officeOutlook';
import { insertIntoPowerPoint } from '@/utils/powerpointTools';
import { renderOfficeCommonApiHtml } from '@/utils/markdown';
import DOMPurify from 'dompurify';

const VERBOSE_LOGGING_ENABLED = import.meta.env.VITE_VERBOSE_LOGGING === 'true';
const verboseLog = VERBOSE_LOGGING_ENABLED ? console.info.bind(console, '[KO-INSERT]') : () => {};

async function doOutlookInsert(
  content: string,
  richHtml: string | undefined,
  copyToClipboard: Function,
  t: (key: string) => string,
) {
  try {
    const mailbox = getOutlookMailbox();
    const item = mailbox?.item;
    if (item?.body?.setSelectedDataAsync) {
      const rawHtmlBody = richHtml || renderOfficeCommonApiHtml(content);
      const htmlBody = DOMPurify.sanitize(rawHtmlBody, { USE_PROFILES: { html: true } });
      await new Promise<void>((resolve, reject) => {
        item.body.setSelectedDataAsync!(
          htmlBody,
          { coercionType: getOfficeHtmlCoercionType() },
          (result: OfficeAsyncResult) => {
            if (isOfficeAsyncSucceeded(result.status)) resolve();
            else reject(new Error(result.error?.message || 'setSelectedDataAsync failed'));
          },
        );
      });
      messageUtil.success(t('insertedToEmail'));
    } else {
      await copyToClipboard(content, true);
    }
  } catch (err) {
    logService.warn('[useOfficeInsert] Outlook error/fallback to clipboard', err);
    await copyToClipboard(content, true);
  }
}

async function doPowerPointInsert(
  content: string,
  copyToClipboard: Function,
  t: (key: string) => string,
) {
  try {
    await insertIntoPowerPoint(content);
    messageUtil.success(t('insertedToSlide'));
  } catch (err) {
    logService.warn('[useOfficeInsert] PowerPoint error/fallback to clipboard', err);
    await copyToClipboard(content, true);
  }
}

async function doExcelInsert(
  content: string,
  copyToClipboard: Function,
  t: (key: string) => string,
) {
  try {
    await Excel.run(async ctx => {
      const range = ctx.workbook.getSelectedRange();
      range.values = [[content]];
      range.format.wrapText = true;
      await ctx.sync();
    });
    messageUtil.success(t('insertedToCell'));
  } catch (err) {
    logService.warn('[useOfficeInsert] Excel error/fallback to clipboard', err);
    await copyToClipboard(content, true);
  }
}

async function doWordInsert(
  content: string,
  type: InsertType,
  richHtml: string | undefined,
  useWordFormatting: boolean,
  copyToClipboard: Function,
  t: (key: string) => string,
) {
  try {
    if (richHtml) {
      const sanitizedHtml = DOMPurify.sanitize(richHtml, { USE_PROFILES: { html: true } });
      await Word.run(async context => {
        const range = context.document.getSelection();
        range.insertHtml(sanitizedHtml, type === 'newLine' ? 'After' : 'Replace');
        await context.sync();
      });
      messageUtil.success(t('inserted'));
    } else if (useWordFormatting) {
      await insertFormattedResult(content, ref(type));
      messageUtil.success(t('inserted'));
    } else {
      await insertResult(content, ref(type));
      messageUtil.success(t('inserted'));
    }
  } catch (err) {
    logService.warn('[useOfficeInsert] Word error/fallback to clipboard', err);
    await copyToClipboard(content, true);
  }
}

interface UseOfficeInsertOptions {
  hostIsOutlook: boolean;
  hostIsPowerPoint: boolean;
  hostIsExcel: boolean;
  hostIsWord: boolean;
  useWordFormatting: Ref<boolean>;

  t: (key: string) => string;
  shouldTreatMessageAsImage: (message: DisplayMessage) => boolean;
  getMessageActionPayload: (message: DisplayMessage) => string;
  copyImageToClipboard: (imageSrc: string, fallback?: boolean) => Promise<void>;
  insertImageToWord: (imageSrc: string, type: InsertType) => Promise<void>;
  insertImageToPowerPoint: (imageSrc: string, type: InsertType) => Promise<void>;
}

export function useOfficeInsert(options: UseOfficeInsertOptions) {
  const {
    hostIsOutlook,
    hostIsPowerPoint,
    hostIsExcel,
    hostIsWord,
    useWordFormatting,

    t,
    shouldTreatMessageAsImage,
    getMessageActionPayload,
    copyImageToClipboard,
    insertImageToWord,
    insertImageToPowerPoint,
  } = options;

  function normalizeInsertionContent(rawContent: string): string {
    return rawContent.replace(/\r\n/g, '\n').replace(/\r/g, '\n').trim();
  }

  async function copyToClipboard(text: string, fallback = false) {
    if (!text.trim()) return;
    const notifySuccess = () => messageUtil.success(t(fallback ? 'copiedFallback' : 'copied'));
    try {
      await navigator.clipboard.writeText(text);
      notifySuccess();
      return;
    } catch (err) {
      logService.warn('Clipboard API writeText failed, trying fallback:', err);
    }
    try {
      const textarea = document.createElement('textarea');
      textarea.value = text;
      textarea.setAttribute('readonly', '');
      textarea.style.position = 'fixed';
      textarea.style.opacity = '0';
      document.body.appendChild(textarea);
      textarea.select();
      const copied = document.execCommand('copy');
      document.body.removeChild(textarea);
      if (copied) notifySuccess();
      else messageUtil.error(t('failedToInsert'));
    } catch {
      messageUtil.error(t('failedToInsert'));
    }
  }

  async function insertToDocument(content: string, type: InsertType, richHtml?: string) {
    const normalizedContent = normalizeInsertionContent(content);
    if (!normalizedContent) return;

    verboseLog('insertToDocument', {
      host: hostIsOutlook
        ? 'outlook'
        : hostIsPowerPoint
          ? 'powerpoint'
          : hostIsExcel
            ? 'excel'
            : 'word',
      type,
      contentLength: normalizedContent.length,
      lineCount: normalizedContent.split('\n').length,
      hasRichHtml: !!richHtml,
    });

    if (hostIsOutlook) {
      await doOutlookInsert(normalizedContent, richHtml, copyToClipboard, t);
      return;
    }

    if (hostIsPowerPoint) {
      await doPowerPointInsert(normalizedContent, copyToClipboard, t);
      return;
    }

    if (hostIsExcel) {
      await doExcelInsert(normalizedContent, copyToClipboard, t);
      return;
    }

    await doWordInsert(
      normalizedContent,
      type,
      richHtml,
      useWordFormatting.value,
      copyToClipboard,
      t,
    );
  }

  async function copyMessageToClipboard(message: DisplayMessage, fallback = false) {
    if (shouldTreatMessageAsImage(message) && message.imageSrc) {
      await copyImageToClipboard(message.imageSrc, fallback);
      return;
    }
    await copyToClipboard(getMessageActionPayload(message), fallback);
  }

  async function insertMessageToDocument(message: DisplayMessage, type: InsertType) {
    if (shouldTreatMessageAsImage(message) && message.imageSrc) {
      if (hostIsWord) {
        try {
          await insertImageToWord(message.imageSrc, type);
          messageUtil.success(t('inserted'));
        } catch {
          await copyImageToClipboard(message.imageSrc, true);
        }
        return;
      }
      if (hostIsPowerPoint) {
        try {
          await insertImageToPowerPoint(message.imageSrc, type);
          messageUtil.success(t('insertedToSlide'));
        } catch {
          await copyImageToClipboard(message.imageSrc, true);
        }
        return;
      }
      if (hostIsExcel) {
        messageUtil.info(t('imageInsertExcelNotSupported'));
        return;
      }
      // Outlook and other hosts: copy to clipboard with a host-appropriate message
      await copyImageToClipboard(message.imageSrc, true);
      messageUtil.info(t('imageInsertOutlookFallback'));
      return;
    }
    await insertToDocument(getMessageActionPayload(message), type, message.richHtml);
  }

  return { copyToClipboard, copyMessageToClipboard, insertToDocument, insertMessageToDocument };
}
