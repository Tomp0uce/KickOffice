import type { ToolDefinition } from '@/types';

import { logService } from '@/utils/logger';
import { executeOfficeAction } from './officeAction';
import { renderOfficeRichHtml, sanitizeHtml } from './markdown';
import { sandboxedEval } from './sandbox';
import { validateOfficeCode } from './officeCodeValidator';

import {
  generateVisualDiff,
  createOfficeTools,
  truncateString,
  buildExecuteWrapper,
  type OfficeToolTemplate,
  getErrorMessage,
  getDetailedOfficeError,
} from './common';
import { getLastRichContext, setLastRichContext } from './richContextStore';
import { reassembleWithFragments, extractTextFromHtml } from './richContentPreserver';
import { OUTLOOK_ACTION_TIMEOUT_MS } from '@/constants/limits';

export type OutlookToolName =
  | 'getEmailBody'
  | 'writeEmailBody'
  | 'getEmailSubject'
  | 'setEmailSubject'
  | 'getEmailRecipients'
  | 'addRecipient'
  | 'getEmailSender'
  | 'eval_outlookjs';

type RecipientField = 'to' | 'cc' | 'bcc';

function getMailbox(): any | null {
  return (window as any).Office?.context?.mailbox ?? null;
}

function getOfficeAsyncStatus(): any {
  return (window as any).Office?.AsyncResultStatus;
}

function getOfficeCoercionType(): any {
  return (window as any).Office?.CoercionType;
}

const runOutlook = <T>(action: () => Promise<T>): Promise<T> =>
  executeOfficeAction(action, 'outlook_action', OUTLOOK_ACTION_TIMEOUT_MS);

type OutlookToolTemplate = OfficeToolTemplate & {
  executeOutlook: (mailbox: any | null, args: Record<string, any>) => Promise<string>;
};

/**
 * QUAL-L2: Helper to bridge Outlook's callback-based Office.AsyncResult with async/await patterns.
 *
 * **Why this exists:**
 * The Outlook JavaScript API uses callback-based patterns (Office.AsyncResult) instead of Promises.
 * This helper wraps the AsyncResult callback pattern to work seamlessly with our async/await codebase.
 *
 * **Pattern:**
 * ```typescript
 * // Instead of this (callback style):
 * mailbox.item.subject.getAsync((result) => {
 *   if (result.status === Office.AsyncResultStatus.Succeeded) {
 *     return result.value
 *   } else {
 *     throw new Error(result.error.message)
 *   }
 * })
 *
 * // We wrap it in a Promise and use this helper:
 * return new Promise<string>((resolve) => {
 *   mailbox.item.subject.getAsync((result) => {
 *     resolve(resolveAsyncResult(result, (value) => value))
 *   })
 * })
 * ```
 *
 * @param result - The Office.AsyncResult object from the callback
 * @param onSuccess - Callback to transform the successful result.value into the desired output
 * @returns The transformed result on success
 * @throws Error if the AsyncResult status is not Succeeded
 */
function resolveAsyncResult(result: any, onSuccess: (value: any) => string): string {
  if (result.status === getOfficeAsyncStatus()?.Succeeded) {
    return onSuccess(result.value);
  }
  throw new Error(`Error: ${result.error?.message || 'unknown error'}`);
}

function normalizeRecipient(recipient: any): { displayName: string; emailAddress: string } {
  if (!recipient) {
    return { displayName: '', emailAddress: '' };
  }

  if (typeof recipient === 'string') {
    return { displayName: '', emailAddress: recipient.trim() };
  }

  return {
    displayName: recipient.displayName || recipient.name || '',
    emailAddress: recipient.emailAddress || recipient.address || '',
  };
}

function normalizeRecipientsInput(recipients: any): any[] {
  if (Array.isArray(recipients)) {
    return recipients.map(normalizeRecipient).filter(r => !!r.emailAddress);
  }

  if (typeof recipients === 'string') {
    return recipients
      .split(',')
      .map(email => normalizeRecipient(email))
      .filter(r => !!r.emailAddress);
  }

  if (recipients && typeof recipients === 'object') {
    const normalized = normalizeRecipient(recipients);
    return normalized.emailAddress ? [normalized] : [];
  }

  return [];
}

function getRecipientField(field: unknown): RecipientField {
  const value = String(field || 'to').toLowerCase();
  if (value === 'cc' || value === 'bcc') return value;
  return 'to';
}

const outlookToolDefinitions = createOfficeTools<
  OutlookToolName,
  OutlookToolTemplate,
  ToolDefinition
>(
  {
    getEmailBody: {
      name: 'getEmailBody',
      category: 'read',
      description:
        'Get the full body text of the current email. Works in both read and compose mode. Automatically captures images for preservation when modifying the email later.',
      inputSchema: {
        type: 'object',
        properties: {},
        required: [],
      },
      executeOutlook: async mailbox => {
        if (!mailbox?.item) return 'No email item available.';

        // First get HTML to capture images for preservation
        const htmlPromise = new Promise<string>((resolve, reject) => {
          mailbox.item.body.getAsync(getOfficeCoercionType().Html, (result: any) => {
            if (result.status === getOfficeAsyncStatus()?.Succeeded) {
              resolve(result.value || '');
            } else {
              reject(new Error('Failed to get HTML'));
            }
          });
        });

        try {
          const htmlContent = await htmlPromise;
          if (htmlContent) {
            const richContext = extractTextFromHtml(htmlContent);
            // Store rich context for later use by writeEmailBody
            if (richContext.hasRichContent) {
              setLastRichContext(richContext);
            }
            // Return clean text
            return richContext.cleanText || '';
          }
        } catch (err) {
          logService.warn('[getEmailBody] Failed to get HTML, falling back to text', err);
        }

        // Fallback to plain text
        return new Promise<string>(resolve => {
          mailbox.item.body.getAsync(getOfficeCoercionType().Text, (result: any) => {
            resolve(resolveAsyncResult(result, value => value || ''));
          });
        });
      },
    },

    writeEmailBody: {
      name: 'writeEmailBody',
      category: 'write',
      description:
        'The PREFERRED tool for modifying the email body. Supports Markdown (bold, italic, lists) and automatically preserves images from the original email. Can replace the whole body, prepend/append, or insert at the cursor. Only works in compose mode. CRITICAL: When replying/forwarding, ALWAYS use mode "Prepend" — it inserts the reply BEFORE the quoted history (standard email convention). NEVER use "Append" for replies (it puts the reply after the thread). NEVER use "Replace" on replies as it deletes the conversation thread. When improving/modifying existing email content, use mode "Replace" to update the email body while preserving images.',
      inputSchema: {
        type: 'object',
        properties: {
          content: {
            type: 'string',
            description: 'The content to write in Markdown format.',
          },
          mode: {
            type: 'string',
            description:
              'How to write the content: "Replace" (full overwrite), "Prepend" (add BEFORE thread history — USE THIS FOR REPLIES), "Append" (add to end), or "Insert" (add at cursor). Default: "Prepend".',
            enum: ['Replace', 'Prepend', 'Append', 'Insert'],
          },
          diffTracking: {
            type: 'boolean',
            description:
              'When mode is "Insert", shows a visual diff comparing content with current selection. Requires originalText. Default: false.',
          },
          originalText: {
            type: 'string',
            description: 'Required if diffTracking is true: the original text to compare against.',
          },
        },
        required: ['content'],
      },
      executeOutlook: async (mailbox, args: Record<string, any>) => {
        const { content, mode = 'Prepend', diffTracking = false, originalText = '' } = args;
        if (!mailbox?.item?.body || typeof mailbox.item.body.setAsync !== 'function') {
          return 'Cannot write email body: compose mode is not available.';
        }

        // Check if we have preserved rich content (images) to reassemble
        const richContext = getLastRichContext();
        let processedContent = content;

        // If the content contains {{PRESERVE_N}} placeholders and we have a rich context, reassemble
        if (richContext?.hasRichContent && content.includes('{{PRESERVE_')) {
          processedContent = reassembleWithFragments(content, richContext);
        }

        const html =
          diffTracking && mode === 'Insert' && originalText
            ? generateVisualDiff(originalText, processedContent)
            : renderOfficeRichHtml(processedContent);

        return new Promise<string>(resolve => {
          const body = mailbox.item.body;

          if (mode === 'Replace') {
            // Safety guard to prevent deleting thread history
            body.getAsync(getOfficeCoercionType().Html, {}, (getResult: any) => {
              if (getResult.status === getOfficeAsyncStatus()?.Succeeded) {
                const existing = getResult.value || '';
                // Common Outlook web/desktop thread markers
                if (
                  existing.includes('<div id="divRplyFwdMsg">') ||
                  existing.includes('<hr tabindex="-1"') ||
                  existing.includes('<hr ')
                ) {
                  logService.warn(
                    'Thread history detected. Overriding "Replace" to "Insert" to protect email history.',
                  );
                  body.setSelectedDataAsync(
                    html,
                    { coercionType: getOfficeCoercionType().Html },
                    (res: any) => {
                      resolve(
                        resolveAsyncResult(
                          res,
                          () => 'Replaced content at cursor (protected history).',
                        ),
                      );
                    },
                  );
                  return;
                }
              }
              // Safe to replace if no history markers found
              body.setAsync(html, { coercionType: getOfficeCoercionType().Html }, (res: any) => {
                resolve(resolveAsyncResult(res, () => 'Successfully replaced email body.'));
              });
            });
          } else if (mode === 'Prepend') {
            // Insert reply BEFORE the quoted thread history — standard email convention.
            // Detects Outlook thread separators and places new content before them.
            body.getAsync(getOfficeCoercionType().Html, {}, (getResult: any) => {
              if (getResult.status !== getOfficeAsyncStatus()?.Succeeded) {
                resolve('Error: Could not read body to prepend.');
                return;
              }
              const existing = getResult.value || '';
              // Common Outlook web/desktop thread history markers
              const separators = [
                '<div id="divRplyFwdMsg">',
                '<hr tabindex="-1"',
                '<hr ',
              ];
              let insertPos = -1;
              for (const sep of separators) {
                const pos = existing.indexOf(sep);
                if (pos !== -1 && (insertPos === -1 || pos < insertPos)) insertPos = pos;
              }
              let newBody: string;
              if (insertPos !== -1) {
                // Insert new reply before thread history
                newBody =
                  existing.substring(0, insertPos) +
                  sanitizeHtml(html) +
                  '<br/><br/>' +
                  existing.substring(insertPos);
              } else {
                // No history found — prepend to body
                newBody = sanitizeHtml(html) + (existing.trim() ? '<br/><br/>' + existing : '');
              }
              body.setAsync(
                newBody,
                { coercionType: getOfficeCoercionType().Html },
                (setResult: any) => {
                  resolve(
                    resolveAsyncResult(setResult, () =>
                      'Successfully prepended reply before email history.',
                    ),
                  );
                },
              );
            });
          } else if (mode === 'Append') {
            body.getAsync(getOfficeCoercionType().Html, {}, (getResult: any) => {
              if (getResult.status !== getOfficeAsyncStatus()?.Succeeded) {
                resolve('Error: Could not read body to append.');
                return;
              }
              const existing = getResult.value || '';
              const separator = existing.trim() ? '<br/><br/>' : '';
              const newBody = existing + separator + sanitizeHtml(html);
              body.setAsync(
                newBody,
                { coercionType: getOfficeCoercionType().Html },
                (setResult: any) => {
                  resolve(
                    resolveAsyncResult(setResult, () => 'Successfully appended to email body.'),
                  );
                },
              );
            });
          } else {
            // Insert at cursor
            body.setSelectedDataAsync(
              html,
              { coercionType: getOfficeCoercionType().Html },
              (res: any) => {
                resolve(
                  resolveAsyncResult(res, () =>
                    diffTracking ? 'Inserted visual diff.' : 'Successfully inserted at cursor.',
                  ),
                );
              },
            );
          }
        });
      },
    },

    getEmailSubject: {
      name: 'getEmailSubject',
      category: 'read',
      description: 'Get the current email subject in read or compose mode.',
      inputSchema: {
        type: 'object',
        properties: {},
        required: [],
      },
      executeOutlook: async mailbox => {
        if (!mailbox?.item) return 'No email item available.';

        if (mailbox.item.subject && typeof mailbox.item.subject.getAsync === 'function') {
          return new Promise<string>(resolve => {
            mailbox.item.subject.getAsync((result: any) => {
              resolve(resolveAsyncResult(result, value => value || ''));
            });
          });
        }

        return mailbox.item.subject || '';
      },
    },

    setEmailSubject: {
      name: 'setEmailSubject',
      category: 'write',
      description: 'Set the email subject in compose mode.',
      inputSchema: {
        type: 'object',
        properties: {
          subject: {
            type: 'string',
            description: 'The new email subject line',
          },
        },
        required: ['subject'],
      },
      executeOutlook: async (mailbox, args: Record<string, any>) => {
        const { subject } = args as Record<string, any>;
        if (!mailbox?.item?.subject || typeof mailbox.item.subject.setAsync !== 'function') {
          return 'Cannot set email subject: compose mode is not available.';
        }

        return new Promise<string>(resolve => {
          mailbox.item.subject.setAsync(subject, (result: any) => {
            resolve(resolveAsyncResult(result, () => 'Successfully updated email subject.'));
          });
        });
      },
    },

    getEmailRecipients: {
      name: 'getEmailRecipients',
      category: 'read',
      description: 'Get the current To, Cc, and Bcc recipients of the email.',
      inputSchema: {
        type: 'object',
        properties: {},
        required: [],
      },
      executeOutlook: async mailbox => {
        if (!mailbox?.item) return 'No email item available.';

        const item = mailbox.item;

        const getField = (field: RecipientField) => {
          const fieldObject = item[field];
          if (fieldObject && typeof fieldObject.getAsync === 'function') {
            return new Promise<any[]>(resolve => {
              fieldObject.getAsync((result: any) => {
                if (
                  result.status === getOfficeAsyncStatus()?.Succeeded &&
                  Array.isArray(result.value)
                ) {
                  resolve(result.value.map(normalizeRecipient));
                } else {
                  resolve([]);
                }
              });
            });
          }

          if (Array.isArray(fieldObject)) {
            return Promise.resolve(fieldObject.map(normalizeRecipient));
          }

          return Promise.resolve([]);
        };

        const [to, cc, bcc] = await Promise.all([getField('to'), getField('cc'), getField('bcc')]);

        return JSON.stringify({ to, cc, bcc });
      },
    },

    addRecipient: {
      name: 'addRecipient',
      category: 'write',
      description: 'Add recipient(s) to To, Cc, or Bcc in compose mode.',
      inputSchema: {
        type: 'object',
        properties: {
          field: {
            type: 'string',
            enum: ['to', 'cc', 'bcc'],
            description: 'Recipient field to update (to, cc, bcc). Defaults to to.',
          },
          recipients: {
            type: 'string',
            description:
              'Recipient(s) to add. Provide a single email address or a comma-separated list of emails.',
          },
        },
        required: ['recipients'],
      },
      executeOutlook: async (mailbox, args: Record<string, any>) => {
        if (!mailbox?.item) return 'No email item available.';

        const field = getRecipientField(args.field);
        const recipients = normalizeRecipientsInput(args.recipients);
        if (recipients.length === 0) {
          return 'No valid recipients provided.';
        }

        const fieldObject = mailbox.item[field];
        if (!fieldObject || typeof fieldObject.addAsync !== 'function') {
          return `Cannot add recipients to ${field}: compose mode is not available.`;
        }

        return new Promise<string>(resolve => {
          fieldObject.addAsync(recipients, (result: any) => {
            resolve(
              resolveAsyncResult(
                result,
                () => `Successfully added ${recipients.length} recipient(s) to ${field}.`,
              ),
            );
          });
        });
      },
    },

    getEmailSender: {
      name: 'getEmailSender',
      category: 'read',
      description: 'Get sender information for the current email.',
      inputSchema: {
        type: 'object',
        properties: {},
        required: [],
      },
      executeOutlook: async mailbox => {
        if (!mailbox?.item) return 'No email item available.';

        const sender = mailbox.item.from || mailbox.item.sender;
        if (!sender) return '';

        return JSON.stringify(normalizeRecipient(sender));
      },
    },

    eval_outlookjs: {
      name: 'eval_outlookjs',
      category: 'write',
      description: `Execute custom Office.js code within the Outlook mailbox context.

**USE THIS TOOL ONLY WHEN:**
- No dedicated tool exists for your operation
- Operations like: attachments, HTML manipulation, advanced metadata

**REQUIRED CODE STRUCTURE:**
\`\`\`javascript
try {
  return new Promise((resolve, reject) => {
    mailbox.item.subject.getAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve({ success: true, result: result.value });
      } else {
        reject({ success: false, error: result.error.message });
      }
    });
  });
} catch (error) {
  return { success: false, error: error.message };
}
\`\`\`

**CRITICAL RULES:**
1. Outlook uses CALLBACKS not async/await (wrap in Promise)
2. ALWAYS wrap in try/catch
3. ONLY use Office.context.mailbox APIs
4. Check AsyncResultStatus before reading result.value`,
      inputSchema: {
        type: 'object',
        properties: {
          code: {
            type: 'string',
            description:
              'JavaScript code following the template. Must use callbacks and Promise wrapper.',
          },
          explanation: {
            type: 'string',
            description: 'Brief explanation of what this code does (required for audit trail).',
          },
        },
        required: ['code', 'explanation'],
      },
      executeOutlook: async (mailbox, args: Record<string, any>) => {
        const { code, explanation } = args;

        // Validate code BEFORE execution
        // Note: Outlook doesn't use context.sync(), so validation is less strict
        const validation = validateOfficeCode(code, 'Outlook');

        if (!validation.valid) {
          return JSON.stringify(
            {
              success: false,
              error: 'Code validation failed. Fix the errors below and try again.',
              validationErrors: validation.errors,
              validationWarnings: validation.warnings,
              suggestion:
                'Refer to the Office.js skill document for correct patterns. Remember: Outlook uses callbacks, not async/await.',
              codeReceived: truncateString(code, 300),
            },
            null,
            2,
          );
        }

        // Log warnings but proceed
        if (validation.warnings.length > 0) {
          logService.warn('[eval_outlookjs] Validation warnings:', validation.warnings);
        }

        try {
          // Execute in sandbox with host restriction
          const result = await sandboxedEval(
            code,
            {
              mailbox,
              Office:
                typeof (window as any).Office !== 'undefined' ? (window as any).Office : undefined,
            },
            'Outlook', // Restrict to Outlook namespace only
          );

          return JSON.stringify(
            {
              success: true,
              result: result ?? null,
              explanation,
              warnings: validation.warnings.length > 0 ? validation.warnings : undefined,
            },
            null,
            2,
          );
        } catch (err: unknown) {
          return JSON.stringify(
            {
              success: false,
              error: getDetailedOfficeError(err),
              explanation,
              codeExecuted: truncateString(code, 200),
              hint: 'Check callback patterns and Promise wrapping. Outlook uses callbacks, not async/await.',
            },
            null,
            2,
          );
        }
      },
    },
  },
  def =>
    async (args = {}) => {
      try {
        return await runOutlook(async () =>
          Promise.race([
            def.executeOutlook(getMailbox(), args),
            new Promise<string>(resolve =>
              setTimeout(
                () =>
                  resolve(
                    `Error: Outlook API request timed out after ${OUTLOOK_ACTION_TIMEOUT_MS / 1000} seconds.`,
                  ),
                OUTLOOK_ACTION_TIMEOUT_MS,
              ),
            ),
          ]),
        );
      } catch (error: unknown) {
        return JSON.stringify(
          {
            success: false,
            error: getErrorMessage(error),
            tool: def.name,
            suggestion: 'Fix the error parameters or context and try again.',
          },
          null,
          2,
        );
      }
    },
);

export function getOutlookToolDefinitions(): ToolDefinition[] {
  return Object.values(outlookToolDefinitions);
}

export { outlookToolDefinitions };
