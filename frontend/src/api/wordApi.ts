import { Ref } from 'vue'

import { WordFormatter } from '@/utils/wordFormatter'

export async function insertResult(result: string, insertType: Ref<string>): Promise<void> {
  // GN-M2: Delegating the plain insertion logic to WordFormatter to avoid code duplication
  await WordFormatter.insertPlainResult(result, insertType)
}

export async function insertFormattedResult(result: string, insertType: Ref<string>): Promise<void> {
  try {
    await WordFormatter.insertFormattedResult(result, insertType)
  } catch (error) {
    console.warn('Formatted insertion failed, falling back to plain text:', error)
    await WordFormatter.insertPlainResult(result, insertType)
  }
}
