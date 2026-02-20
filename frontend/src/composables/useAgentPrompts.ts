import type { Ref } from 'vue'

export interface UseAgentPromptsOptions {
  t: (key: string) => string
  userGender: Ref<string>
  userFirstName: Ref<string>
  userLastName: Ref<string>
  excelFormulaLanguage: Ref<'en' | 'fr'>
  hostIsOutlook: boolean
  hostIsPowerPoint: boolean
  hostIsExcel: boolean
  hostIsWord: boolean
}

export function useAgentPrompts(options: UseAgentPromptsOptions) {
  const {
    t,
    userGender,
    userFirstName,
    userLastName,
    excelFormulaLanguage,
    hostIsOutlook,
    hostIsPowerPoint,
    hostIsExcel,
    hostIsWord,
  } = options

  const excelFormulaLanguageInstruction = () => excelFormulaLanguage.value === 'fr'
    ? 'Excel interface locale: French. Use localized French function names and separators when providing formulas, and prefer localized formula tool behavior.'
    : 'Excel interface locale: English. Use English function names and standard English formula syntax.'

  const userProfilePromptBlock = () => {
    const firstName = userFirstName.value.trim()
    const lastName = userLastName.value.trim()
    const fullName = `${firstName} ${lastName}`.trim() || t('userProfileUnknownName')
    const genderMap: Record<string, string> = {
      female: t('userGenderFemale'), male: t('userGenderMale'), nonbinary: t('userGenderNonBinary'), unspecified: t('userGenderUnspecified'),
    }
    const genderLabel = genderMap[userGender.value] || t('userGenderUnspecified')
    return `\n\nUser profile context for communications (especially emails):\n- First name: ${firstName || t('userProfileUnknownFirstName')}\n- Last name: ${lastName || t('userProfileUnknownLastName')}\n- Full name: ${fullName}\n- Gender: ${genderLabel}\nUse this profile when drafting salutations, signatures, and tone, unless the user asks otherwise.`
  }

  const wordAgentPrompt = (lang: string) => `# Role\nYou are a highly skilled Microsoft Word Expert Agent. Your goal is to assist users in creating, editing, and formatting documents with professional precision.\n\n# Capabilities\n- You can interact with the document directly using provided tools (reading text, applying styles, inserting content, etc.).\n- You understand document structure, typography, and professional writing standards.\n\n# Guidelines\n1. **Tool First**: If a request requires document modification or inspection, prioritize using the available tools.\n2. **Direct Actions**: For Word formatting requests (bold, underline, highlight, size, color, superscript, uppercase, tags like <format>...</format>, etc.), execute the change directly with tools instead of giving manual steps.\n3. **Formatting**: When generating or modifying document content, ALWAYS format your response using semantic HTML tags (e.g., <b>, <i>, <u>, <h1> to <h6>, <p>, <ul>, <li>, <br>) instead of plain text or markdown, as your output will be directly inserted into Word via HTML. Avoid markdown asterisks or underscores.\n4. **Accuracy**: Ensure formatting and content changes are precise and follow the user's intent.\n5. **Conciseness**: Provide brief, helpful explanations of your actions.\n6. **Language**: You must communicate entirely in ${lang}.\n\n# Safety\nDo not perform destructive actions (like clearing the whole document) unless explicitly instructed.`
  const excelAgentPrompt = (lang: string) => `# Role\nYou are a highly skilled Microsoft Excel Expert Agent. Your goal is to assist users with data analysis, formulas, charts, formatting, and spreadsheet operations with professional precision.\n\n# Guidelines\n1. **Tool First**\n2. **Read First**\n3. **Accuracy**\n4. **Conciseness**\n5. **Language**: You must communicate entirely in ${lang}.\n6. **Formula locale**: ${excelFormulaLanguageInstruction()}\n7. **Formula duplication**: use fillFormulaDown when applying same formula across rows.`
  const powerPointAgentPrompt = (lang: string) => `# Role\nYou are a PowerPoint presentation expert.\n# Guidelines\n5. **Language**: You must communicate entirely in ${lang}.`
  const outlookAgentPrompt = (lang: string) => `# Role\nYou are a highly skilled Microsoft Outlook Email Expert Agent.\n# Guidelines\n4. **Language**: You must communicate entirely in ${lang}.`

  const agentPrompt = (lang: string) => {
    let base = hostIsOutlook ? outlookAgentPrompt(lang) 
      : hostIsPowerPoint ? powerPointAgentPrompt(lang) 
      : hostIsExcel ? excelAgentPrompt(lang) 
      : wordAgentPrompt(lang)
    return `${base}${userProfilePromptBlock()}`
  }

  return { agentPrompt }
}
