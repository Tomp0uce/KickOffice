export const languageMap: IStringKeyMap = {
  en: 'English',
  es: 'Espa\u00f1ol',
  fr: 'Fran\u00e7ais',
  de: 'Deutsch',
  it: 'Italiano',
  pt: 'Portugu\u00eas',
  'zh-cn': '\u7b80\u4f53\u4e2d\u6587',
  ja: '\u65e5\u672c\u8a9e',
  ko: '\ud55c\uad6d\uc5b4',
  nl: 'Nederlands',
  pl: 'Polski',
  ar: '\u0627\u0644\u0639\u0631\u0628\u064a\u0629',
  ru: '\u0420\u0443\u0441\u0441\u043a\u0438\u0439',
}

export const buildInPrompt = {
  translate: {
    system: (language: string) =>
      `You are an expert polyglot translator. Your task is to provide professional, context-aware translations into ${language}.
      Maintain formatting, keep the original tone, and ensure the output is idiomatic and elegant.`,
    user: (text: string, language: string) =>
      `Task: Translate the following text into ${language}.
      Constraints:
      1. Provide a natural-sounding translation suitable for native speakers.
      2. If the text is technical, use appropriate terminology.
      3. OUTPUT ONLY the translated text. Do not include "Here is the translation" or any explanations.

      Text: ${text}`,
  },

  polish: {
    system: (language: string) =>
      `You are a professional editor and stylist. Your goal is to make the text more professional, engaging, and clear in ${language}.`,
    user: (text: string, language: string) =>
      `Task: Polish the following text for better flow and impact.
      Improvements:
      - Correct grammar, spelling, and punctuation.
      - Enhance vocabulary while maintaining the original meaning.
      - Improve sentence structure and eliminate redundancy.
      - Ensure the tone is consistent and professional.
      Constraints:
      1. Respond in ${language}.
      2. OUTPUT ONLY the polished text without any commentary.

      Text: ${text}`,
  },

  academic: {
    system: (language: string) =>
      `You are a senior academic editor for high-impact journals. You specialize in formal, precise, and objective scholarly writing in ${language}.`,
    user: (text: string, language: string) =>
      `Task: Rewrite the following text to meet professional academic standards.
      Requirements:
      - Use formal, objective language and avoid colloquialisms.
      - Ensure logical transitions and precise scientific terminology.
      - Maintain a third-person perspective unless the context requires otherwise.
      - Optimize for clarity and conciseness as per peer-review expectations.
      Constraints:
      1. Respond in ${language}.
      2. OUTPUT ONLY the revised text. No pre-amble or meta-talk.

      Text: ${text}`,
  },

  summary: {
    system: (language: string) =>
      `You are an expert document analyst. You excel at distilling complex information into clear, actionable summaries in ${language}.`,
    user: (text: string, language: string) =>
      `Task: Summarize the following text.
      Structure:
      - Capture the core message and primary supporting points.
      - Aim for approximately 100 words (or 3-5 key bullet points).
      - Ensure the summary is self-contained and easy to understand.
      Constraints:
      1. Respond in ${language}.
      2. OUTPUT ONLY the summary.

      Text: ${text}`,
  },

  grammar: {
    system: (language: string) =>
      `You are a meticulous proofreader. Your sole focus is linguistic accuracy, including syntax, morphology, and orthography in ${language}.`,
    user: (text: string, language: string) =>
      `Task: Check and correct the grammar of the following text.
      Focus:
      - Fix all spelling and punctuation errors.
      - Correct subject-verb agreement and tense inconsistencies.
      - Ensure proper sentence structure.
      Constraints:
      1. If the text is already perfect, respond exactly with: "No grammatical issues found."
      2. Otherwise, provide ONLY the corrected text without explaining the changes.
      3. Respond in ${language}.

      Text: ${text}`,
  },
}

export const getBuiltInPrompt = () => {
  const stored = localStorage.getItem('customBuiltInPrompts')
  if (!stored) {
    return buildInPrompt
  }

  try {
    const customPrompts = JSON.parse(stored)
    const result = { ...buildInPrompt }

    Object.keys(customPrompts).forEach(key => {
      const typedKey = key as keyof typeof buildInPrompt
      if (result[typedKey]) {
        result[typedKey] = {
          system: (language: string) => customPrompts[key].system.replace(/\$\{language\}/g, language),
          user: (text: string, language: string) =>
            customPrompts[key].user.replace(/\$\{text\}/g, text).replace(/\$\{language\}/g, language),
        }
      }
    })

    return result
  } catch (error) {
    console.error('Error loading custom built-in prompts:', error)
    return buildInPrompt
  }
}
