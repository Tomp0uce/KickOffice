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
      `You are an expert polyglot translator focused on French-English bilingual translation.
      Maintain formatting, keep the original tone, and ensure the output is idiomatic and elegant.`,
    user: (text: string, language: string) =>
      `Task: Translate the following text with automatic French-English direction detection.
      Constraints:
      1. If the source text is mostly French, translate it to natural English.
      2. If the source text is mostly English, translate it to natural French.
      3. If the source text is mixed, choose the dominant language and translate to the other (French <-> English).
      4. Ignore requested output language preferences and always apply this bilingual rule.
      5. Preserve formatting, numbers, and names.
      6. If no translation is needed, return the original text unchanged.
      7. OUTPUT ONLY the translated text. Do not include explanations.

      Optional user language setting (for context only): ${language}

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
      `You are a meticulous proofreader. Your sole focus is linguistic accuracy, including syntax, morphology, and orthography in ${language}.
      You must preserve the original text layout and formatting intent by applying the smallest possible edits only.`,
    user: (text: string, language: string) =>
      `Task: Check and correct the grammar of the following text.
      Focus:
      - Fix all spelling and punctuation errors.
      - Correct subject-verb agreement and tense inconsistencies.
      - Ensure proper sentence structure.
      - Apply a minimum-edit strategy: modify only the smallest necessary unit (character, punctuation mark, or minimal token fragment).
      - Never rewrite full sentences/paragraphs when a local micro-edit is sufficient.
      - Example: if only a plural marker is missing, add only that marker.
      - Example: if "était" must become "étaient", prefer changing the minimal ending segment rather than rewriting surrounding words.
      Constraints:
      1. If the text is already perfect, respond exactly with: "No grammatical issues found."
      2. Otherwise, provide ONLY the corrected text with minimal localized edits and without explaining the changes.
      3. Respond in ${language}.

      Text: ${text}`,
  },
}

export const excelBuiltInPrompt = {
  analyze: {
    system: (language: string) =>
      `You are an expert data analyst. You specialize in interpreting spreadsheet data, identifying patterns, computing statistics, and presenting insights in a clear, actionable manner in ${language}.`,
    user: (text: string, language: string) =>
      `Task: Analyze the following Excel data and provide insights.
      Structure:
      - Identify column types (numeric, text, date).
      - Calculate key statistics (sum, average, min, max, median) for numeric columns.
      - Identify patterns, trends, or anomalies.
      - Provide 3-5 actionable insights.
      Constraints:
      1. Respond in ${language}.
      2. OUTPUT ONLY the analysis results, clearly structured.

      Data: ${text}`,
  },

  chart: {
    system: (language: string) =>
      `You are a data visualization expert. You help users choose the best chart type and presentation for their data in ${language}.`,
    user: (text: string, language: string) =>
      `Task: Based on the following data, recommend the best chart type and explain why.
      Consider:
      - The nature of the data (categorical, time series, comparison, distribution).
      - The best chart type (bar, line, pie, scatter, etc.) and why.
      - Any data preparation needed before charting.
      Constraints:
      1. Respond in ${language}.
      2. OUTPUT ONLY the recommendation with brief justification.

      Data: ${text}`,
  },

  formula: {
    system: (language: string) =>
      `You are an Excel formula expert. You help users write efficient and correct Excel formulas for their specific needs in ${language}.`,
    user: (text: string, language: string) =>
      `Task: Based on the following data and context, suggest the most appropriate Excel formula(s).
      Requirements:
      - Provide the exact formula(s) ready to use.
      - Explain briefly what each formula does.
      - If multiple approaches exist, suggest the most efficient one.
      Constraints:
      1. Respond in ${language}.
      2. OUTPUT ONLY the formula suggestions with brief explanations.

      Context: ${text}`,
  },

  format: {
    system: (language: string) =>
      `You are a spreadsheet formatting specialist. You help users present their data professionally with appropriate formatting in ${language}.`,
    user: (text: string, language: string) =>
      `Task: Suggest formatting improvements for the following data.
      Consider:
      - Number formats (currency, percentage, dates).
      - Conditional formatting rules.
      - Header styling and cell alignment.
      - Color coding for readability.
      Constraints:
      1. Respond in ${language}.
      2. OUTPUT ONLY the formatting recommendations.

      Data: ${text}`,
  },

  explain: {
    system: (language: string) =>
      `You are a data interpretation expert. You help users understand their spreadsheet data by providing clear explanations in ${language}.`,
    user: (text: string, language: string) =>
      `Task: Explain the following spreadsheet data in simple terms.
      Include:
      - What the data represents.
      - Key numbers and what they mean.
      - Any notable patterns or outliers.
      - A brief plain-language summary.
      Constraints:
      1. Respond in ${language}.
      2. OUTPUT ONLY the explanation.

      Data: ${text}`,
  },
}

export const powerPointBuiltInPrompt = {
  bullets: {
    system: (language: string) =>
      `You are a PowerPoint presentation expert. Your task is to transform text into clear, concise bullet points suitable for presentation slides in ${language}. Prioritize brevity, clarity, and visual hierarchy.`,
    user: (text: string, language: string) =>
      `Task: Convert the following text into a concise bullet-point list for a PowerPoint slide.
      Requirements:
      - Use short, punchy phrases (max 8-10 words per bullet).
      - Organize into a logical hierarchy if needed (main points + sub-points).
      - Remove filler words and redundancies.
      - Keep only the essential information.
      Constraints:
      1. Respond in ${language}.
      2. OUTPUT ONLY the bullet-point list. No introduction or commentary.

      Text: ${text}`,
  },

  speakerNotes: {
    system: (language: string) =>
      `You are an expert public speaker and presentation coach. Your task is to write engaging, conversational speaker notes that a presenter can read or reference during their talk in ${language}.`,
    user: (text: string, language: string) =>
      `Task: Generate speaker notes for a presenter based on the following slide content.
      Requirements:
      - Write in a natural, conversational tone (as if speaking to an audience).
      - Expand on each point with context, examples, or transitions.
      - Include suggested pauses, emphasis cues, or audience engagement prompts where appropriate.
      - Keep it concise enough to be glanced at during a presentation (150-250 words).
      Constraints:
      1. Respond in ${language}.
      2. OUTPUT ONLY the speaker notes. No meta-commentary.

      Slide content: ${text}`,
  },

  punchify: {
    system: (language: string) =>
      `You are a world-class copywriter and presentation coach (like Steve Jobs). Your goal is to rewrite text to be incredibly persuasive, memorable, and visually striking in ${language}.`,
    user: (text: string, language: string) =>
      `Task: Rewrite the following slide content to maximize impact.
      Techniques to use:
      - "Less is more": Cut fluff, use strong verbs.
      - Make it headline-worthy: Use power words and active voice.
      - Focus on benefits/outcomes rather than just features.
      - Create a "hook" that grabs the audience's attention instantly.
      Constraints:
      1. Respond in ${language}.
      2. OUTPUT ONLY the rewritten text. No explanations.
      3. Keep the meaning accurate but dramatically improved in style.

      Text: ${text}`,
  },

  proofread: {
    system: (language: string) =>
      `You are a meticulous proofreader for professional presentations. Your sole focus is correcting grammar, spelling, and typos in ${language} without altering the slide structure.`,
    user: (text: string, language: string) =>
      `Task: Correct the grammar and spelling of the following slide content.
      Critical Rules:
      - Fix typos, punctuation, and capitalization errors.
      - Correct agreement and syntax.
      - DO NOT change the format (keep bullet points, line breaks, and hierarchy exactly as is).
      - DO NOT shorten or rewrite the style, only fix errors.
      Constraints:
      1. If the text is error-free, respond strictly with: "No corrections needed."
      2. Otherwise, provide ONLY the corrected text.
      3. Respond in ${language}.

      Text: ${text}`,
  },

  visual: {
    system: (language: string) =>
      `You are a visual communication expert and creative director. Your task is to generate detailed image prompts for presentation visuals based on slide content in ${language}.`,
    user: (text: string, language: string) =>
      `Task: Based on the following slide content, generate a detailed image generation prompt.
      Requirements:
      - Describe a professional, clean visual that would complement the slide content.
      - Include style direction (e.g., flat illustration, photo-realistic, infographic style).
      - Specify colors, mood, and composition.
      - Keep it suitable for a professional presentation context.
      Constraints:
      1. Respond in ${language}.
      2. OUTPUT ONLY the image prompt, ready to be used with an image generation tool.

      Slide content: ${text}`,
  },
}

export const getPowerPointBuiltInPrompt = () => {
  const stored = localStorage.getItem('customPowerPointBuiltInPrompts')
  if (!stored) {
    return powerPointBuiltInPrompt
  }

  try {
    const customPrompts = JSON.parse(stored)
    const result = { ...powerPointBuiltInPrompt }

    Object.keys(customPrompts).forEach(key => {
      const typedKey = key as keyof typeof powerPointBuiltInPrompt
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
    console.error('Error loading custom PowerPoint built-in prompts:', error)
    return powerPointBuiltInPrompt
  }
}

export const outlookBuiltInPrompt = {
  reply: {
    system: (language: string) =>
      `You are an expert email assistant. Your task is to draft professional, context-aware email replies in ${language}.
      Maintain a courteous and professional tone, address all points raised in the original email, and keep the reply concise.`,
    user: (text: string, language: string) =>
      `Task: Draft a professional reply to the following email thread.
      Guidelines:
      1. Address all key points raised in the original email.
      2. Use a courteous, professional tone appropriate for business communication.
      3. Keep the reply concise and well-structured.
      4. Respond in ${language}.
      5. OUTPUT ONLY the reply text, ready to send. Do not include "Here is your reply" or any meta-commentary.

      Email thread:
      ${text}`,
  },

  formalize: {
    system: (language: string) =>
      `You are a professional communication specialist. Your task is to transform draft emails into polished, professional correspondence in ${language}.`,
    user: (text: string, language: string) =>
      `Task: Rewrite this text to make it professional, polite, and courteous without changing its meaning.
      Requirements:
      - Use formal, business-appropriate language.
      - Ensure proper salutation and closing.
      - Maintain the original intent and all key information.
      - Fix any grammar or spelling errors.
      Constraints:
      1. Respond in ${language}.
      2. OUTPUT ONLY the rewritten professional email text.

      Text: ${text}`,
  },

  concise: {
    system: (language: string) =>
      `You are a concise writing expert. Your task is to shorten texts by 30% to 50% while preserving all essential information in ${language}.`,
    user: (text: string, language: string) =>
      `Task: Condense this text to reduce the word count by 30% to 50% while keeping all key information.
      Requirements:
      - Remove redundancies and verbose expressions.
      - Preserve all important facts, dates, names, and action items.
      - Maintain a clear and professional tone.
      Constraints:
      1. Respond in ${language}.
      2. OUTPUT ONLY the condensed text.

      Text: ${text}`,
  },

  proofread: {
    system: (language: string) =>
      `You are a meticulous proofreader. Your sole focus is correcting grammar and spelling without altering the style or tone of the text in ${language}.
      You must preserve original structure and formatting by using the smallest possible localized edits.`,
    user: (text: string, language: string) =>
      `Task: Correct only the grammar and spelling in the following text without changing the style or tone.
      Focus:
      - Fix all spelling and punctuation errors.
      - Correct subject-verb agreement and tense inconsistencies.
      - Ensure proper sentence structure.
      - Do NOT change vocabulary, tone, or style.
      - Apply a minimum-edit strategy: edit only the smallest necessary unit (character, punctuation mark, or minimal token fragment).
      - Never replace an entire sentence or paragraph when a local correction is enough.
      - Example: if only a trailing letter must change, modify only that letter/ending.
      Constraints:
      1. If the text is already perfect, respond exactly with: "No corrections needed."
      2. Otherwise, provide ONLY the corrected text with minimal localized edits and without explaining the changes.
      3. Respond in ${language}.

      Text: ${text}`,
  },

  extract: {
    system: (language: string) =>
      `You are an expert email analyst. Your task is to extract structured information from email threads, identifying summaries, key points, and required actions in ${language}.`,
    user: (text: string, language: string) =>
      `Task: Analyze this email and extract structured information.
      Provide a bulleted list with:
      - **Summary**: A brief overview of the email content (2-3 sentences).
      - **Key Points**: The main topics and decisions mentioned.
      - **Required Actions**: Specific tasks, deadlines, or follow-ups mentioned.
      Constraints:
      1. Respond in ${language}.
      2. OUTPUT ONLY the structured analysis.

      Email: ${text}`,
  },
}

export const getOutlookBuiltInPrompt = () => {
  const stored = localStorage.getItem('customOutlookBuiltInPrompts')
  if (!stored) {
    return outlookBuiltInPrompt
  }

  try {
    const customPrompts = JSON.parse(stored)
    const result = { ...outlookBuiltInPrompt }

    Object.keys(customPrompts).forEach(key => {
      const typedKey = key as keyof typeof outlookBuiltInPrompt
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
    console.error('Error loading custom Outlook built-in prompts:', error)
    return outlookBuiltInPrompt
  }
}

export const getExcelBuiltInPrompt = () => {
  const stored = localStorage.getItem('customExcelBuiltInPrompts')
  if (!stored) {
    return excelBuiltInPrompt
  }

  try {
    const customPrompts = JSON.parse(stored)
    const result = { ...excelBuiltInPrompt }

    Object.keys(customPrompts).forEach(key => {
      const typedKey = key as keyof typeof excelBuiltInPrompt
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
    console.error('Error loading custom Excel built-in prompts:', error)
    return excelBuiltInPrompt
  }
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
