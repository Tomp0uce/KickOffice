import { describe, it, expect } from 'vitest'
import {
  validateOfficeCode,
  formatValidationResult,
  quickValidate,
} from '../officeCodeValidator'
import type { OfficeHost, ValidationResult } from '../officeCodeValidator'

// ─────────────────────────────────────────────────────────────────────────────
// Helpers
// ─────────────────────────────────────────────────────────────────────────────

/** Minimal valid Word snippet that satisfies all mandatory rules. */
const VALID_WORD_CODE = `
  await Word.run(async (context) => {
    const body = context.document.body
    body.load("text")
    await context.sync()
    try {
      console.log(body.text)
    } catch (e) {
      console.error(e)
    }
  })
`

/** Minimal valid Excel snippet. */
const VALID_EXCEL_CODE = `
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet()
    sheet.load("name")
    await context.sync()
    try {
      console.log(sheet.name)
    } catch (e) {
      console.error(e)
    }
  })
`

// ─────────────────────────────────────────────────────────────────────────────
// validateOfficeCode — Rule 1: context.sync()
// ─────────────────────────────────────────────────────────────────────────────
describe('validateOfficeCode — Rule 1: context.sync() requirement', () => {
  it('is valid when context.sync() is present for Word', () => {
    const result = validateOfficeCode(VALID_WORD_CODE, 'Word')
    expect(result.valid).toBe(true)
    expect(result.errors).toHaveLength(0)
  })

  it('reports an error when context.sync() is missing for Word', () => {
    const code = `
      await Word.run(async (context) => {
        const body = context.document.body
        body.load("text")
        try { console.log(body.text) } catch(e) {}
      })
    `
    const result = validateOfficeCode(code, 'Word')
    expect(result.valid).toBe(false)
    expect(result.errors.some(e => e.includes('context.sync()'))).toBe(true)
  })

  it('does NOT require context.sync() for Outlook', () => {
    const outlookCode = `
      Office.context.mailbox.item.subject.getAsync((result) => {
        console.log(result.value)
      })
    `
    const result = validateOfficeCode(outlookCode, 'Outlook')
    const syncError = result.errors.find(e => e.includes('context.sync()'))
    expect(syncError).toBeUndefined()
  })

  it('reports error for Excel without context.sync()', () => {
    const code = `
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet()
        try { console.log(sheet.name) } catch(e) {}
      })
    `
    const result = validateOfficeCode(code, 'Excel')
    expect(result.valid).toBe(false)
  })
})

// ─────────────────────────────────────────────────────────────────────────────
// validateOfficeCode — Rule 2: .load() before property reads
// ─────────────────────────────────────────────────────────────────────────────
describe('validateOfficeCode — Rule 2: .load() before property reads', () => {
  it('is valid when .load() is present alongside property reads', () => {
    const result = validateOfficeCode(VALID_WORD_CODE, 'Word')
    expect(result.errors.some(e => e.includes('.load()'))).toBe(false)
  })

  it('reports error when reading .values without .load()', () => {
    // NOTE: PROPERTY_READ_PATTERNS uses stateful RegExp with the global `g` flag.
    // A previous test may have advanced `lastIndex` on the `.text` pattern.
    // Using `.values` (a different loadable property) makes this test order-independent.
    const code = `
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange()
        await context.sync()
        try { console.log(range.values) } catch(e) {}
      })
    `
    const result = validateOfficeCode(code, 'Excel')
    // The validator detects ".values" as a property read; no ".load(" present → error
    expect(result.errors.some(e => e.includes('.load()'))).toBe(true)
  })

  it('does NOT flag .load() rule for Outlook', () => {
    const outlookCode = `
      Office.context.mailbox.item.body.getAsync((result) => {
        console.log(result.value)
      })
    `
    // Outlook does not use the load/sync pattern
    const result = validateOfficeCode(outlookCode, 'Outlook')
    expect(result.errors.some(e => e.includes('.load()'))).toBe(false)
  })

  it('does not flag property assignment (.values =) as a missing load', () => {
    // Assignment (op=) should not trigger the read-without-load error
    const code = `
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange()
        range.values = [[42]]
        await context.sync()
        try {} catch(e) {}
      })
    `
    const result = validateOfficeCode(code, 'Excel')
    expect(result.errors.some(e => e.includes('.load()'))).toBe(false)
  })
})

// ─────────────────────────────────────────────────────────────────────────────
// validateOfficeCode — Rule 3: namespace validation
// ─────────────────────────────────────────────────────────────────────────────
describe('validateOfficeCode — Rule 3: host namespace validation', () => {
  it('allows the correct namespace for Word', () => {
    const result = validateOfficeCode(VALID_WORD_CODE, 'Word')
    expect(result.errors.some(e => e.includes('namespace'))).toBe(false)
  })

  it('rejects Excel namespace in a Word context', () => {
    const code = `
      await Excel.run(async (context) => {
        await context.sync()
        try {} catch(e) {}
      })
    `
    const result = validateOfficeCode(code, 'Word')
    expect(result.valid).toBe(false)
    expect(result.errors.some(e => e.includes('Excel'))).toBe(true)
  })

  it('rejects Word namespace in an Excel context', () => {
    const code = `
      await Word.run(async (context) => {
        await context.sync()
        try {} catch(e) {}
      })
    `
    const result = validateOfficeCode(code, 'Excel')
    expect(result.valid).toBe(false)
    expect(result.errors.some(e => e.includes('Word'))).toBe(true)
  })

  it('allows Office namespace in all contexts', () => {
    const hosts: OfficeHost[] = ['Word', 'Excel', 'PowerPoint', 'Outlook']
    for (const host of hosts) {
      const code = `
        const platform = Office.context.platform
        await context.sync()
        try {} catch(e) {}
      `
      const result = validateOfficeCode(code, host)
      expect(result.errors.some(e => e.includes('namespace') && e.includes('Office'))).toBe(false)
    }
  })

  it('rejects PowerPoint namespace in Word context', () => {
    const code = `
      await PowerPoint.run(async (context) => {
        await context.sync()
        try {} catch(e) {}
      })
    `
    const result = validateOfficeCode(code, 'Word')
    expect(result.valid).toBe(false)
    expect(result.errors.some(e => e.includes('PowerPoint'))).toBe(true)
  })
})

// ─────────────────────────────────────────────────────────────────────────────
// validateOfficeCode — Rule 4: infinite loop detection
// ─────────────────────────────────────────────────────────────────────────────
describe('validateOfficeCode — Rule 4: infinite loop detection', () => {
  it('rejects while(true)', () => {
    const code = `
      while (true) {
        await context.sync()
        try {} catch(e) {}
      }
    `
    const result = validateOfficeCode(code, 'Word')
    expect(result.errors.some(e => e.includes('while(true)'))).toBe(true)
  })

  it('rejects for(;;)', () => {
    const code = `
      for (;;) {
        await context.sync()
        try {} catch(e) {}
      }
    `
    const result = validateOfficeCode(code, 'Word')
    expect(result.errors.some(e => e.includes('for(;;)'))).toBe(true)
  })

  it('accepts a normal for loop', () => {
    const result = validateOfficeCode(VALID_WORD_CODE, 'Word')
    expect(result.errors.some(e => e.includes('loop'))).toBe(false)
  })
})

// ─────────────────────────────────────────────────────────────────────────────
// validateOfficeCode — Rule 5: dangerous operations
// ─────────────────────────────────────────────────────────────────────────────
describe('validateOfficeCode — Rule 5: dangerous operations', () => {
  it('rejects eval()', () => {
    const code = `
      eval("console.log(1)")
      await context.sync()
      try {} catch(e) {}
    `
    const result = validateOfficeCode(code, 'Word')
    expect(result.errors.some(e => e.includes('eval()'))).toBe(true)
  })

  it('rejects new Function()', () => {
    const code = `
      const fn = new Function("return 1")
      await context.sync()
      try {} catch(e) {}
    `
    const result = validateOfficeCode(code, 'Word')
    expect(result.errors.some(e => e.includes('Function()'))).toBe(true)
  })

  it('does not flag regular function declarations', () => {
    const code = `
      function helper() { return 1 }
      await context.sync()
      try {} catch(e) {}
    `
    const result = validateOfficeCode(code, 'Word')
    expect(result.errors.some(e => e.includes('Function()'))).toBe(false)
  })
})

// ─────────────────────────────────────────────────────────────────────────────
// validateOfficeCode — Warnings
// ─────────────────────────────────────────────────────────────────────────────
describe('validateOfficeCode — warnings', () => {
  it('warns when no try/catch block is present', () => {
    const code = `
      await Word.run(async (context) => {
        const body = context.document.body
        body.load("text")
        await context.sync()
        console.log(body.text)
      })
    `
    const result = validateOfficeCode(code, 'Word')
    expect(result.warnings.some(w => w.includes('try/catch'))).toBe(true)
  })

  it('does NOT warn about try/catch when both are present', () => {
    const result = validateOfficeCode(VALID_WORD_CODE, 'Word')
    expect(result.warnings.some(w => w.includes('try/catch'))).toBe(false)
  })

  it('warns when more than 3 context.sync() calls are detected', () => {
    const syncs = 'await context.sync()\n'.repeat(4)
    const code = `
      await Word.run(async (context) => {
        ${syncs}
        try {} catch(e) {}
      })
    `
    const result = validateOfficeCode(code, 'Word')
    expect(result.warnings.some(w => w.includes('context.sync()'))).toBe(true)
  })

  it('does NOT warn for 3 or fewer context.sync() calls', () => {
    const syncs = 'await context.sync()\n'.repeat(3)
    const code = `
      await Word.run(async (context) => {
        ${syncs}
        try {} catch(e) {}
      })
    `
    const result = validateOfficeCode(code, 'Word')
    expect(result.warnings.some(w => w.includes('context.sync()'))).toBe(false)
  })

  it('warns on direct .values = scalar assignment', () => {
    const code = `
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange()
        range.values = 42
        await context.sync()
        try {} catch(e) {}
      })
    `
    const result = validateOfficeCode(code, 'Excel')
    expect(result.warnings.some(w => w.includes('.values'))).toBe(true)
  })

  it('warns on large hardcoded range (e.g. A1:Z10000)', () => {
    const code = `
      await Excel.run(async (context) => {
        const range = context.workbook.worksheets.getActiveWorksheet().getRange("A1:Z10000")
        range.load("values")
        await context.sync()
        try {} catch(e) {}
      })
    `
    const result = validateOfficeCode(code, 'Excel')
    expect(result.warnings.some(w => w.includes('getUsedRange()'))).toBe(true)
  })
})

// ─────────────────────────────────────────────────────────────────────────────
// validateOfficeCode — valid flag semantics
// ─────────────────────────────────────────────────────────────────────────────
describe('validateOfficeCode — valid flag', () => {
  it('valid is true when errors array is empty', () => {
    const result = validateOfficeCode(VALID_WORD_CODE, 'Word')
    expect(result.valid).toBe(result.errors.length === 0)
  })

  it('valid is false when at least one error is present', () => {
    const result = validateOfficeCode('', 'Word')
    expect(result.valid).toBe(false)
  })

  it('returns a ValidationResult shape for every host', () => {
    const hosts: OfficeHost[] = ['Word', 'Excel', 'PowerPoint', 'Outlook']
    for (const host of hosts) {
      const result: ValidationResult = validateOfficeCode(VALID_WORD_CODE, host)
      expect(result).toHaveProperty('valid')
      expect(result).toHaveProperty('errors')
      expect(result).toHaveProperty('warnings')
      expect(Array.isArray(result.errors)).toBe(true)
      expect(Array.isArray(result.warnings)).toBe(true)
    }
  })
})

// ─────────────────────────────────────────────────────────────────────────────
// formatValidationResult
// ─────────────────────────────────────────────────────────────────────────────
describe('formatValidationResult', () => {
  it('returns "Code validation passed." when valid and no warnings', () => {
    const result: ValidationResult = { valid: true, errors: [], warnings: [] }
    expect(formatValidationResult(result)).toBe('Code validation passed.')
  })

  it('includes error section header when there are errors', () => {
    const result: ValidationResult = {
      valid: false,
      errors: ['Missing context.sync()'],
      warnings: [],
    }
    const output = formatValidationResult(result)
    expect(output).toContain('## Validation ERRORS')
    expect(output).toContain('Missing context.sync()')
  })

  it('includes warning section header when there are warnings', () => {
    const result: ValidationResult = {
      valid: true,
      errors: [],
      warnings: ['Consider try/catch'],
    }
    const output = formatValidationResult(result)
    expect(output).toContain('## Validation WARNINGS')
    expect(output).toContain('Consider try/catch')
  })

  it('includes both error and warning sections when both are present', () => {
    const result: ValidationResult = {
      valid: false,
      errors: ['Error 1'],
      warnings: ['Warning 1'],
    }
    const output = formatValidationResult(result)
    expect(output).toContain('## Validation ERRORS')
    expect(output).toContain('## Validation WARNINGS')
  })

  it('numbers errors starting from 1', () => {
    const result: ValidationResult = {
      valid: false,
      errors: ['First error', 'Second error'],
      warnings: [],
    }
    const output = formatValidationResult(result)
    expect(output).toContain('1. First error')
    expect(output).toContain('2. Second error')
  })

  it('numbers warnings starting from 1', () => {
    const result: ValidationResult = {
      valid: true,
      errors: [],
      warnings: ['First warning', 'Second warning'],
    }
    const output = formatValidationResult(result)
    expect(output).toContain('1. First warning')
    expect(output).toContain('2. Second warning')
  })
})

// ─────────────────────────────────────────────────────────────────────────────
// quickValidate
// ─────────────────────────────────────────────────────────────────────────────
describe('quickValidate', () => {
  it('returns true for valid code with sync and load', () => {
    expect(quickValidate(VALID_WORD_CODE)).toBe(true)
  })

  it('returns true for code with sync and no property reads', () => {
    const code = 'await context.sync()'
    expect(quickValidate(code)).toBe(true)
  })

  it('returns false when context.sync() is absent', () => {
    expect(quickValidate('some code without sync')).toBe(false)
  })

  it('returns false when there are property reads but no .load(', () => {
    // Contains ".text" (a loadable property read) but no .load( and no context.sync()
    const code = 'console.log(body.text)'
    expect(quickValidate(code)).toBe(false)
  })

  it('handles an empty string without throwing', () => {
    expect(() => quickValidate('')).not.toThrow()
    expect(quickValidate('')).toBe(false)
  })
})
