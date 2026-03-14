/**
 * Office.js Code Validator
 *
 * Pre-execution validation for eval_* tools.
 * Rejects code that doesn't follow Office.js patterns before it can cause errors.
 */

export type OfficeHost = 'Word' | 'Excel' | 'PowerPoint' | 'Outlook';

export interface ValidationResult {
  valid: boolean;
  errors: string[];
  warnings: string[];
}

/**
 * Properties that require .load() before reading.
 * These are common across all Office hosts.
 */
const LOADABLE_PROPERTIES = [
  'text',
  'values',
  'formulas',
  'items',
  'name',
  'size',
  'bold',
  'italic',
  'color',
  'font',
  'paragraphs',
  'cells',
  'rows',
  'columns',
  'address',
  'rowCount',
  'columnCount',
  'style',
  'format',
  'id',
  'shapes',
  'slides',
];

/**
 * Patterns that indicate reading a property (need load first).
 */
const PROPERTY_READ_PATTERNS = LOADABLE_PROPERTIES.map(
  prop => new RegExp(`\\.${prop}(?!\\s*[=:])\\b`, 'g'), // Match .prop but not .prop = or .prop:
);

/**
 * Validate Office.js code before execution.
 *
 * @param code - The JavaScript code to validate
 * @param host - The Office host (Word, Excel, PowerPoint, Outlook)
 * @returns ValidationResult with errors and warnings
 */
export function validateOfficeCode(code: string, host: OfficeHost): ValidationResult {
  const errors: string[] = [];
  const warnings: string[] = [];

  // ========== CRITICAL ERRORS ==========

  // Rule 1: Must have context.sync() (except for Outlook which uses callbacks)
  if (host !== 'Outlook' && !code.includes('context.sync()')) {
    errors.push(
      'Missing `await context.sync()`. ' +
        'Office.js requires sync() to execute queued operations. ' +
        'Add `await context.sync();` after loading properties and after making changes.',
    );
  }

  // Rule 2: Check for property reads without load() (except for Outlook which uses callbacks)
  if (host !== 'Outlook') {
    const hasLoad = /\.load\s*\(/.test(code);
    let hasPropertyReads = false;

    for (const pattern of PROPERTY_READ_PATTERNS) {
      if (pattern.test(code)) {
        hasPropertyReads = true;
        break;
      }
    }

    if (hasPropertyReads && !hasLoad) {
      errors.push(
        'Reading Office.js properties without `.load()`. ' +
          'Before accessing properties like .text, .values, .items, you must call `.load("propertyName")` then `await context.sync()`. ' +
          'Example: `range.load("text"); await context.sync(); console.log(range.text);`',
      );
    }
  }

  // Rule 3: Host namespace validation
  const namespaceErrors = validateNamespaces(code, host);
  errors.push(...namespaceErrors);

  // Rule 4: Infinite loop detection
  if (/while\s*\(\s*true\s*\)/.test(code)) {
    errors.push('Infinite loop detected: `while(true)` is not allowed.');
  }
  if (/for\s*\(\s*;\s*;\s*\)/.test(code)) {
    errors.push('Infinite loop detected: `for(;;)` is not allowed.');
  }

  // Rule 5: Dangerous operations
  if (/eval\s*\(/.test(code)) {
    errors.push('`eval()` is not allowed inside eval_* tools.');
  }
  if (/Function\s*\(/.test(code)) {
    errors.push('`new Function()` is not allowed.');
  }

  // ========== WARNINGS ==========

  // Warning 1: No try/catch
  if (!code.includes('try') || !code.includes('catch')) {
    warnings.push(
      'No try/catch block detected. ' +
        'Wrap your code in try/catch to handle Office.js errors gracefully.',
    );
  }

  // Warning 2: Multiple syncs without batching
  const syncCount = (code.match(/context\.sync\(\)/g) || []).length;
  if (syncCount > 3) {
    warnings.push(
      `Found ${syncCount} context.sync() calls. ` +
        'Consider batching operations to reduce round-trips to Office.',
    );
  }

  // Warning 3: Direct property assignment without checking
  if (/\.values\s*=\s*[^[{]/.test(code)) {
    warnings.push(
      'Direct assignment to .values detected. ' +
        'Remember Excel values must be 2D arrays: `range.values = [[value]]`',
    );
  }

  // Warning 4: getRange with large hardcoded range
  if (/getRange\s*\(\s*['"`][A-Z]+1?\s*:\s*[A-Z]+\d{4,}/.test(code)) {
    warnings.push(
      'Large hardcoded range detected. ' +
        'Consider using `getUsedRange()` instead for better performance.',
    );
  }

  return {
    valid: errors.length === 0,
    errors,
    warnings,
  };
}

/**
 * Validate that code only uses the correct host namespace.
 */
function validateNamespaces(code: string, host: OfficeHost): string[] {
  const errors: string[] = [];

  // Define which namespaces are allowed for each host
  const allowedNamespaces: Record<OfficeHost, string[]> = {
    Word: ['Word', 'Office'],
    Excel: ['Excel', 'Office'],
    PowerPoint: ['PowerPoint', 'Office'],
    Outlook: ['Office'], // Outlook uses Office.context.mailbox
  };

  const allHostNamespaces = ['Word', 'Excel', 'PowerPoint'];
  const allowed = allowedNamespaces[host];

  for (const ns of allHostNamespaces) {
    // Check if namespace is used (e.g., "Word." or "Word.run")
    const nsPattern = new RegExp(`\\b${ns}\\s*\\.`, 'g');
    if (nsPattern.test(code) && !allowed.includes(ns)) {
      errors.push(
        `Invalid namespace: Cannot use \`${ns}\` APIs in ${host} context. ` +
          `You are running in ${host} — only ${allowed.join(', ')} APIs are available.`,
      );
    }
  }

  return errors;
}

/**
 * Format validation result for display to the AI.
 */
export function formatValidationResult(result: ValidationResult): string {
  if (result.valid && result.warnings.length === 0) {
    return 'Code validation passed.';
  }

  let output = '';

  if (result.errors.length > 0) {
    output += '## Validation ERRORS (must fix):\n';
    result.errors.forEach((error, i) => {
      output += `${i + 1}. ${error}\n`;
    });
    output += '\n';
  }

  if (result.warnings.length > 0) {
    output += '## Validation WARNINGS (recommended to fix):\n';
    result.warnings.forEach((warning, i) => {
      output += `${i + 1}. ${warning}\n`;
    });
  }

  return output;
}

/**
 * Quick check if code is likely valid (for fast-path).
 */
export function quickValidate(code: string): boolean {
  return (
    code.includes('context.sync()') &&
    (code.includes('.load(') || !PROPERTY_READ_PATTERNS.some(p => p.test(code)))
  );
}
