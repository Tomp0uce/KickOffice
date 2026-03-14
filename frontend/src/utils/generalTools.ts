import type { ToolInputSchema, ToolCategory } from '@/types';
import { getErrorMessage } from './common';
import { evaluate } from 'mathjs';
import { getBash, writeFile as vfsWrite, readFile as vfsRead, listUploads } from '@/utils/vfs';

export type GeneralToolName =
  | 'getCurrentDate'
  | 'calculateMath'
  | 'executeBash'
  | 'vfsWriteFile'
  | 'vfsReadFile'
  | 'vfsListFiles';

export interface GeneralToolDefinition {
  name: GeneralToolName;
  category: ToolCategory;
  description: string;
  inputSchema: ToolInputSchema;
  execute: (args: Record<string, any>) => Promise<string>;
}

const generalToolDefinitions: GeneralToolDefinition[] = [
  {
    name: 'getCurrentDate',
    category: 'read',
    description:
      'Returns the current date and time. Useful for adding timestamps, dates to documents, or understanding temporal context.',
    inputSchema: {
      type: 'object',
      properties: {
        format: {
          type: 'string',
          description:
            'Format: "full" (date and time), "date" (date only), "time" (time only), "iso" (ISO 8601)',
          enum: ['full', 'date', 'time', 'iso'],
        },
      },
      required: [],
    },
    execute: async args => {
      const { format = 'full' } = args;
      const now = new Date();

      switch (format) {
        case 'date':
          return now.toLocaleDateString('en-US', {
            year: 'numeric',
            month: 'long',
            day: 'numeric',
          });
        case 'time':
          return now.toLocaleTimeString('en-US', {
            hour: '2-digit',
            minute: '2-digit',
            second: '2-digit',
          });
        case 'iso':
          return now.toISOString();
        case 'full':
        default:
          return now.toLocaleString('en-US', {
            year: 'numeric',
            month: 'long',
            day: 'numeric',
            hour: '2-digit',
            minute: '2-digit',
            second: '2-digit',
          });
      }
    },
  },
  {
    name: 'calculateMath',
    category: 'write',
    description:
      'Evaluates mathematical expressions safely. Supports basic arithmetic (+, -, *, /), parentheses, and common math functions.',
    inputSchema: {
      type: 'object',
      properties: {
        expression: {
          type: 'string',
          description: 'The mathematical expression to evaluate (e.g., "2 + 2 * 3")',
        },
      },
      required: ['expression'],
    },
    execute: async args => {
      const { expression } = args as Record<string, any>;
      try {
        const result = evaluate(expression);

        if (typeof result !== 'number' && typeof result !== 'bigint') {
          return `Calculation completed, but result is not a simple number: ${result}`;
        }

        return `${expression} = ${result}`;
      } catch (error: unknown) {
        return `Error evaluating expression: ${getErrorMessage(error)}`;
      }
    },
  },
  {
    name: 'executeBash',
    category: 'write',
    description:
      'Execute a bash command in a sandboxed in-memory shell (VFS). Use this for data processing, scripting, and shell tasks. The shell is stateful within a session. You can write custom reusable bash functions to /home/user/scripts/ (using vfsWriteFile) and call them here. Available utilities include: ls, cat, grep, awk, sed, find, sort, uniq, wc, cut, head, tail, and base64.',
    inputSchema: {
      type: 'object',
      properties: {
        command: {
          type: 'string',
          description:
            'The bash command to execute. Examples: "ls /home/user/uploads", "cat data.txt | grep error", "source /home/user/scripts/my_tool.sh && my_tool arg"',
        },
      },
      required: ['command'],
    },
    execute: async args => {
      const { command } = args as { command: string };
      try {
        const shell = getBash();
        const result = await shell.exec(command);
        const output = [result.stdout, result.stderr].filter(Boolean).join('\n');
        return output || '(no output)';
      } catch (error: unknown) {
        return `Shell error: ${getErrorMessage(error)}`;
      }
    },
  },
  {
    name: 'vfsWriteFile',
    category: 'write',
    description:
      'Write text content to a file in the sandboxed virtual filesystem (VFS). Best used to create custom tools by writing reusable bash scripts to /home/user/scripts/<name>.sh, or processing output to /home/user/uploads/.',
    inputSchema: {
      type: 'object',
      properties: {
        path: {
          type: 'string',
          description: 'File path (e.g., "/home/user/scripts/extract.sh" or "output.txt")',
        },
        content: {
          type: 'string',
          description: 'Text content to write to the file',
        },
      },
      required: ['path', 'content'],
    },
    execute: async args => {
      const { path, content } = args as { path: string; content: string };
      try {
        await vfsWrite(path, content);
        const fullPath = path.startsWith('/') ? path : `/home/user/uploads/${path}`;
        return `File written successfully to ${fullPath}`;
      } catch (error: unknown) {
        return `Error writing file: ${getErrorMessage(error)}`;
      }
    },
  },
  {
    name: 'vfsReadFile',
    category: 'read',
    description: 'Read text content from a file in the sandboxed virtual filesystem (VFS).',
    inputSchema: {
      type: 'object',
      properties: {
        path: {
          type: 'string',
          description: 'File path to read (e.g., "data.csv" or "/home/user/uploads/data.csv")',
        },
      },
      required: ['path'],
    },
    execute: async args => {
      const { path } = args as { path: string };
      try {
        return await vfsRead(path);
      } catch (error: unknown) {
        return `Error reading file: ${getErrorMessage(error)}`;
      }
    },
  },
  {
    name: 'vfsListFiles',
    category: 'read',
    description: 'List all files available in the sandboxed virtual filesystem uploads directory.',
    inputSchema: {
      type: 'object',
      properties: {},
      required: [],
    },
    execute: async () => {
      try {
        const files = await listUploads();
        if (files.length === 0) return 'No files in VFS uploads directory.';
        return `Files in /home/user/uploads/:\n${files.map(f => `  - ${f}`).join('\n')}`;
      } catch (error: unknown) {
        return `Error listing files: ${getErrorMessage(error)}`;
      }
    },
  },
];

export function getGeneralToolDefinitions(): GeneralToolDefinition[] {
  return generalToolDefinitions;
}
