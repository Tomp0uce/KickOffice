import type { ToolDefinition } from '@/types';
import { logService } from '@/utils/logger';
import { executeOfficeAction } from './officeAction';
import {
  createOfficeTools,
  buildExecuteWrapper,
  type OfficeToolTemplate,
  getErrorMessage,
  createEvalExecutor,
  buildScreenshotResult,
} from './common';
import { localStorageKey } from './enum';
import { createMutationDetector } from './mutationDetector';
import { extractChartData } from '@/api/backend';
import { getVfs, getVfsSandboxContext } from '@/utils/vfs';

const runExcel = <T>(action: (context: Excel.RequestContext) => Promise<T>): Promise<T> =>
  executeOfficeAction(() => Excel.run(action));

// ============================================================
// Screenshot headers composition — ported from Office Agents
// packages/excel/src/lib/tools/screenshot-range.ts
// ============================================================
const HEADER_WIDTH = 40;
const HEADER_HEIGHT = 20;
const HEADER_BG = '#f0f0f0';
const HEADER_BORDER = '#c0c0c0';
const HEADER_FONT = 'bold 11px Calibri, Arial, sans-serif';
const HEADER_TEXT_COLOR = '#333333';

/** Convert 0-based column index to Excel column letter (0→A, 1→B, 26→AA). Ported from Office Agents. */
function columnIndexToLetter(index: number): string {
  let letter = '';
  let temp = index;
  while (temp >= 0) {
    letter = String.fromCharCode((temp % 26) + 65) + letter;
    temp = Math.floor(temp / 26) - 1;
  }
  return letter;
}

/** Parse start row/col from an A1-style range address (e.g. "Sheet1!B3:D10" → {startRow:2, startCol:1}). */
function parseRangeStart(rangeAddress: string): { startRow: number; startCol: number } {
  const addr = rangeAddress.includes('!') ? rangeAddress.split('!')[1] : rangeAddress;
  const startCell = addr.split(':')[0];
  const match = startCell.match(/^([A-Z]+)(\d+)$/i);
  if (!match) return { startRow: 0, startCol: 0 };
  const colStr = match[1].toUpperCase();
  let col = 0;
  for (let i = 0; i < colStr.length; i++) {
    col = col * 26 + (colStr.charCodeAt(i) - 64);
  }
  return { startRow: parseInt(match[2], 10) - 1, startCol: col - 1 };
}

/**
 * Composite an Excel range screenshot with row/column headers using Canvas.
 * Ported from Office Agents screenshot-range.ts
 */
function compositeWithHeaders(
  imageBase64: string,
  startRow: number,
  startCol: number,
  colWidths: number[],
  rowHeights: number[],
): Promise<string> {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => {
      const totalColWidth = colWidths.reduce((a, b) => a + b, 0);
      const totalRowHeight = rowHeights.reduce((a, b) => a + b, 0);
      const scaleX = totalColWidth > 0 ? img.width / totalColWidth : 1;
      const scaleY = totalRowHeight > 0 ? img.height / totalRowHeight : 1;

      const canvas = document.createElement('canvas');
      canvas.width = HEADER_WIDTH + img.width;
      canvas.height = HEADER_HEIGHT + img.height;
      const ctx = canvas.getContext('2d');
      if (!ctx) return reject(new Error('Failed to get 2d canvas context'));

      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, canvas.width, canvas.height);
      ctx.drawImage(img, HEADER_WIDTH, HEADER_HEIGHT);

      // Column headers
      ctx.fillStyle = HEADER_BG;
      ctx.fillRect(HEADER_WIDTH, 0, img.width, HEADER_HEIGHT);
      ctx.font = HEADER_FONT;
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      let x = HEADER_WIDTH;
      for (let i = 0; i < colWidths.length; i++) {
        const w = colWidths[i] * scaleX;
        ctx.strokeStyle = HEADER_BORDER;
        ctx.strokeRect(x, 0, w, HEADER_HEIGHT);
        ctx.fillStyle = HEADER_TEXT_COLOR;
        ctx.fillText(columnIndexToLetter(startCol + i), x + w / 2, HEADER_HEIGHT / 2);
        x += w;
      }

      // Row headers
      ctx.fillStyle = HEADER_BG;
      ctx.fillRect(0, HEADER_HEIGHT, HEADER_WIDTH, img.height);
      let y = HEADER_HEIGHT;
      for (let i = 0; i < rowHeights.length; i++) {
        const h = rowHeights[i] * scaleY;
        ctx.strokeStyle = HEADER_BORDER;
        ctx.strokeRect(0, y, HEADER_WIDTH, h);
        ctx.fillStyle = HEADER_TEXT_COLOR;
        ctx.fillText(String(startRow + i + 1), HEADER_WIDTH / 2, y + h / 2);
        y += h;
      }

      // Top-left corner cell
      ctx.fillStyle = HEADER_BG;
      ctx.fillRect(0, 0, HEADER_WIDTH, HEADER_HEIGHT);
      ctx.strokeStyle = HEADER_BORDER;
      ctx.strokeRect(0, 0, HEADER_WIDTH, HEADER_HEIGHT);

      resolve(canvas.toDataURL('image/png').split(',')[1]);
    };
    img.onerror = () => reject(new Error('Failed to load range image for header composition'));
    img.src = `data:image/png;base64,${imageBase64}`;
  });
}

// ============================================================
// Mutation detection patterns — ported from Office Agents dirty-tracker
// packages/excel/src/lib/dirty-tracker.ts
// ============================================================
/** Detect if code contains write operations. Ported from Office Agents. */
const looksLikeMutation = createMutationDetector([
  /\.(values|formulas|formulasLocal|numberFormat|numberFormatLocal)\s*=/,
  /\.clear\s*\(/,
  /\.delete\s*\(/,
  /\.insert\s*\(/,
  /\.copyFrom\s*\(/,
  /\.add\s*\(/,
  /\.merge\s*\(/,
  /\.unmerge\s*\(/,
  /\.format\.\w+\s*=/,
  /\.set\s*\(/,
]);

// Agent-modified cells are marked with text underline (font.underline = single).
// This is auto-applied by setCellRange and cleared by clearAgentHighlights.

/** Coerce a CSV string value to its native type. Ported from Office Agents csv-to-sheet. */
function coerceValue(value: string): string | number | boolean {
  if (value === '') return '';
  const lower = value.toLowerCase();
  if (lower === 'true') return true;
  if (lower === 'false') return false;
  const num = Number(value);
  if (!isNaN(num) && value.trim() !== '') return num;
  return value;
}

/** Safely resolves a worksheet by name, falling back to active sheet. Throws a clear error if the sheet doesn't exist. */
async function safeGetSheet(
  context: Excel.RequestContext,
  sheetName?: string,
): Promise<Excel.Worksheet> {
  if (!sheetName) return context.workbook.worksheets.getActiveWorksheet();
  const sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
  await context.sync();
  if (sheet.isNullObject) {
    throw new Error(
      `Worksheet "${sheetName}" not found. Use getWorksheetInfo to list available worksheets.`,
    );
  }
  return sheet;
}

type ExcelToolTemplate = OfficeToolTemplate<Excel.RequestContext> & {
  executeExcel: (context: Excel.RequestContext, args: Record<string, any>) => Promise<string>;
};

export type ExcelToolName =
  | 'getSelectedCells'
  | 'setCellRange'
  | 'getWorksheetData'
  | 'createTable'
  | 'modifyStructure'
  | 'formatRange'
  | 'sortRange'
  | 'getWorksheetInfo'
  | 'clearRange'
  | 'searchAndReplace'
  | 'addWorksheet'
  | 'getNamedRanges'
  | 'applyConditionalFormatting'
  | 'findData'
  | 'getAllObjects'
  | 'manageObject'
  | 'protectWorksheet'
  | 'setNamedRange'
  | 'getConditionalFormattingRules'
  | 'eval_officejs'
  | 'screenshotRange'
  | 'getRangeAsCsv'
  | 'modifyWorkbookStructure'
  | 'detectDataHeaders'
  | 'importCsvToSheet'
  | 'clearAgentHighlights'
  | 'imageToSheet';

function getExcelFormulaLanguage(): 'en' | 'fr' {
  const configured = localStorage.getItem(localStorageKey.excelFormulaLanguage);
  return configured === 'fr' ? 'fr' : 'en';
}

const excelToolDefinitions = createOfficeTools<ExcelToolName, ExcelToolTemplate, ToolDefinition>(
  {
    getSelectedCells: {
      name: 'getSelectedCells',
      category: 'read',
      description:
        'Get the values, formulas, address, and dimensions of the currently selected cells in Excel. Returns a JSON object with address, rowCount, columnCount, the 2D values array, and the 2D formulas array (cells without formulas show their value as-is in the formulas array). Always use formulas (not values) when explaining how a cell is calculated.',
      inputSchema: {
        type: 'object',
        properties: {},
        required: [],
      },
      executeExcel: async context => {
        const range = context.workbook.getSelectedRange();
        range.load('values, formulas, address, rowCount, columnCount');
        await context.sync();
        return JSON.stringify(
          {
            address: range.address,
            rowCount: range.rowCount,
            columnCount: range.columnCount,
            values: range.values,
            formulas: range.formulas,
          },
          null,
          2,
        );
      },
    },

    getWorksheetData: {
      name: 'getWorksheetData',
      category: 'read',
      description:
        'Get data from a worksheet. By default, reads all data from the used range of the active worksheet. Optionally specify a worksheet name and/or a specific range address. Returns the values, formulas, address, row count, and column count. Use formulas (not values) when you need to understand how cells are calculated.',
      inputSchema: {
        type: 'object',
        properties: {
          sheetName: {
            type: 'string',
            description: 'Optional worksheet name. Uses active sheet if omitted.',
          },
          address: {
            type: 'string',
            description: 'Optional range address (e.g., "A1:D10"). Uses used range if omitted.',
          },
        },
        required: [],
      },
      executeExcel: async (context, args: Record<string, any>) => {
        const { sheetName, address } = args || {};

        const sheet = sheetName
          ? await safeGetSheet(context, sheetName)
          : context.workbook.worksheets.getActiveWorksheet();

        const range = address ? sheet.getRange(address) : sheet.getUsedRange();
        range.load('values, formulas, address, rowCount, columnCount');
        await context.sync();

        return JSON.stringify(
          {
            worksheet: sheetName || '(active)',
            address: range.address,
            rowCount: range.rowCount,
            columnCount: range.columnCount,
            values: range.values,
            formulas: range.formulas,
          },
          null,
          2,
        );
      },
    },

    createTable: {
      name: 'createTable',
      category: 'write',
      description:
        'Convert a range to an Excel structured table (ListObject), with optional table name and style.',
      inputSchema: {
        type: 'object',
        properties: {
          address: {
            type: 'string',
            description:
              'Range address to convert into a table (e.g., "A1:D20"). Uses selection if omitted.',
          },
          hasHeaders: {
            type: 'boolean',
            description: 'Whether the first row contains headers. Default: true.',
          },
          tableName: {
            type: 'string',
            description: 'Optional table name (must be unique).',
          },
          style: {
            type: 'string',
            description: 'Optional table style name (e.g., "TableStyleMedium2").',
          },
        },
        required: [],
      },
      executeExcel: async (context, args: Record<string, any>) => {
        const { address, hasHeaders = true, tableName, style = 'TableStyleMedium2' } = args;

        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = address ? sheet.getRange(address) : context.workbook.getSelectedRange();
        const table = sheet.tables.add(range, hasHeaders);

        if (tableName) table.name = tableName;
        if (style) table.style = style;

        table.load('name');
        await context.sync();
        return `Successfully created table "${table.name}"${address ? ` from ${address}` : ' from selection'}`;
      },
    },

    setCellRange: {
      name: 'setCellRange',
      category: 'write',
      description:
        'PREFERRED tool for ALL write operations in Excel. Write values OR formulas to a range, apply formatting, and optionally fill down a formula to a larger range — all in one call. Automatically underlines modified cells (font underline) so the user can review changes; use `clearAgentHighlights` to remove the underline when done. Use `copyToRange` to fill a formula from the first row of `address` down to a larger range (e.g., address="C2:C2", copyToRange="C2:C50"). For multi-cell writes, always prefer passing a 2D array to `values` over calling this tool multiple times.',
      inputSchema: {
        type: 'object',
        properties: {
          address: {
            type: 'string',
            description: 'Target cell or range in A1 notation (e.g., "A1", "B2:D10"). Required.',
          },
          sheetName: {
            type: 'string',
            description: 'Optional worksheet name. Uses active sheet if omitted.',
          },
          values: {
            type: 'array',
            description:
              'A 2D array of values to write, e.g. [["Name","Score"],["Alice",95],["Bob",true],[null,3.14]]. Each cell value can be: string, number, boolean, null, or Date object. Use null to skip/clear a cell. Mutually exclusive with `formulas`.',
            items: {
              type: 'array',
              items: {
                anyOf: [
                  { type: 'string' },
                  { type: 'number' },
                  { type: 'boolean' },
                  { type: 'null' },
                ],
              },
            },
          },
          formulas: {
            type: 'array',
            description:
              'A 2D array of formulas to write, e.g. [["=SUM(A2:A10)"],["=AVERAGE(B2:B10)"]]. Each formula must start with "=". Mutually exclusive with `values`.',
            items: { type: 'array', items: { type: 'string' } },
          },
          formatting: {
            type: 'object',
            description: 'Optional formatting to apply to the range after writing.',
            properties: {
              bold: { type: 'boolean' },
              fillColor: { type: 'string', description: 'Hex color (e.g., "#FFFF00")' },
              fontColor: { type: 'string', description: 'Hex color (e.g., "#000000")' },
              numberFormat: {
                type: 'string',
                description: 'Number format string (e.g., "0.00%", "#,##0", "dd/mm/yyyy")',
              },
              horizontalAlignment: { type: 'string', enum: ['Left', 'Center', 'Right'] },
            },
          },
          copyToRange: {
            type: 'string',
            description:
              'Optional. Fill the formula/values from `address` down to a larger range (e.g., "C2:C50"). The source range (`address`) must be the first row of this range.',
          },
        },
        required: ['address'],
      },
      executeExcel: async (context, args: Record<string, any>) => {
        const { address, sheetName, values, formulas, formatting, copyToRange } = args;
        const formulaLocale = getExcelFormulaLanguage();

        const sheet = await safeGetSheet(context, sheetName);

        const range = sheet.getRange(address);

        // Write values or formulas
        if (formulas) {
          if (formulaLocale === 'fr') {
            range.formulasLocal = formulas;
          } else {
            range.formulas = formulas;
          }
        } else if (values !== undefined) {
          range.values = values;
        }

        // Auto-mark modified cells with text underline so the user can review changes
        if (values !== undefined || formulas) {
          range.format.font.underline = Excel.RangeUnderlineStyle.single;
        }

        // Apply formatting
        if (formatting) {
          if (formatting.bold !== undefined) range.format.font.bold = formatting.bold;
          if (formatting.fillColor) range.format.fill.color = formatting.fillColor;
          if (formatting.fontColor) range.format.font.color = formatting.fontColor;
          if (formatting.numberFormat) range.numberFormat = [[formatting.numberFormat]];
          if (formatting.horizontalAlignment) {
            const alignMap: Record<string, any> = {
              Left: Excel.HorizontalAlignment.left,
              Center: Excel.HorizontalAlignment.center,
              Right: Excel.HorizontalAlignment.right,
            };
            range.format.horizontalAlignment =
              alignMap[formatting.horizontalAlignment] ?? Excel.HorizontalAlignment.general;
          }
        }

        await context.sync();

        // Fill-down to copyToRange
        if (copyToRange) {
          const fullRange = sheet.getRange(copyToRange);
          range.autoFill(fullRange, Excel.AutoFillType.fillDefault);
          await context.sync();
          return `Successfully wrote to ${address} and filled down to ${copyToRange}`;
        }

        return `Successfully wrote to ${address}${sheetName ? ` on sheet "${sheetName}"` : ''}`;
      },
    },

    modifyStructure: {
      name: 'modifyStructure',
      category: 'write',
      description:
        'Insert, delete, hide, unhide, or freeze rows/columns. Use this instead of eval_officejs for structural changes. Examples: insert 3 rows before row 5, delete column B, hide rows 10-15, freeze the first row.',
      inputSchema: {
        type: 'object',
        properties: {
          operation: {
            type: 'string',
            enum: ['insert', 'delete', 'hide', 'unhide', 'freeze', 'unfreeze'],
            description: 'Operation to perform.',
          },
          dimension: {
            type: 'string',
            enum: ['rows', 'columns'],
            description: 'Whether to operate on rows or columns.',
          },
          reference: {
            type: 'string',
            description:
              'Row number(s) or column letter(s) to target, e.g. "5" (row 5), "3:7" (rows 3-7), "B" (column B), "B:D" (columns B-D). For freeze, use the first row/column number to freeze before (e.g., "2" freezes row 1).',
          },
          count: {
            type: 'number',
            description: 'Number of rows/columns to insert (for insert operation). Defaults to 1.',
          },
          sheetName: {
            type: 'string',
            description: 'Optional worksheet name. Uses active sheet if omitted.',
          },
        },
        required: ['operation', 'dimension'],
      },
      executeExcel: async (context, args: Record<string, any>) => {
        const { operation, dimension, reference, count = 1, sheetName } = args;

        const sheet = await safeGetSheet(context, sheetName);

        if (operation === 'freeze' || operation === 'unfreeze') {
          if (operation === 'unfreeze') {
            sheet.freezePanes.unfreeze();
          } else {
            const ref = parseInt(reference, 10) || 1;
            if (dimension === 'rows') {
              sheet.freezePanes.freezeRows(ref);
            } else {
              sheet.freezePanes.freezeColumns(ref);
            }
          }
          await context.sync();
          return `Successfully ${operation}d ${dimension}`;
        }

        if (dimension === 'rows') {
          const rangeRef = reference
            ? `${reference}:${reference.includes(':') ? reference.split(':')[1] : reference}`
            : '1:1';
          const rowRange = sheet.getRange(rangeRef);
          if (operation === 'insert') {
            rowRange.insert(Excel.InsertShiftDirection.down);
            if (count > 1) {
              // Insert additional rows
              for (let i = 1; i < count; i++) {
                const insertRef = reference ? `${reference}:${reference}` : '1:1';
                sheet.getRange(insertRef).insert(Excel.InsertShiftDirection.down);
              }
            }
          } else if (operation === 'delete') {
            rowRange.delete(Excel.DeleteShiftDirection.up);
          } else if (operation === 'hide') {
            rowRange.rowHidden = true;
          } else if (operation === 'unhide') {
            rowRange.rowHidden = false;
          }
        } else {
          const colRef = reference || 'A:A';
          const colRange = sheet.getRange(colRef);
          if (operation === 'insert') {
            colRange.insert(Excel.InsertShiftDirection.right);
          } else if (operation === 'delete') {
            colRange.delete(Excel.DeleteShiftDirection.left);
          } else if (operation === 'hide') {
            colRange.columnHidden = true;
          } else if (operation === 'unhide') {
            colRange.columnHidden = false;
          }
        }

        await context.sync();
        return `Successfully ${operation}d ${dimension}${reference ? ` ${reference}` : ''}${sheetName ? ` on sheet "${sheetName}"` : ''}`;
      },
    },

    manageObject: {
      name: 'manageObject',
      category: 'write',
      description:
        "Create, update, or delete charts and pivot tables. For create/update, specify an explicit sheetName and source range so the agent can target any sheet without depending on the user's current selection.",
      inputSchema: {
        type: 'object',
        properties: {
          operation: {
            type: 'string',
            enum: ['create', 'update', 'delete'],
            description: 'The operation to perform.',
          },
          objectType: {
            type: 'string',
            enum: ['chart', 'pivotTable'],
            description: 'The type of object to manage.',
          },
          sheetName: {
            type: 'string',
            description: 'Name of the worksheet to operate on. If omitted, uses the active sheet.',
          },
          source: {
            type: 'string',
            description:
              'For create/update: data range address used as the chart source (e.g. "A1:D50"). Required for create.',
          },
          chartType: {
            type: 'string',
            description: 'For chart create/update: chart type.',
            enum: [
              'ColumnClustered',
              'ColumnStacked',
              'Line',
              'LineMarkers',
              'Pie',
              'BarClustered',
              'Area',
              'Doughnut',
              'XYScatter',
              'Waterfall',
              'Treemap',
              'Funnel',
            ],
          },
          title: {
            type: 'string',
            description: 'For chart create/update: chart title.',
          },
          name: {
            type: 'string',
            description: 'For update/delete: name of the existing chart or pivot table to target.',
          },
          anchor: {
            type: 'string',
            description:
              'For chart create: cell address where the chart top-left corner will be placed (e.g. "F1"). Defaults to a position beside the source data.',
          },
          hasHeaders: {
            type: 'boolean',
            description:
              'For chart create: whether the first row or column of the source range contains headers/labels (not data). When true, labels are used as axis categories instead of being plotted as a data series. Default: true.',
          },
          seriesBy: {
            type: 'string',
            description:
              'For chart create: how to interpret data series. "columns" means each column is a series, "rows" means each row is a series. Default: "columns".',
            enum: ['columns', 'rows'],
          },
        },
        required: ['operation', 'objectType'],
      },
      executeExcel: async (context, args: Record<string, any>) => {
        const {
          operation,
          objectType,
          sheetName,
          source,
          chartType,
          title,
          name,
          anchor,
          seriesBy = 'columns',
          hasHeaders = true,
        } = args as Record<string, any>;

        const chartTypeMap: Record<string, any> = {
          ColumnClustered: Excel.ChartType.columnClustered,
          ColumnStacked: Excel.ChartType.columnStacked,
          Line: Excel.ChartType.line,
          LineMarkers: Excel.ChartType.lineMarkers,
          Pie: Excel.ChartType.pie,
          BarClustered: Excel.ChartType.barClustered,
          Area: Excel.ChartType.area,
          Doughnut: Excel.ChartType.doughnut,
          XYScatter: Excel.ChartType.xyscatter,
          Waterfall: Excel.ChartType.waterfall,
          Treemap: Excel.ChartType.treemap,
          Funnel: Excel.ChartType.funnel,
        };

        // Resolve target sheet
        const sheet = await safeGetSheet(context, sheetName);

        if (operation === 'create') {
          if (objectType === 'chart') {
            if (!source) return 'Error: source range is required to create a chart.';
            const dataRange = sheet.getRange(source);
            dataRange.load(['values', 'columnCount', 'rowCount']);
            await context.sync();

            if (
              !dataRange.values ||
              (dataRange.values.length <= 1 && dataRange.values[0]?.length <= 1)
            ) {
              return 'Error: source range is too small. Provide a range with headers and data (at least 2 rows or 2 columns).';
            }

            // XL-M1 Fix: Split the label column/row from the data range so Excel never
            // plots it as an extra data series (especially problematic with numeric labels like years).
            let plotRange: Excel.Range = dataRange;
            let categoryRange: Excel.Range | null = null;

            if (hasHeaders) {
              if (seriesBy === 'columns' && dataRange.columnCount > 1) {
                // Category labels = first column, EXCLUDING the top-left header cell so it
                // never shows up as a data category (e.g. "Mois" must not appear in the axis).
                categoryRange = dataRange
                  .getColumn(0)
                  .getOffsetRange(1, 0)
                  .getResizedRange(-1, 0);
                // Plot data = remaining columns (row 0 used as series name headers by Excel)
                plotRange = dataRange.getOffsetRange(0, 1).getResizedRange(0, -1);
              } else if (seriesBy === 'rows' && dataRange.rowCount > 1) {
                // Category labels = first row, EXCLUDING the top-left header cell
                categoryRange = dataRange
                  .getRow(0)
                  .getOffsetRange(0, 1)
                  .getResizedRange(0, -1);
                // Plot data = remaining rows (column 0 used as series name headers by Excel)
                plotRange = dataRange.getOffsetRange(1, 0).getResizedRange(-1, 0);
              }
            }

            const excelChartType = chartTypeMap[chartType] || Excel.ChartType.columnClustered;
            const seriesByEnum =
              seriesBy === 'rows' ? Excel.ChartSeriesBy.rows : Excel.ChartSeriesBy.columns;
            const chart = sheet.charts.add(excelChartType, plotRange, seriesByEnum);

            // Explicitly bind the category axis to the label range to ensure correct axis labels
            if (categoryRange) {
              try {
                chart.axes.categoryAxis.setCategoryNames(categoryRange);
              } catch (e) {
                logService.warn('[ExcelTools] Failed to explicitly set category names', e);
              }
            }

            if (title) chart.title.text = title;
            chart.width = 400;
            chart.height = 300;

            if (anchor) {
              const anchorRange = sheet.getRange(anchor);
              chart.setPosition(anchorRange, undefined);
            }

            await context.sync();
            return `Successfully created ${chartType || 'ColumnClustered'} chart${title ? ` titled "${title}"` : ''} from range ${source}${sheetName ? ` on sheet "${sheetName}"` : ''}. Series interpreted by ${seriesBy}${hasHeaders ? ' (label column/row excluded from plot range)' : ''}.`;
          }

          if (objectType === 'pivotTable') {
            if (!source) return 'Error: source range is required to create a pivot table.';
            if (!name)
              return 'Error: name is required to create a pivot table (used as the pivot table name).';
            const destRange = anchor ? sheet.getRange(anchor) : sheet.getRange('A1');
            sheet.pivotTables.add(name, source, destRange);
            await context.sync();
            return `Successfully created pivot table "${name}" from range ${source}.`;
          }
        }

        if (operation === 'update') {
          if (objectType === 'chart') {
            if (!name) return 'Error: name is required to update a chart.';
            const chart = sheet.charts.getItem(name);
            if (chartType) chart.chartType = chartTypeMap[chartType] || chart.chartType;
            if (title) chart.title.text = title;
            if (source) {
              const newDataRange = sheet.getRange(source);
              const updateSeriesBy =
                seriesBy === 'rows' ? Excel.ChartSeriesBy.rows : Excel.ChartSeriesBy.columns;
              chart.setData(newDataRange, updateSeriesBy);
            }
            await context.sync();
            return `Successfully updated chart "${name}".`;
          }
        }

        if (operation === 'delete') {
          if (!name) return 'Error: name is required to delete an object.';
          if (objectType === 'chart') {
            sheet.charts.getItem(name).delete();
          } else {
            sheet.pivotTables.getItem(name).delete();
          }
          await context.sync();
          return `Successfully deleted ${objectType} "${name}"${sheetName ? ` from sheet "${sheetName}"` : ''}.`;
        }

        return `Error: unsupported operation "${operation}" for objectType "${objectType}".`;
      },
    },

    formatRange: {
      name: 'formatRange',
      category: 'format',
      description:
        '⚠️ DEPRECATED: Use setCellRange with formatting parameter instead. This tool is redundant and will be removed in a future version. Apply formatting to the selected range or a specific range address. Can set fill color, font color, bold, italic, font size, and borders.',
      inputSchema: {
        type: 'object',
        properties: {
          address: {
            type: 'string',
            description: 'Optional cell address. If not provided, formats the current selection.',
          },
          fillColor: {
            type: 'string',
            description: 'Background fill color as hex (e.g., "#FF0000" for red)',
          },
          fontColor: {
            type: 'string',
            description: 'Font color as hex (e.g., "#FFFFFF" for white)',
          },
          bold: {
            type: 'boolean',
            description: 'Make text bold',
          },
          italic: {
            type: 'boolean',
            description: 'Make text italic',
          },
          fontSize: {
            type: 'number',
            description: 'Font size in points',
          },
          borders: {
            type: 'boolean',
            description: 'Add borders around all cells',
          },
          horizontalAlignment: {
            type: 'string',
            description: 'Horizontal alignment',
            enum: ['Left', 'Center', 'Right'],
          },
          verticalAlignment: {
            type: 'string',
            description: 'Vertical alignment',
            enum: ['Top', 'Center', 'Bottom', 'Justify', 'Distributed'],
          },
          wrapText: {
            type: 'boolean',
            description: 'Wrap text within cells',
          },
          fontName: {
            type: 'string',
            description: 'Font family name (e.g., "Calibri", "Arial")',
          },
          borderStyle: {
            type: 'string',
            description: 'Default border style to apply to all borders.',
            enum: [
              'continuous',
              'dash',
              'dashDot',
              'dashDotDot',
              'dot',
              'double',
              'none',
              'slantDashDot',
            ],
          },
          borderColor: {
            type: 'string',
            description: 'Default border color (hex).',
          },
          borderWeight: {
            type: 'string',
            description: 'Default border weight to apply to all borders.',
            enum: ['hairline', 'thin', 'medium', 'thick'],
          },
          borderTopStyle: { type: 'string' },
          borderBottomStyle: { type: 'string' },
          borderLeftStyle: { type: 'string' },
          borderRightStyle: { type: 'string' },
          borderInsideHorizontalStyle: { type: 'string' },
          borderInsideVerticalStyle: { type: 'string' },
          borderTopColor: { type: 'string' },
          borderBottomColor: { type: 'string' },
          borderLeftColor: { type: 'string' },
          borderRightColor: { type: 'string' },
          borderInsideHorizontalColor: { type: 'string' },
          borderInsideVerticalColor: { type: 'string' },
          borderTopWeight: { type: 'string' },
          borderBottomWeight: { type: 'string' },
          borderLeftWeight: { type: 'string' },
          borderRightWeight: { type: 'string' },
          borderInsideHorizontalWeight: { type: 'string' },
          borderInsideVerticalWeight: { type: 'string' },
        },
        required: [],
      },
      executeExcel: async (context, args: Record<string, any>) => {
        const {
          address,
          fillColor,
          fontColor,
          bold,
          italic,
          fontSize,
          borders,
          horizontalAlignment,
          verticalAlignment,
          wrapText,
          fontName,
          borderStyle,
          borderColor,
          borderWeight,
        } = args as Record<string, any>;

        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = address ? sheet.getRange(address) : context.workbook.getSelectedRange();

        if (fillColor) range.format.fill.color = fillColor;
        if (fontColor) range.format.font.color = fontColor;
        if (bold !== undefined) range.format.font.bold = bold;
        if (italic !== undefined) range.format.font.italic = italic;
        if (fontSize) range.format.font.size = fontSize;
        if (fontName) range.format.font.name = fontName;
        if (wrapText !== undefined) range.format.wrapText = wrapText;
        if (horizontalAlignment) {
          range.format.horizontalAlignment = horizontalAlignment as Excel.HorizontalAlignment;
        }
        if (verticalAlignment) {
          range.format.verticalAlignment = verticalAlignment as Excel.VerticalAlignment;
        }

        const borderStyleMap: Record<string, any> = {
          continuous: Excel.BorderLineStyle.continuous,
          dash: Excel.BorderLineStyle.dash,
          dashDot: Excel.BorderLineStyle.dashDot,
          dashDotDot: Excel.BorderLineStyle.dashDotDot,
          dot: Excel.BorderLineStyle.dot,
          double: Excel.BorderLineStyle.double,
          none: Excel.BorderLineStyle.none,
          slantDashDot: Excel.BorderLineStyle.slantDashDot,
        };

        const borderWeightMap: Record<string, any> = {
          hairline: Excel.BorderWeight.hairline,
          thin: Excel.BorderWeight.thin,
          medium: Excel.BorderWeight.medium,
          thick: Excel.BorderWeight.thick,
        };

        const setBorder = (
          borderIndex: Excel.BorderIndex,
          overrides: { style?: string; color?: string; weight?: string },
        ) => {
          try {
            const border = range.format.borders.getItem(borderIndex);
            const styleToApply = overrides.style ?? borderStyle;
            const colorToApply = overrides.color ?? borderColor;
            const weightToApply = overrides.weight ?? borderWeight;

            if (styleToApply)
              border.style = borderStyleMap[styleToApply] ?? Excel.BorderLineStyle.continuous;
            if (colorToApply) border.color = colorToApply;
            if (weightToApply)
              border.weight = borderWeightMap[weightToApply] ?? Excel.BorderWeight.thin;
          } catch {
            // Some border types may not apply to single cells
          }
        };

        setBorder(Excel.BorderIndex.edgeTop, {
          style: args.borderTopStyle,
          color: args.borderTopColor,
          weight: args.borderTopWeight,
        });
        setBorder(Excel.BorderIndex.edgeBottom, {
          style: args.borderBottomStyle,
          color: args.borderBottomColor,
          weight: args.borderBottomWeight,
        });
        setBorder(Excel.BorderIndex.edgeLeft, {
          style: args.borderLeftStyle,
          color: args.borderLeftColor,
          weight: args.borderLeftWeight,
        });
        setBorder(Excel.BorderIndex.edgeRight, {
          style: args.borderRightStyle,
          color: args.borderRightColor,
          weight: args.borderRightWeight,
        });
        setBorder(Excel.BorderIndex.insideHorizontal, {
          style: args.borderInsideHorizontalStyle,
          color: args.borderInsideHorizontalColor,
          weight: args.borderInsideHorizontalWeight,
        });
        setBorder(Excel.BorderIndex.insideVertical, {
          style: args.borderInsideVerticalStyle,
          color: args.borderInsideVerticalColor,
          weight: args.borderInsideVerticalWeight,
        });

        if (borders) {
          const borderItems = [
            Excel.BorderIndex.edgeTop,
            Excel.BorderIndex.edgeBottom,
            Excel.BorderIndex.edgeLeft,
            Excel.BorderIndex.edgeRight,
            Excel.BorderIndex.insideHorizontal,
            Excel.BorderIndex.insideVertical,
          ];
          for (const border of borderItems) {
            try {
              const b = range.format.borders.getItem(border);
              b.style = Excel.BorderLineStyle.continuous;
              b.color = '#000000';
            } catch {
              // Some border types may not apply to single cells
            }
          }
        }

        await context.sync();
        return `Successfully applied formatting${address ? ` to ${address}` : ' to selection'}`;
      },
    },

    sortRange: {
      name: 'sortRange',
      category: 'write',
      description:
        'Sort a data range by a specific column. Pass an explicit address (e.g. "A1:D50") to sort a known range, or omit it to sort the current selection.',
      inputSchema: {
        type: 'object',
        properties: {
          address: {
            type: 'string',
            description:
              'Optional range address to sort (e.g. "A1:D50", "Sheet2!B2:E100"). If omitted, the current user selection is used.',
          },
          columnIndex: {
            type: 'number',
            description: 'Zero-based column index to sort by (default: 0)',
          },
          ascending: {
            type: 'boolean',
            description: 'Sort ascending (true) or descending (false). Default: true.',
          },
          hasHeaders: {
            type: 'boolean',
            description: 'Whether the range has headers in the first row. Default: true.',
          },
        },
        required: [],
      },
      executeExcel: async (context, args: Record<string, any>) => {
        const {
          address,
          columnIndex = 0,
          ascending = true,
          hasHeaders = true,
        } = args as Record<string, any>;

        const range = address
          ? context.workbook.worksheets.getActiveWorksheet().getRange(address)
          : context.workbook.getSelectedRange();
        range.load('values, rowCount, columnCount');
        await context.sync();

        hasHeaders ? range.getResizedRange(-1, 0).getOffsetRange(1, 0) : range;

        // Manual sort as fallback-safe approach
        const values = range.values.slice();
        const headers = hasHeaders ? [values.shift()!] : [];
        values.sort((a, b) => {
          const va = a[columnIndex];
          const vb = b[columnIndex];
          if (va < vb) return ascending ? -1 : 1;
          if (va > vb) return ascending ? 1 : -1;
          return 0;
        });

        range.values = [...headers, ...values];
        await context.sync();
        return `Successfully sorted data by column ${columnIndex} (${ascending ? 'ascending' : 'descending'})`;
      },
    },

    getWorksheetInfo: {
      name: 'getWorksheetInfo',
      category: 'read',
      description:
        'Get information about the active worksheet including name, used range dimensions, and worksheet count.',
      inputSchema: {
        type: 'object',
        properties: {},
        required: [],
      },
      executeExcel: async context => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.load('name, id, position');
        const usedRange = sheet.getUsedRangeOrNullObject();
        usedRange.load('address, rowCount, columnCount');

        const sheets = context.workbook.worksheets;
        sheets.load('items/name');
        await context.sync();

        const sheetNames = sheets.items.map((s: any) => s.name);

        return JSON.stringify(
          {
            activeName: sheet.name,
            position: sheet.position,
            usedRange: usedRange.isNullObject
              ? null
              : {
                  address: usedRange.address,
                  rowCount: usedRange.rowCount,
                  columnCount: usedRange.columnCount,
                },
            totalSheets: sheetNames.length,
            sheetNames,
          },
          null,
          2,
        );
      },
    },

    clearRange: {
      name: 'clearRange',
      category: 'write',
      description: 'Clear the contents, formatting, or both from the selected range.',
      inputSchema: {
        type: 'object',
        properties: {
          address: {
            type: 'string',
            description: 'Optional range address. Uses selection if not provided.',
          },
          clearType: {
            type: 'string',
            description:
              'What to clear: "contents" (values only), "formats" (formatting only), or "all" (both)',
            enum: ['contents', 'formats', 'all'],
          },
        },
        required: [],
      },
      executeExcel: async (context, args: Record<string, any>) => {
        const { address, clearType = 'all' } = args;

        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = address ? sheet.getRange(address) : context.workbook.getSelectedRange();

        switch (clearType) {
          case 'contents':
            range.clear(Excel.ClearApplyTo.contents);
            break;
          case 'formats':
            range.clear(Excel.ClearApplyTo.formats);
            break;
          default:
            range.clear(Excel.ClearApplyTo.all);
        }

        await context.sync();
        return `Successfully cleared ${clearType}${address ? ` from ${address}` : ' from selection'}`;
      },
    },

    searchAndReplace: {
      name: 'searchAndReplace',
      category: 'write',
      description: 'Search for a value in the used range and optionally replace it.',
      inputSchema: {
        type: 'object',
        properties: {
          searchText: {
            type: 'string',
            description: 'Text to search for',
          },
          replaceText: {
            type: 'string',
            description: 'Optional replacement text. If omitted, just searches without replacing.',
          },
        },
        required: ['searchText'],
      },
      executeExcel: async (context, args: Record<string, any>) => {
        const { searchText, replaceText } = args as Record<string, any>;

        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const usedRange = sheet.getUsedRange();
        usedRange.load('values, rowCount, columnCount');
        await context.sync();

        let matchCount = 0;
        const newValues = usedRange.values.map((row: any[]) =>
          row.map((cell: any) => {
            const cellStr = String(cell);
            if (cellStr.includes(searchText)) {
              matchCount++;
              if (replaceText !== undefined) {
                return cellStr.replace(
                  new RegExp(searchText.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g'),
                  replaceText,
                );
              }
            }
            return cell;
          }),
        );

        if (replaceText !== undefined && matchCount > 0) {
          usedRange.values = newValues;
          await context.sync();
          return `Found and replaced ${matchCount} occurrence(s) of "${searchText}" with "${replaceText}"`;
        }

        return `Found ${matchCount} occurrence(s) of "${searchText}"`;
      },
    },

    addWorksheet: {
      name: 'addWorksheet',
      category: 'write',
      description: 'Add a new worksheet to the workbook.',
      inputSchema: {
        type: 'object',
        properties: {
          name: {
            type: 'string',
            description: 'Name for the new worksheet',
          },
        },
        required: [],
      },
      executeExcel: async (context: Excel.RequestContext, args: Record<string, any>) => {
        const { name } = args as Record<string, any>;

        const sheet = context.workbook.worksheets.add(name || undefined);
        sheet.activate();
        sheet.load('name');
        await context.sync();
        return `Successfully created and activated worksheet "${sheet.name}"`;
      },
    },

    protectWorksheet: {
      name: 'protectWorksheet',
      category: 'write',
      description: 'Protect or unprotect a worksheet with optional password and permissions.',
      inputSchema: {
        type: 'object',
        properties: {
          sheetName: {
            type: 'string',
            description: 'Optional worksheet name. Uses active worksheet if omitted.',
          },
          action: {
            type: 'string',
            description: 'Protection action.',
            enum: ['protect', 'unprotect'],
          },
          password: {
            type: 'string',
            description: 'Optional password.',
          },
          allowAutoFilter: {
            type: 'boolean',
            description: 'Allow users to use auto filters while protected.',
          },
          allowFormatCells: {
            type: 'boolean',
            description: 'Allow users to format cells while protected.',
          },
          allowInsertRows: {
            type: 'boolean',
            description: 'Allow users to insert rows while protected.',
          },
        },
        required: ['action'],
      },
      executeExcel: async (context: Excel.RequestContext, args: Record<string, any>) => {
        const {
          sheetName,
          action,
          password,
          allowAutoFilter = false,
          allowFormatCells = false,
          allowInsertRows = false,
        } = args as Record<string, any>;

        const sheet = await safeGetSheet(context, sheetName);

        if (action === 'protect') {
          sheet.protection.protect(
            {
              allowAutoFilter,
              allowFormatCells,
              allowInsertRows,
              selectionMode: Excel.ProtectionSelectionMode.normal,
            },
            password,
          );
        } else {
          sheet.protection.unprotect(password);
        }

        await context.sync();
        return `Successfully ${action === 'protect' ? 'protected' : 'unprotected'} worksheet "${sheetName ?? 'active'}"`;
      },
    },

    getNamedRanges: {
      name: 'getNamedRanges',
      category: 'read',
      description: 'List workbook named ranges and their formulas/references.',
      inputSchema: {
        type: 'object',
        properties: {},
        required: [],
      },
      executeExcel: async context => {
        const names = context.workbook.names;
        names.load('items/name,items/formula,items/value');
        await context.sync();

        return JSON.stringify(
          {
            totalNamedRanges: names.items.length,
            items: names.items.map((item: any) => ({
              name: item.name,
              formula: item.formula,
              value: item.value,
            })),
          },
          null,
          2,
        );
      },
    },

    setNamedRange: {
      name: 'setNamedRange',
      category: 'write',
      description: 'Create or update a workbook named range.',
      inputSchema: {
        type: 'object',
        properties: {
          name: {
            type: 'string',
            description: 'Named range identifier.',
          },
          rangeAddress: {
            type: 'string',
            description: 'Range formula/address reference (e.g., "=Sheet1!$A$1:$B$10").',
          },
        },
        required: ['name', 'rangeAddress'],
      },
      executeExcel: async (context: Excel.RequestContext, args: Record<string, any>) => {
        const { name, rangeAddress } = args as Record<string, any>;

        context.workbook.names.add(name, rangeAddress);
        await context.sync();
        return `Successfully set named range "${name}" = ${rangeAddress}`;
      },
    },

    applyConditionalFormatting: {
      name: 'applyConditionalFormatting',
      category: 'format',
      description:
        'Create conditional formatting rules on a range, including cell value, text contains, custom formulas, color scales, data bars, and icon sets. Can also clear existing rules first.',
      inputSchema: {
        type: 'object',
        properties: {
          address: {
            type: 'string',
            description: 'Target range address in A1 notation (e.g., "A2:A100").',
          },
          ruleType: {
            type: 'string',
            description: 'Conditional formatting rule type to apply.',
            enum: ['cellValue', 'containsText', 'custom', 'colorScale', 'dataBar', 'iconSet'],
          },
          clearExisting: {
            type: 'boolean',
            description:
              'If true, clear existing conditional formats on the target range before applying the new rule.',
          },
          operator: {
            type: 'string',
            description: 'Operator for cellValue rule.',
            enum: [
              'between',
              'notBetween',
              'equalTo',
              'notEqualTo',
              'greaterThan',
              'greaterThanOrEqual',
              'lessThan',
              'lessThanOrEqual',
            ],
          },
          formula1: {
            type: 'string',
            description:
              'First formula/value for cellValue or custom rule (examples: "100", "=A2>AVERAGE($A$2:$A$100)"). IMPORTANT: Always use English function names and comma separators here, regardless of the Excel formula language setting — the conditional formatting API does not support localized formulas.',
          },
          formula2: {
            type: 'string',
            description:
              'Second formula/value for between/notBetween operators. Always use English syntax.',
          },
          text: {
            type: 'string',
            description: 'Text to match for containsText rule.',
          },
          textOperator: {
            type: 'string',
            description: 'Operator for containsText rule.',
            enum: ['contains', 'beginsWith', 'endsWith', 'notContains'],
          },
          fillColor: {
            type: 'string',
            description: 'Fill color (hex) for matching cells, e.g. "#FFCDD2".',
          },
          fontColor: {
            type: 'string',
            description: 'Font color (hex) for matching cells.',
          },
          bold: {
            type: 'boolean',
            description: 'Set font bold style on matching cells.',
          },
          colorScaleMinColor: {
            type: 'string',
            description: 'Minimum color for colorScale rules.',
          },
          colorScaleMidColor: {
            type: 'string',
            description: 'Midpoint color for colorScale rules.',
          },
          colorScaleMaxColor: {
            type: 'string',
            description: 'Maximum color for colorScale rules.',
          },
          dataBarColor: {
            type: 'string',
            description: 'Bar color for dataBar rules.',
          },
          iconSetStyle: {
            type: 'string',
            description: 'Icon set style for iconSet rules.',
            enum: [
              'threeTrafficLights1',
              'threeArrows',
              'threeSymbols',
              'fourArrows',
              'fourTrafficLights',
              'fiveArrows',
            ],
          },
          stopIfTrue: {
            type: 'boolean',
            description: 'Whether evaluation should stop if this rule is true.',
          },
        },
        required: ['address', 'ruleType'],
      },
      executeExcel: async (context, args: Record<string, any>) => {
        const {
          address,
          ruleType,
          clearExisting = false,
          operator = 'greaterThan',
          formula1 = '0',
          formula2,
          text,
          textOperator = 'contains',
          fillColor,
          fontColor,
          bold,
          colorScaleMinColor = '#F8696B',
          colorScaleMidColor = '#FFEB84',
          colorScaleMaxColor = '#63BE7B',
          dataBarColor = '#5B9BD5',
          iconSetStyle = 'threeTrafficLights1',
          stopIfTrue,
        } = args;

        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const targetRange = sheet.getRange(address);
        const conditionalFormats: any = targetRange.conditionalFormats;

        if (clearExisting) {
          conditionalFormats.clearAll();
        }

        const ruleTypeMap: Record<string, any> = {
          cellValue: Excel.ConditionalFormatType.cellValue,
          containsText: Excel.ConditionalFormatType.containsText,
          custom: Excel.ConditionalFormatType.custom,
          colorScale: Excel.ConditionalFormatType.colorScale,
          dataBar: Excel.ConditionalFormatType.dataBar,
          iconSet: Excel.ConditionalFormatType.iconSet,
        };

        const selectedType = ruleTypeMap[ruleType];
        if (!selectedType) {
          throw new Error(`Unsupported conditional formatting ruleType: ${ruleType}`);
        }

        const cf: any = conditionalFormats.add(selectedType);

        if (ruleType === 'cellValue') {
          const operatorMap: Record<string, any> = {
            between: Excel.ConditionalCellValueOperator.between,
            notBetween: Excel.ConditionalCellValueOperator.notBetween,
            equalTo: Excel.ConditionalCellValueOperator.equalTo,
            notEqualTo: Excel.ConditionalCellValueOperator.notEqualTo,
            greaterThan: Excel.ConditionalCellValueOperator.greaterThan,
            greaterThanOrEqual: Excel.ConditionalCellValueOperator.greaterThanOrEqual,
            lessThan: Excel.ConditionalCellValueOperator.lessThan,
            lessThanOrEqual: Excel.ConditionalCellValueOperator.lessThanOrEqual,
          };
          cf.cellValue.rule = {
            formula1,
            ...(formula2 ? { formula2 } : {}),
            operator: operatorMap[operator] ?? Excel.ConditionalCellValueOperator.greaterThan,
          };
        } else if (ruleType === 'containsText') {
          const textOperatorMap: Record<string, any> = {
            contains: Excel.ConditionalTextOperator.contains,
            beginsWith: Excel.ConditionalTextOperator.beginsWith,
            endsWith: Excel.ConditionalTextOperator.endsWith,
            notContains: Excel.ConditionalTextOperator.notContains,
          };
          cf.textComparison.rule = {
            operator: textOperatorMap[textOperator] ?? Excel.ConditionalTextOperator.contains,
            text: text ?? '',
          };
        } else if (ruleType === 'custom') {
          cf.custom.rule.formula = formula1;
        } else if (ruleType === 'colorScale') {
          cf.colorScale.criteria = {
            minimum: {
              color: colorScaleMinColor,
              type: Excel.ConditionalFormatColorCriterionType.lowestValue,
            },
            midpoint: {
              color: colorScaleMidColor,
              type: Excel.ConditionalFormatColorCriterionType.percentile,
              formula: '50',
            },
            maximum: {
              color: colorScaleMaxColor,
              type: Excel.ConditionalFormatColorCriterionType.highestValue,
            },
          };
        } else if (ruleType === 'dataBar') {
          cf.dataBar.barColor = dataBarColor;
          cf.dataBar.lowerBoundRule = { type: Excel.ConditionalFormatRuleType.lowestValue };
          cf.dataBar.upperBoundRule = { type: Excel.ConditionalFormatRuleType.highestValue };
        } else if (ruleType === 'iconSet') {
          const iconSetMap: Record<string, any> = {
            threeTrafficLights1: Excel.IconSet.threeTrafficLights1,
            threeArrows: Excel.IconSet.threeArrows,
            threeSymbols: Excel.IconSet.threeSymbols,
            fourArrows: Excel.IconSet.fourArrows,
            fourTrafficLights: Excel.IconSet.fourTrafficLights,
            fiveArrows: Excel.IconSet.fiveArrows,
          };
          cf.iconSet.style = iconSetMap[iconSetStyle] ?? Excel.IconSet.threeTrafficLights1;
        }

        const applyTextAndFillFormat = (format: any) => {
          if (!format) return;
          if (fillColor) format.fill.color = fillColor;
          if (fontColor) format.font.color = fontColor;
          if (bold !== undefined) format.font.bold = bold;
        };

        if (ruleType === 'cellValue') {
          applyTextAndFillFormat(cf.cellValue?.format);
        } else if (ruleType === 'containsText') {
          applyTextAndFillFormat(cf.textComparison?.format);
        } else if (ruleType === 'custom') {
          applyTextAndFillFormat(cf.custom?.format);
        }

        if (stopIfTrue !== undefined) cf.stopIfTrue = stopIfTrue;

        await context.sync();
        return `Successfully applied ${ruleType} conditional formatting on ${address}`;
      },
    },

    getConditionalFormattingRules: {
      name: 'getConditionalFormattingRules',
      category: 'format',
      description:
        'Read conditional formatting rules from a target range (or from the worksheet used range when no address is provided).',
      inputSchema: {
        type: 'object',
        properties: {
          address: {
            type: 'string',
            description: 'Optional range address (e.g., "A1:D20"). Uses used range if omitted.',
          },
        },
        required: [],
      },
      executeExcel: async (context: Excel.RequestContext, args: Record<string, any>) => {
        const { address } = args as Record<string, any>;

        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const targetRange = address ? sheet.getRange(address) : sheet.getUsedRangeOrNullObject();
        targetRange.load('address,isNullObject');
        await context.sync();

        if (targetRange.isNullObject) {
          return 'No used range found on the active worksheet, so no conditional formatting rules were read.';
        }

        const conditionalFormats = targetRange.conditionalFormats;
        conditionalFormats.load('items/type,items/priority,items/stopIfTrue');
        await context.sync();

        return JSON.stringify(
          {
            address: targetRange.address,
            totalRules: conditionalFormats.items.length,
            rules: conditionalFormats.items.map((rule: any) => ({
              type: rule.type,
              priority: rule.priority,
              stopIfTrue: rule.stopIfTrue,
            })),
          },
          null,
          2,
        );
      },
    },

    findData: {
      name: 'findData',
      category: 'read',
      description:
        'Find text or values across the spreadsheet. Returns matching cells with their addresses and values. Options for regex, match case, entire cell match, and formula search. Supports pagination via maxResults and offset.',
      inputSchema: {
        type: 'object',
        properties: {
          searchTerm: { type: 'string', description: 'The text or pattern to search for' },
          matchCase: { type: 'boolean', description: 'Case sensitive. Default: false' },
          matchEntireCell: {
            type: 'boolean',
            description: 'Match entire cell content. Default: false',
          },
          useRegex: { type: 'boolean', description: 'Use regex pattern. Default: false' },
          searchInFormulas: {
            type: 'boolean',
            description:
              'Search in cell formulas instead of display values. Useful for debugging formulas. Default: false.',
          },
          maxResults: { type: 'number', description: 'Max results to return. Default: 50.' },
          offset: {
            type: 'number',
            description: 'Skip first N matches for pagination. Default: 0.',
          },
        },
        required: ['searchTerm'],
      },
      executeExcel: async (context, args: Record<string, any>) => {
        const {
          searchTerm,
          matchCase = false,
          matchEntireCell = false,
          useRegex = false,
          searchInFormulas = false,
        } = args as Record<string, any>;

        let pattern: RegExp | null = null;
        if (useRegex) {
          try {
            pattern = new RegExp(searchTerm, matchCase ? '' : 'i');
          } catch {
            return JSON.stringify({
              error: `Invalid regex pattern: "${searchTerm}". Please provide a valid regular expression.`,
            });
          }
        }

        const sheets = context.workbook.worksheets;
        sheets.load('items');
        await context.sync();

        const allMatches: any[] = [];

        // Load values or formulas depending on search mode — formula search ported from Office Agents
        const dataProperty = searchInFormulas ? 'formulas' : 'values';

        for (const sheet of sheets.items) {
          sheet.load('name');
          const usedRange = sheet.getUsedRangeOrNullObject();
          usedRange.load(`${dataProperty},address,rowCount,columnCount`);
          await context.sync();

          if (usedRange.isNullObject) continue;

          const startMatch = usedRange.address.split('!')[1]?.match(/([A-Z]+)(\d+)/);
          const startCol = startMatch
            ? startMatch[1]
                .split('')
                .reduce((acc: number, c: string) => acc * 26 + c.charCodeAt(0) - 64, 0) - 1
            : 0;
          const startRow = startMatch ? parseInt(startMatch[2], 10) - 1 : 0;
          const colLetter = (idx: number) => {
            let letter = '';
            let temp = idx;
            while (temp >= 0) {
              letter = String.fromCharCode((temp % 26) + 65) + letter;
              temp = Math.floor(temp / 26) - 1;
            }
            return letter;
          };

          const dataGrid = searchInFormulas ? usedRange.formulas : usedRange.values;
          for (let r = 0; r < usedRange.rowCount; r++) {
            for (let c = 0; c < usedRange.columnCount; c++) {
              const val = dataGrid[r][c];
              const target = String(val ?? '');
              let isMatch = false;
              if (pattern) {
                isMatch = pattern.test(target);
              } else {
                const compVal = matchCase ? target : target.toLowerCase();
                const compTerm = matchCase ? searchTerm : searchTerm.toLowerCase();
                isMatch = matchEntireCell ? compVal === compTerm : compVal.includes(compTerm);
              }
              if (isMatch) {
                allMatches.push({
                  sheet: sheet.name,
                  address: `${colLetter(startCol + c)}${startRow + r + 1}`,
                  value: val,
                });
              }
            }
          }
        }

        const offset = args.offset || 0;
        const maxResults = args.maxResults || 50;
        const page = allMatches.slice(offset, offset + maxResults);
        const hasMore = offset + maxResults < allMatches.length;

        return JSON.stringify(
          {
            matches: page,
            totalFound: allMatches.length,
            returned: page.length,
            offset,
            hasMore,
            nextOffset: hasMore ? offset + maxResults : null,
          },
          null,
          2,
        );
      },
    },

    getAllObjects: {
      name: 'getAllObjects',
      category: 'read',
      description:
        'List all charts and pivot tables. By default scans the active sheet only. Pass allSheets: true to scan all sheets in the workbook (may be slow on large workbooks).',
      inputSchema: {
        type: 'object',
        properties: {
          allSheets: {
            type: 'boolean',
            description:
              'When true, list objects from ALL sheets. When false (default), list only the active sheet.',
          },
        },
        required: [],
      },
      executeExcel: async (context, args: Record<string, any>) => {
        const allSheets = args.allSheets === true; // default false

        if (!allSheets) {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          sheet.load('name');
          const charts = sheet.charts;
          const pivotTables = sheet.pivotTables;
          charts.load('items/name, items/id');
          pivotTables.load('items/name, items/id');
          await context.sync();

          return JSON.stringify(
            {
              charts: charts.items.map((c: any) => ({
                name: c.name,
                id: c.id,
                sheetName: sheet.name,
              })),
              pivotTables: pivotTables.items.map((p: any) => ({
                name: p.name,
                id: p.id,
                sheetName: sheet.name,
              })),
            },
            null,
            2,
          );
        }

        // Workbook-wide scan
        const worksheets = context.workbook.worksheets;
        worksheets.load('items/name');
        await context.sync();

        for (const sheet of worksheets.items) {
          sheet.charts.load('items/name, items/id');
          sheet.pivotTables.load('items/name, items/id');
        }
        await context.sync();

        const allCharts: any[] = [];
        const allPivots: any[] = [];
        for (const sheet of worksheets.items) {
          for (const c of sheet.charts.items) {
            allCharts.push({ name: c.name, id: c.id, sheetName: sheet.name });
          }
          for (const p of sheet.pivotTables.items) {
            allPivots.push({ name: p.name, id: p.id, sheetName: sheet.name });
          }
        }

        return JSON.stringify({ charts: allCharts, pivotTables: allPivots }, null, 2);
      },
    },

    eval_officejs: {
      name: 'eval_officejs',
      category: 'write',
      description: `Execute custom Office.js code within an Excel.run context.

**USE THIS TOOL ONLY WHEN:**
- No dedicated tool exists for your operation
- Operations like: sorting, autofilter, freeze panes, hyperlinks, row/column operations, data validation, number formats, cell comments, named ranges, sheet operations, etc.

**REQUIRED CODE STRUCTURE:**
\`\`\`javascript
try {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getUsedRange();

  range.load('values,address');
  await context.sync();

  // Your operations here

  await context.sync();
  return { success: true, result: 'Operation completed' };
} catch (error) {
  return { success: false, error: error.message };
}
\`\`\`

**CRITICAL RULES:**
1. ALWAYS call \`.load()\` before reading properties
2. ALWAYS call \`await context.sync()\` after load and after modifications
3. ALWAYS wrap in try/catch
4. ONLY use Excel namespace (not Word, PowerPoint)
5. Values MUST be 2D arrays: \`range.values = [[value]]\`
6. **Formula language**: Use \`range.formulasLocal\` (not \`range.formulas\`) when writing formulas if the user's Excel locale is French. Check the \`excelFormulaLanguage\` setting from the agent context. When in doubt, use English syntax and \`range.formulas\`.`,
      inputSchema: {
        type: 'object',
        properties: {
          code: {
            type: 'string',
            description:
              'JavaScript code following the template. Must include load(), sync(), and try/catch.',
          },
          explanation: {
            type: 'string',
            description: 'Brief explanation of what this code does (required for audit trail).',
          },
        },
        required: ['code', 'explanation'],
      },
      executeExcel: createEvalExecutor<Excel.RequestContext>({
        host: 'Excel',
        toolName: 'eval_officejs',
        suggestion:
          'Refer to the Office.js skill document for correct patterns. Remember: Excel values must be 2D arrays.',
        mutationDetector: looksLikeMutation,
        buildSandboxContext: (context) => ({
          context,
          Excel: typeof Excel !== 'undefined' ? Excel : undefined,
          Office: typeof Office !== 'undefined' ? Office : undefined,
          ...getVfsSandboxContext(),
        }),
      }),
    },
    screenshotRange: {
      name: 'screenshotRange',
      category: 'read',
      description:
        'Capture a visual screenshot of an Excel range as PNG image. Use this to verify visual formatting, chart rendering, or analyze existing content visually. Requires ExcelApi 1.7+.',
      inputSchema: {
        type: 'object',
        properties: {
          sheetName: {
            type: 'string',
            description: 'Worksheet name. Uses active sheet if omitted.',
          },
          range: {
            type: 'string',
            description: 'A1 notation range, e.g. "A1:F20". Uses used range if omitted.',
          },
        },
        required: [],
      },
      executeExcel: async (context: Excel.RequestContext, args: Record<string, any>) => {
        const sheet = await safeGetSheet(context, args.sheetName);
        const targetRange = args.range ? sheet.getRange(args.range) : sheet.getUsedRange();

        // Load dimensions for header composition — ported from Office Agents
        targetRange.load('address, rowCount, columnCount');
        await context.sync();

        const numCols = targetRange.columnCount;
        const numRows = targetRange.rowCount;

        // Load column widths and row heights
        const cols: Excel.Range[] = [];
        for (let i = 0; i < numCols; i++) {
          const col = targetRange.getColumn(i);
          col.format.load('columnWidth');
          cols.push(col);
        }
        const rows: Excel.Range[] = [];
        for (let i = 0; i < numRows; i++) {
          const row = targetRange.getRow(i);
          row.format.load('rowHeight');
          rows.push(row);
        }

        const imageResult = (targetRange as any).getImage();
        await context.sync();

        const base64 = imageResult.value as string;
        const colWidths = cols.map(c => c.format.columnWidth);
        const rowHeights = rows.map(r => r.format.rowHeight);
        const { startRow, startCol } = parseRangeStart(targetRange.address);

        // Composite with row/column headers for vision model accuracy
        const composited = await compositeWithHeaders(
          base64,
          startRow,
          startCol,
          colWidths,
          rowHeights,
        );

        return buildScreenshotResult(
          composited,
          `Screenshot of range ${args.range || 'used range'} on sheet ${args.sheetName || 'active'} (with row/column headers)`,
        );
      },
    },

    getRangeAsCsv: {
      name: 'getRangeAsCsv',
      category: 'read',
      description:
        'Get a range of cells as CSV text. More token-efficient than JSON for large datasets. Use for data analysis when formatting details are not needed.',
      inputSchema: {
        type: 'object',
        properties: {
          sheetName: {
            type: 'string',
            description: 'Worksheet name. Uses active sheet if omitted.',
          },
          range: {
            type: 'string',
            description: 'A1 notation range, e.g. "A1:F100". Uses used range if omitted.',
          },
          maxRows: { type: 'number', description: 'Maximum rows to return. Default: 500.' },
        },
        required: [],
      },
      executeExcel: async (context: Excel.RequestContext, args: Record<string, any>) => {
        const sheet = await safeGetSheet(context, args.sheetName);
        const targetRange = args.range ? sheet.getRange(args.range) : sheet.getUsedRange();
        targetRange.load('values,rowCount,columnCount');
        await context.sync();

        const maxRows = args.maxRows || 500;
        const values = targetRange.values;
        const rows = values.slice(0, maxRows);

        const csv = rows
          .map((row: any[]) =>
            row
              .map((cell: any) => {
                const str = String(cell ?? '');
                if (str.includes(',') || str.includes('"') || str.includes('\n')) {
                  return '"' + str.replace(/"/g, '""') + '"';
                }
                return str;
              })
              .join(','),
          )
          .join('\n');

        const hasMore = values.length > maxRows;
        return `Rows: ${rows.length}/${values.length}${hasMore ? ' (truncated, use offset parameter)' : ''}\n\n${csv}`;
      },
    },

    modifyWorkbookStructure: {
      name: 'modifyWorkbookStructure',
      category: 'write',
      description:
        'Create, delete, rename, or duplicate a worksheet. Use instead of addWorksheet for operations that need delete/rename/duplicate.',
      inputSchema: {
        type: 'object',
        properties: {
          operation: {
            type: 'string',
            enum: ['create', 'delete', 'rename', 'duplicate'],
            description: 'The operation to perform.',
          },
          sheetName: {
            type: 'string',
            description: 'Name of the target sheet (for delete/rename/duplicate).',
          },
          newName: {
            type: 'string',
            description: 'New name for the sheet (for create/rename/duplicate).',
          },
          tabColor: {
            type: 'string',
            description: 'Optional hex color for the sheet tab, e.g. "#FF0000".',
          },
        },
        required: ['operation'],
      },
      executeExcel: async (context: Excel.RequestContext, args: Record<string, any>) => {
        const { operation, sheetName, newName, tabColor } = args;
        const sheets = context.workbook.worksheets;

        switch (operation) {
          case 'create': {
            const newSheet = sheets.add(newName || undefined);
            if (tabColor) newSheet.tabColor = tabColor;
            newSheet.activate();
            newSheet.load('name');
            await context.sync();
            return `Worksheet "${newSheet.name}" created.`;
          }
          case 'delete': {
            if (!sheetName) throw new Error('sheetName is required for delete operation.');
            const sheet = sheets.getItemOrNullObject(sheetName);
            await context.sync();
            if (sheet.isNullObject) throw new Error(`Worksheet "${sheetName}" not found.`);
            sheet.delete();
            await context.sync();
            return `Worksheet "${sheetName}" deleted.`;
          }
          case 'rename': {
            if (!sheetName) throw new Error('sheetName is required for rename operation.');
            if (!newName) throw new Error('newName is required for rename operation.');
            const sheet = await safeGetSheet(context, sheetName);
            sheet.name = newName;
            await context.sync();
            return `Worksheet "${sheetName}" renamed to "${newName}".`;
          }
          case 'duplicate': {
            if (!sheetName) throw new Error('sheetName is required for duplicate operation.');
            const sheet = await safeGetSheet(context, sheetName);
            const copy = (sheet as any).copy();
            if (newName) copy.name = newName;
            await context.sync();
            return `Worksheet "${sheetName}" duplicated${newName ? ` as "${newName}"` : ''}.`;
          }
          default:
            throw new Error(`Unknown operation: ${operation}`);
        }
      },
    },

    detectDataHeaders: {
      name: 'detectDataHeaders',
      category: 'read',
      description:
        'Analyze a range to detect whether it has column headers (first row = text labels) and/or row headers (first column = text labels). Returns detection results and the exact hasHeaders + seriesBy values to pass to manageObject when creating a chart. ALWAYS call this before creating a chart from user data.',
      inputSchema: {
        type: 'object',
        properties: {
          address: {
            type: 'string',
            description:
              'Cell range to analyze (e.g. "A1:D20"). Uses current selection if omitted.',
          },
          sheetName: {
            type: 'string',
            description: 'Worksheet name. Uses active sheet if omitted.',
          },
        },
        required: [],
      },
      executeExcel: async (context, args) => {
        const sheet = await safeGetSheet(context, args.sheetName);
        const range = args.address
          ? sheet.getRange(args.address)
          : context.workbook.getSelectedRange();
        range.load(['values', 'rowCount', 'columnCount', 'address']);
        await context.sync();

        const values = range.values as any[][];
        if (!values || values.length < 2 || values[0].length < 2) {
          return JSON.stringify({
            error: 'Range too small to detect headers (needs at least 2 rows and 2 columns).',
          });
        }

        const rows = values.length;
        const cols = values[0].length;

        const isText = (v: any) => typeof v === 'string' && v.trim().length > 0;
        const isNumeric = (v: any) =>
          typeof v === 'number' ||
          (typeof v === 'string' &&
            v.trim() !== '' &&
            !isNaN(Number(v.replace(/\s/g, '').replace(',', '.'))));

        // Column headers: first row is mostly text (not numbers)
        const firstRow = values[0];
        const firstRowText = firstRow.filter(isText).length;
        const firstRowNum = firstRow.filter(isNumeric).length;
        const hasColumnHeaders = firstRowText > 0 && firstRowText >= firstRowNum;

        // Row headers: first column (excluding row 0) is mostly text
        const firstColData = values.slice(1).map(r => r[0]);
        const firstColText = firstColData.filter(isText).length;
        const firstColNum = firstColData.filter(isNumeric).length;
        const hasRowHeaders = firstColText > 0 && firstColText >= firstColNum;

        // Verify data body is predominantly numeric
        const dataRowStart = hasColumnHeaders ? 1 : 0;
        const dataColStart = hasRowHeaders ? 1 : 0;
        let total = 0,
          numeric = 0;
        for (let r = dataRowStart; r < rows; r++) {
          for (let c = dataColStart; c < cols; c++) {
            const v = values[r][c];
            if (v !== '' && v !== null && v !== undefined) {
              total++;
              if (isNumeric(v)) numeric++;
            }
          }
        }
        const dataIsNumeric = total === 0 || numeric / total > 0.6;

        // Determine suggested parameters for chart creation
        const suggestedHasHeaders = hasColumnHeaders || hasRowHeaders;
        // If headers are in the first column only → data series are in rows
        const suggestedSeriesBy = hasRowHeaders && !hasColumnHeaders ? 'rows' : 'columns';

        const columnLabels = hasColumnHeaders ? firstRow.filter(isText) : [];
        const rowLabels = hasRowHeaders ? firstColData.filter(isText) : [];

        return JSON.stringify(
          {
            rangeAddress: (range as any).address,
            rowCount: rows,
            columnCount: cols,
            hasColumnHeaders,
            hasRowHeaders,
            dataIsNumeric,
            columnLabels,
            rowLabels,
            suggestedHasHeaders,
            suggestedSeriesBy,
            recommendation: `Use hasHeaders: ${suggestedHasHeaders}, seriesBy: "${suggestedSeriesBy}" when creating a chart from this range.`,
          },
          null,
          2,
        );
      },
    },

    // ============================================================
    // importCsvToSheet — ported from Office Agents csv-to-sheet custom command
    // packages/excel/src/lib/vfs/custom-commands.ts
    // ============================================================
    importCsvToSheet: {
      name: 'importCsvToSheet',
      category: 'write',
      description:
        'Import a CSV file from the VFS into an Excel worksheet. Reads the CSV, auto-detects data types (numbers, booleans, text), and writes to the specified sheet/cell. Use this after the user uploads a CSV file.',
      inputSchema: {
        type: 'object',
        properties: {
          filePath: {
            type: 'string',
            description: 'Path to CSV file in VFS (e.g., "/home/user/uploads/data.csv")',
          },
          sheetName: {
            type: 'string',
            description: 'Target worksheet name. Uses active sheet if omitted.',
          },
          startCell: {
            type: 'string',
            description: 'Starting cell in A1 notation (e.g., "A1"). Defaults to "A1".',
          },
          delimiter: {
            type: 'string',
            description: 'CSV delimiter character. Defaults to ",".',
          },
          overwrite: {
            type: 'boolean',
            description:
              'If true, overwrites existing cell data. Default is false (fails if cells contain data).',
          },
        },
        required: ['filePath'],
      },
      executeExcel: async (context: Excel.RequestContext, args: Record<string, any>) => {
        const {
          filePath,
          sheetName,
          startCell = 'A1',
          delimiter = ',',
          overwrite = false,
        } = args;

        const { readFile } = await import('@/utils/vfs');
        const csvContent = await readFile(filePath);
        if (!csvContent || !csvContent.trim()) {
          throw new Error(`File "${filePath}" is empty or not found.`);
        }

        // Parse CSV handling quoted fields — ported from Office Agents
        const lines = csvContent.trim().split('\n');
        const data: (string | number | boolean)[][] = [];
        for (const line of lines) {
          const row: (string | number | boolean)[] = [];
          let current = '';
          let inQuotes = false;
          for (let i = 0; i < line.length; i++) {
            const char = line[i];
            if (char === '"') {
              inQuotes = !inQuotes;
            } else if (char === delimiter && !inQuotes) {
              row.push(coerceValue(current.trim()));
              current = '';
            } else {
              current += char;
            }
          }
          row.push(coerceValue(current.trim()));
          data.push(row);
        }

        if (data.length === 0) throw new Error('CSV file contains no data rows.');

        // Pad rows to equal length
        const maxCols = Math.max(...data.map(r => r.length));
        for (const row of data) {
          while (row.length < maxCols) row.push('');
        }

        const sheet = await safeGetSheet(context, sheetName);
        const targetRange = sheet
          .getRange(startCell)
          .getResizedRange(data.length - 1, maxCols - 1);

        if (!overwrite) {
          targetRange.load('values');
          await context.sync();
          const hasData = targetRange.values.some((row: any[]) =>
            row.some((cell: any) => cell !== '' && cell !== null),
          );
          if (hasData) {
            throw new Error(
              'Target range contains existing data. Use overwrite=true to replace, or choose a different startCell.',
            );
          }
        }

        targetRange.values = data as any[][];
        await context.sync();

        return JSON.stringify({
          success: true,
          rowsImported: data.length,
          columnsImported: maxCols,
          targetRange: `${startCell}`,
          hasMutated: true,
        });
      },
    },

    // ============================================================
    // clearAgentHighlights — allows user to remove agent modification highlights
    // ============================================================
    clearAgentHighlights: {
      name: 'clearAgentHighlights',
      category: 'write',
      description:
        'Clear the text underline markings automatically applied to cells modified by the agent. Use this when the user has reviewed the changes and wants to remove the visual underline indicators.',
      inputSchema: {
        type: 'object',
        properties: {
          range: {
            type: 'string',
            description:
              'A1 notation range to clear markings from (e.g., "A1:D10"). Clears used range if omitted.',
          },
          sheetName: {
            type: 'string',
            description: 'Worksheet name. Uses active sheet if omitted.',
          },
        },
        required: [],
      },
      executeExcel: async (context: Excel.RequestContext, args: Record<string, any>) => {
        const sheet = await safeGetSheet(context, args.sheetName);
        const targetRange = args.range ? sheet.getRange(args.range) : sheet.getUsedRange();
        targetRange.load('address');
        await context.sync();

        // Clear the font underline marking automatically applied by setCellRange
        targetRange.format.font.underline = Excel.RangeUnderlineStyle.none;
        await context.sync();

        return JSON.stringify({
          success: true,
          message: `Cleared text underline markings on ${targetRange.address}. Agent modification indicators removed.`,
        });
      },
    },

    // ============================================================
    // imageToSheet — ported from Office Agents image-to-sheet custom command
    // packages/excel/src/lib/vfs/custom-commands.ts
    // Converts an image to "pixel art" in Excel cells using fill colors.
    // ============================================================
    imageToSheet: {
      name: 'imageToSheet',
      category: 'write',
      description:
        'Convert an uploaded image to pixel art in Excel by setting cell background colors. Reads an image from the VFS, downsamples it, and paints each pixel as a cell color. Max 200x200 pixels. Cells are resized to equal squares.',
      inputSchema: {
        type: 'object',
        properties: {
          filePath: {
            type: 'string',
            description: 'Path to image file in VFS (e.g., "/home/user/uploads/logo.png")',
          },
          width: {
            type: 'number',
            description: 'Target width in pixels/columns (max 200). E.g., 64.',
          },
          height: {
            type: 'number',
            description: 'Target height in pixels/rows (max 200). E.g., 64.',
          },
          sheetName: {
            type: 'string',
            description: 'Target worksheet name. Uses active sheet if omitted.',
          },
          startCell: {
            type: 'string',
            description: 'Top-left cell (e.g., "A1"). Defaults to "A1".',
          },
          cellSize: {
            type: 'number',
            description: 'Cell width/height in points (1-50). Default: 6.',
          },
        },
        required: ['filePath', 'width', 'height'],
      },
      executeExcel: async (context: Excel.RequestContext, args: Record<string, any>) => {
        const {
          filePath,
          width: targetW,
          height: targetH,
          sheetName,
          startCell = 'A1',
          cellSize = 6,
        } = args;

        if (targetW > 200 || targetH > 200 || targetW < 1 || targetH < 1) {
          throw new Error('Dimensions must be between 1 and 200.');
        }
        if (cellSize < 1 || cellSize > 50) {
          throw new Error('Cell size must be between 1 and 50 points.');
        }

        // Read image from VFS
        const vfs = getVfs();
        const fullPath = filePath.startsWith('/') ? filePath : `/home/user/uploads/${filePath}`;
        const data = await vfs.readFileBuffer(fullPath);

        // Decode and downsample image using Canvas
        const blob = new Blob([data as BlobPart]);
        const url = URL.createObjectURL(blob);
        let pixels: Uint8ClampedArray;
        try {
          const img = new Image();
          img.src = url;
          await new Promise<void>((resolve, reject) => {
            img.onload = () => resolve();
            img.onerror = () => reject(new Error('Failed to decode image'));
          });

          const canvas = document.createElement('canvas');
          canvas.width = targetW;
          canvas.height = targetH;
          const ctx = canvas.getContext('2d');
          if (!ctx) throw new Error('Failed to create canvas 2D context');
          ctx.drawImage(img, 0, 0, targetW, targetH);
          pixels = ctx.getImageData(0, 0, targetW, targetH).data;
        } finally {
          URL.revokeObjectURL(url);
        }

        // Build color-to-ranges map with RLE (run-length encoding) for efficiency
        const rgbToHex = (r: number, g: number, b: number) =>
          `#${r.toString(16).padStart(2, '0')}${g.toString(16).padStart(2, '0')}${b.toString(16).padStart(2, '0')}`;

        const sheet = await safeGetSheet(context, sheetName);
        const { startRow, startCol } = parseRangeStart(startCell);

        const colorRanges = new Map<string, string[]>();
        for (let y = 0; y < targetH; y++) {
          const rowNum = startRow + y + 1;
          let x = 0;
          while (x < targetW) {
            const i = (y * targetW + x) * 4;
            const hex = rgbToHex(pixels[i], pixels[i + 1], pixels[i + 2]);
            const runStart = x;
            x++;
            while (x < targetW) {
              const j = (y * targetW + x) * 4;
              if (
                pixels[j] !== pixels[i] ||
                pixels[j + 1] !== pixels[i + 1] ||
                pixels[j + 2] !== pixels[i + 2]
              )
                break;
              x++;
            }
            const addr =
              runStart === x - 1
                ? `${columnIndexToLetter(startCol + runStart)}${rowNum}`
                : `${columnIndexToLetter(startCol + runStart)}${rowNum}:${columnIndexToLetter(startCol + x - 1)}${rowNum}`;
            let ranges = colorRanges.get(hex);
            if (!ranges) {
              ranges = [];
              colorRanges.set(hex, ranges);
            }
            ranges.push(addr);
          }
        }

        // Set cell sizes and clear values
        const endCol = columnIndexToLetter(startCol + targetW - 1);
        const endRow = startRow + targetH;
        const fullRange = sheet.getRange(`${startCell}:${endCol}${endRow}`);
        fullRange.format.columnWidth = cellSize;
        fullRange.format.rowHeight = cellSize;
        const emptyValues: string[][] = Array.from({ length: targetH }, () =>
          Array.from({ length: targetW }, () => ''),
        );
        fullRange.values = emptyValues;
        await context.sync();

        // Apply colors in batches — ported from Office Agents
        const RANGES_PER_BATCH = 1000;
        let queued = 0;
        for (const [color, ranges] of colorRanges.entries()) {
          for (let i = 0; i < ranges.length; i += RANGES_PER_BATCH) {
            const batch = ranges.slice(i, i + RANGES_PER_BATCH);
            const areas = (sheet as any).getRanges(batch.join(','));
            areas.format.fill.color = color;
            queued += batch.length;
            if (queued >= RANGES_PER_BATCH) {
              await context.sync();
              queued = 0;
            }
          }
        }
        await context.sync();

        return JSON.stringify({
          success: true,
          pixelsWidth: targetW,
          pixelsHeight: targetH,
          totalCells: targetW * targetH,
          uniqueColors: colorRanges.size,
          cellSizePoints: cellSize,
          hasMutated: true,
        });
      },
    },
  },
  buildExecuteWrapper<ExcelToolTemplate>('executeExcel', runExcel),
);

/** Backend-calling tool: extracts data points from a chart image via pixel analysis. */
const extractChartDataTool: ToolDefinition = {
  name: 'extract_chart_data',
  category: 'read',
  description:
    'Extract numerical data points from a chart/graph image using pixel color analysis. ' +
    'You MUST first analyze the image yourself (via vision) to determine: (1) the axis ranges, ' +
    '(2) the color(s) of the data series, (3) the chart type, AND (4) the bounding box of the plot area ' +
    '(the rectangle delimited by the X and Y axes, excluding titles, legends and labels). ' +
    'Estimate plotAreaBox as fractions of the image (0.0–1.0): xMinPx = where the Y-axis line sits (left edge), ' +
    'xMaxPx = rightmost gridline/tick (right edge), yMinPx = topmost gridline (top edge), ' +
    'yMaxPx = where the X-axis line sits (bottom edge). ' +
    'Returns a JSON array of {x, y} points that you can write into Excel with setCellRange and chart with manageObject. ' +
    'The imageId is provided in the <uploaded_images> context block when the user uploads a chart image. ' +
    '**MULTI-CURVE CHARTS**: If the chart contains multiple data series (e.g., 3 different colored lines), ' +
    'call this tool ONCE PER SERIES with the specific targetColor for each series. ' +
    'First identify all series colors (e.g., red="#FF0000", blue="#0000FF", green="#00FF00"), ' +
    'then call extract_chart_data for each color separately. ' +
    'Write each series to adjacent Excel columns (e.g., columns A-B for series 1, C-D for series 2, etc.).',
  inputSchema: {
    type: 'object',
    properties: {
      imageId: {
        type: 'string',
        description:
          'The imageId from the <uploaded_images> context block (UUID returned by the upload endpoint).',
      },
      xAxisRange: {
        type: 'array',
        description:
          'Real-world min and max values of the X axis as [min, max]. Example: [2000, 2024]. Determine this by reading the axis labels in the image.',
        items: { type: 'number' },
      } as any,
      yAxisRange: {
        type: 'array',
        description:
          'Real-world min and max values of the Y axis as [min, max]. Example: [0, 50]. Determine this by reading the axis labels in the image.',
        items: { type: 'number' },
      } as any,
      targetColor: {
        type: 'string',
        description:
          'Hex color of the data series line/bars/points in the chart. Example: "#FF0000" for red, "#0000FF" for blue. Determine this by observing the chart image.',
      },
      plotAreaBox: {
        type: 'object',
        description:
          "Bounding box of the chart's plot area (the area delimited by the axes, excluding labels and legends). " +
          'Provide values as fractions of the image dimensions between 0.0 and 1.0. ' +
          'xMinPx: fraction from left where the Y-axis line is (e.g. 0.12). ' +
          'xMaxPx: fraction from left where the rightmost tick/gridline is (e.g. 0.95). ' +
          'yMinPx: fraction from top where the topmost gridline is (e.g. 0.08). ' +
          'yMaxPx: fraction from top where the X-axis line is (e.g. 0.88).',
        properties: {
          xMinPx: {
            type: 'number',
            description: 'Left edge of plot area as fraction [0,1] of image width.',
          },
          xMaxPx: {
            type: 'number',
            description: 'Right edge of plot area as fraction [0,1] of image width.',
          },
          yMinPx: {
            type: 'number',
            description: 'Top edge of plot area as fraction [0,1] of image height.',
          },
          yMaxPx: {
            type: 'number',
            description: 'Bottom edge of plot area as fraction [0,1] of image height.',
          },
        },
        required: ['xMinPx', 'xMaxPx', 'yMinPx', 'yMaxPx'],
      },
      chartType: {
        type: 'string',
        description:
          'Type of chart: "line", "scatter", "bar" (horizontal bars), or "area". Defaults to "line".',
        enum: ['line', 'scatter', 'bar', 'area'],
      },
      colorTolerance: {
        type: 'number',
        description:
          'Color matching tolerance (0-441, Euclidean RGB distance). Higher = more permissive. Default: 120. Increase if few points are returned.',
      },
      numPoints: {
        type: 'number',
        description: 'Desired number of output data points (5-200). Default: 40.',
      },
    },
    required: ['imageId', 'xAxisRange', 'yAxisRange', 'targetColor', 'plotAreaBox'],
  },
  execute: async args => {
    try {
      const result = await extractChartData({
        imageId: args.imageId,
        xAxisRange: args.xAxisRange,
        yAxisRange: args.yAxisRange,
        targetColor: args.targetColor,
        plotAreaBox: args.plotAreaBox,
        chartType: args.chartType,
        colorTolerance: args.colorTolerance,
        numPoints: args.numPoints,
      });

      if (result.warning) {
        return JSON.stringify(
          { success: false, warning: result.warning, pixelsMatched: 0 },
          null,
          2,
        );
      }

      return JSON.stringify(
        {
          success: true,
          pointCount: result.points.length,
          pixelsMatched: result.pixelsMatched,
          points: result.points,
        },
        null,
        2,
      );
    } catch (error: unknown) {
      return JSON.stringify(
        {
          success: false,
          error: getErrorMessage(error),
          suggestion:
            'Check that the imageId is valid and the image was recently uploaded. Adjust colorTolerance or targetColor if needed.',
        },
        null,
        2,
      );
    }
  },
};

export function getExcelToolDefinitions(): ToolDefinition[] {
  return [...Object.values(excelToolDefinitions), extractChartDataTool];
}

export { excelToolDefinitions };

/**
 * Direct helper: clear all agent modification underline highlights across the active workbook.
 * Called directly from the UI "Valider les modifications IA" button — bypasses the agent loop.
 * Returns a user-readable result string.
 */
export async function clearAllAgentHighlightsInWorkbook(): Promise<string> {
  try {
    return await runExcel(async (context: Excel.RequestContext) => {
      const sheets = context.workbook.worksheets;
      sheets.load('items/name');
      await context.sync();
      for (const sheet of sheets.items) {
        try {
          const used = sheet.getUsedRange();
          used.format.font.underline = Excel.RangeUnderlineStyle.none;
        } catch {
          // Sheet may be empty — ignore
        }
      }
      await context.sync();
      return `✓ Surlignements IA supprimés sur ${sheets.items.length} feuille${sheets.items.length === 1 ? '' : 's'}.`;
    });
  } catch (err: unknown) {
    logService.error('[ExcelTools] clearAllAgentHighlightsInWorkbook failed', err);
    const msg = err instanceof Error ? err.message : String(err);
    return `Erreur : ${msg}`;
  }
}
