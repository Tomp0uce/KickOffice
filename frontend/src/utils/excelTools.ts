import type { ToolProperty, ToolDefinition } from '@/types'
import { executeOfficeAction } from './officeAction'
import { createOfficeTools } from './common'
import { localStorageKey } from './enum'
import { sandboxedEval } from './sandbox'
import { validateOfficeCode } from './officeCodeValidator'

const runExcel = <T>(action: (context: Excel.RequestContext) => Promise<T>): Promise<T> =>
  executeOfficeAction(() => Excel.run(action))

/** Safely resolves a worksheet by name, falling back to active sheet. Throws a clear error if the sheet doesn't exist. */
async function safeGetSheet(context: Excel.RequestContext, sheetName?: string): Promise<Excel.Worksheet> {
  if (!sheetName) return context.workbook.worksheets.getActiveWorksheet()
  const sheet = context.workbook.worksheets.getItemOrNullObject(sheetName)
  await context.sync()
  if (sheet.isNullObject) {
    throw new Error(`Worksheet "${sheetName}" not found. Use getWorksheetInfo to list available worksheets.`)
  }
  return sheet
}

type ExcelToolTemplate = Omit<ToolDefinition, 'execute'> & {
  executeExcel: (context: Excel.RequestContext, args: Record<string, any>) => Promise<string>
}

export type ExcelToolName =
  | 'getSelectedCells'
  | 'setCellRange'
  | 'getWorksheetData'
  | 'createTable'
  | 'modifyStructure'
  | 'formatRange'
  | 'sortRange'
  | 'getWorksheetInfo'
  | 'getDataFromSheet'
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


function getExcelFormulaLanguage(): 'en' | 'fr' {
  const configured = localStorage.getItem(localStorageKey.excelFormulaLanguage)
  return configured === 'fr' ? 'fr' : 'en'
}


const excelToolDefinitions = createOfficeTools<ExcelToolName, ExcelToolTemplate, ToolDefinition>({
  getSelectedCells: {
    name: 'getSelectedCells',
    category: 'read',
    description:
      'Get the values, address, and dimensions of the currently selected cells in Excel. Returns a JSON object with address, rowCount, columnCount, and the 2D values array.',
    inputSchema: {
      type: 'object',
      properties: {},
      required: [],
    },
    executeExcel: async (context) => {
      
        const range = context.workbook.getSelectedRange()
        range.load('values, address, rowCount, columnCount')
        await context.sync()
        return JSON.stringify(
          {
            address: range.address,
            rowCount: range.rowCount,
            columnCount: range.columnCount,
            values: range.values,
          },
          null,
          2,
        )
      },
  },


  getWorksheetData: {
    name: 'getWorksheetData',
    category: 'read',
    description:
      'Get all data from the used range of the active worksheet. Returns the values, address, row count, and column count.',
    inputSchema: {
      type: 'object',
      properties: {},
      required: [],
    },
    executeExcel: async (context) => {
      
        const sheet = context.workbook.worksheets.getActiveWorksheet()
        const usedRange = sheet.getUsedRange()
        usedRange.load('values, address, rowCount, columnCount')
        await context.sync()
        return JSON.stringify(
          {
            address: usedRange.address,
            rowCount: usedRange.rowCount,
            columnCount: usedRange.columnCount,
            values: usedRange.values,
          },
          null,
          2,
        )
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
          description: 'Range address to convert into a table (e.g., "A1:D20"). Uses selection if omitted.',
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
      const { address, hasHeaders = true, tableName, style = 'TableStyleMedium2' } = args
      
        const sheet = context.workbook.worksheets.getActiveWorksheet()
        const range = address ? sheet.getRange(address) : context.workbook.getSelectedRange()
        const table = sheet.tables.add(range, hasHeaders)

        if (tableName) table.name = tableName
        if (style) table.style = style

        table.load('name')
        await context.sync()
        return `Successfully created table "${table.name}"${address ? ` from ${address}` : ' from selection'}`
      },
  },

  setCellRange: {
    name: 'setCellRange',
    category: 'write',
    description:
      'PREFERRED tool for ALL write operations in Excel. Write values OR formulas to a range, apply formatting, and optionally fill down a formula to a larger range — all in one call. Use `copyToRange` to fill a formula from the first row of `address` down to a larger range (e.g., address="C2:C2", copyToRange="C2:C50"). For multi-cell writes, always prefer passing a 2D array to `values` over calling this tool multiple times.',
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
          description: 'A 2D array of values to write, e.g. [["Name","Score"],["Alice",95]]. Use null to skip a cell. Mutually exclusive with `formulas`.',
          items: { type: 'array', items: { type: 'string' } },
        },
        formulas: {
          type: 'array',
          description: 'A 2D array of formulas to write, e.g. [["=SUM(A2:A10)"],["=AVERAGE(B2:B10)"]]. Each formula must start with "=". Mutually exclusive with `values`.',
          items: { type: 'array', items: { type: 'string' } },
        },
        formatting: {
          type: 'object',
          description: 'Optional formatting to apply to the range after writing.',
          properties: {
            bold: { type: 'boolean' },
            fillColor: { type: 'string', description: 'Hex color (e.g., "#FFFF00")' },
            fontColor: { type: 'string', description: 'Hex color (e.g., "#000000")' },
            numberFormat: { type: 'string', description: 'Number format string (e.g., "0.00%", "#,##0", "dd/mm/yyyy")' },
            horizontalAlignment: { type: 'string', enum: ['Left', 'Center', 'Right'] },
          },
        },
        copyToRange: {
          type: 'string',
          description: 'Optional. Fill the formula/values from `address` down to a larger range (e.g., "C2:C50"). The source range (`address`) must be the first row of this range.',
        },
      },
      required: ['address'],
    },
    executeExcel: async (context, args: Record<string, any>) => {
      const { address, sheetName, values, formulas, formatting, copyToRange } = args
      const formulaLocale = getExcelFormulaLanguage()

      const sheet = await safeGetSheet(context, sheetName)

      const range = sheet.getRange(address)

      // Write values or formulas
      if (formulas) {
        if (formulaLocale === 'fr') {
          range.formulasLocal = formulas
        } else {
          range.formulas = formulas
        }
      } else if (values !== undefined) {
        range.values = values
      }

      // Apply formatting
      if (formatting) {
        if (formatting.bold !== undefined) range.format.font.bold = formatting.bold
        if (formatting.fillColor) range.format.fill.color = formatting.fillColor
        if (formatting.fontColor) range.format.font.color = formatting.fontColor
        if (formatting.numberFormat) range.numberFormat = [[formatting.numberFormat]]
        if (formatting.horizontalAlignment) {
          const alignMap: Record<string, any> = {
            Left: Excel.HorizontalAlignment.left,
            Center: Excel.HorizontalAlignment.center,
            Right: Excel.HorizontalAlignment.right,
          }
          range.format.horizontalAlignment = alignMap[formatting.horizontalAlignment] ?? Excel.HorizontalAlignment.general
        }
      }

      await context.sync()

      // Fill-down to copyToRange
      if (copyToRange) {
        const fullRange = sheet.getRange(copyToRange)
        range.autoFill(fullRange, Excel.AutoFillType.fillDefault)
        await context.sync()
        return `Successfully wrote to ${address} and filled down to ${copyToRange}`
      }

      return `Successfully wrote to ${address}${sheetName ? ` on sheet "${sheetName}"` : ''}`
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
          description: 'Row number(s) or column letter(s) to target, e.g. "5" (row 5), "3:7" (rows 3-7), "B" (column B), "B:D" (columns B-D). For freeze, use the first row/column number to freeze before (e.g., "2" freezes row 1).',
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
      const { operation, dimension, reference, count = 1, sheetName } = args

      const sheet = await safeGetSheet(context, sheetName)

      if (operation === 'freeze' || operation === 'unfreeze') {
        if (operation === 'unfreeze') {
          sheet.freezePanes.unfreeze()
        } else {
          const ref = parseInt(reference, 10) || 1
          if (dimension === 'rows') {
            sheet.freezePanes.freezeRows(ref)
          } else {
            sheet.freezePanes.freezeColumns(ref)
          }
        }
        await context.sync()
        return `Successfully ${operation}d ${dimension}`
      }

      if (dimension === 'rows') {
        const rangeRef = reference ? `${reference}:${reference.includes(':') ? reference.split(':')[1] : reference}` : '1:1'
        const rowRange = sheet.getRange(rangeRef)
        if (operation === 'insert') {
          rowRange.insert(Excel.InsertShiftDirection.down)
          if (count > 1) {
            // Insert additional rows
            for (let i = 1; i < count; i++) {
              const insertRef = reference ? `${reference}:${reference}` : '1:1'
              sheet.getRange(insertRef).insert(Excel.InsertShiftDirection.down)
            }
          }
        } else if (operation === 'delete') {
          rowRange.delete(Excel.DeleteShiftDirection.up)
        } else if (operation === 'hide') {
          rowRange.rowHidden = true
        } else if (operation === 'unhide') {
          rowRange.rowHidden = false
        }
      } else {
        const colRef = reference || 'A:A'
        const colRange = sheet.getRange(colRef)
        if (operation === 'insert') {
          colRange.insert(Excel.InsertShiftDirection.right)
        } else if (operation === 'delete') {
          colRange.delete(Excel.DeleteShiftDirection.left)
        } else if (operation === 'hide') {
          colRange.columnHidden = true
        } else if (operation === 'unhide') {
          colRange.columnHidden = false
        }
      }

      await context.sync()
      return `Successfully ${operation}d ${dimension}${reference ? ` ${reference}` : ''}${sheetName ? ` on sheet "${sheetName}"` : ''}`
    },
  },


  manageObject: {
    name: 'manageObject',
    category: 'write',
    description:
      'Create, update, or delete charts and pivot tables. For create/update, specify an explicit sheetName and source range so the agent can target any sheet without depending on the user\'s current selection.',
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
          description: 'For create/update: data range address used as the chart source (e.g. "A1:D50"). Required for create.',
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
          description: 'For chart create: cell address where the chart top-left corner will be placed (e.g. "F1"). Defaults to a position beside the source data.',
        },
      },
      required: ['operation', 'objectType'],
    },
    executeExcel: async (context, args: Record<string, any>) => {
      const { operation, objectType, sheetName, source, chartType, title, name, anchor } = args as Record<string, any>

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
      }

      // Resolve target sheet
      const sheet = await safeGetSheet(context, sheetName)

      if (operation === 'create') {
        if (objectType === 'chart') {
          if (!source) return 'Error: source range is required to create a chart.'
          const dataRange = sheet.getRange(source)
          dataRange.load('values')
          await context.sync()

          if (!dataRange.values || (dataRange.values.length <= 1 && dataRange.values[0]?.length <= 1)) {
            return 'Error: source range is too small. Provide a range with headers and data (at least 2 rows or 2 columns).'
          }

          const excelChartType = chartTypeMap[chartType] || Excel.ChartType.columnClustered
          const chart = sheet.charts.add(excelChartType, dataRange, Excel.ChartSeriesBy.auto)

          if (title) chart.title.text = title
          chart.width = 400
          chart.height = 300

          if (anchor) {
            const anchorRange = sheet.getRange(anchor)
            chart.setPosition(anchorRange, undefined)
          }

          await context.sync()
          return `Successfully created ${chartType || 'ColumnClustered'} chart${title ? ` titled "${title}"` : ''} from range ${source}${sheetName ? ` on sheet "${sheetName}"` : ''}.`
        }

        if (objectType === 'pivotTable') {
          if (!source) return 'Error: source range is required to create a pivot table.'
          if (!name) return 'Error: name is required to create a pivot table (used as the pivot table name).'
          const destRange = anchor ? sheet.getRange(anchor) : sheet.getRange('A1')
          sheet.pivotTables.add(name, source, destRange)
          await context.sync()
          return `Successfully created pivot table "${name}" from range ${source}.`
        }
      }

      if (operation === 'update') {
        if (objectType === 'chart') {
          if (!name) return 'Error: name is required to update a chart.'
          const chart = sheet.charts.getItem(name)
          if (chartType) chart.chartType = chartTypeMap[chartType] || chart.chartType
          if (title) chart.title.text = title
          if (source) {
            const newDataRange = sheet.getRange(source)
            chart.setData(newDataRange, Excel.ChartSeriesBy.auto)
          }
          await context.sync()
          return `Successfully updated chart "${name}".`
        }
      }

      if (operation === 'delete') {
        if (!name) return 'Error: name is required to delete an object.'
        if (objectType === 'chart') {
          sheet.charts.getItem(name).delete()
        } else {
          sheet.pivotTables.getItem(name).delete()
        }
        await context.sync()
        return `Successfully deleted ${objectType} "${name}"${sheetName ? ` from sheet "${sheetName}"` : ''}.`
      }

      return `Error: unsupported operation "${operation}" for objectType "${objectType}".`
    },
  },

  formatRange: {
    name: 'formatRange',
    category: 'format',
    description:
      'Apply formatting to the selected range or a specific range address. Can set fill color, font color, bold, italic, font size, and borders.',
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
          enum: ['continuous', 'dash', 'dashDot', 'dashDotDot', 'dot', 'double', 'none', 'slantDashDot'],
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
      } = args as Record<string, any>
      
        const sheet = context.workbook.worksheets.getActiveWorksheet()
        const range = address ? sheet.getRange(address) : context.workbook.getSelectedRange()

        if (fillColor) range.format.fill.color = fillColor
        if (fontColor) range.format.font.color = fontColor
        if (bold !== undefined) range.format.font.bold = bold
        if (italic !== undefined) range.format.font.italic = italic
        if (fontSize) range.format.font.size = fontSize
        if (fontName) range.format.font.name = fontName
        if (wrapText !== undefined) range.format.wrapText = wrapText
        if (horizontalAlignment) {
          range.format.horizontalAlignment = horizontalAlignment as Excel.HorizontalAlignment
        }
        if (verticalAlignment) {
          range.format.verticalAlignment = verticalAlignment as Excel.VerticalAlignment
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
        }

        const borderWeightMap: Record<string, any> = {
          hairline: Excel.BorderWeight.hairline,
          thin: Excel.BorderWeight.thin,
          medium: Excel.BorderWeight.medium,
          thick: Excel.BorderWeight.thick,
        }

        const setBorder = (borderIndex: Excel.BorderIndex, overrides: { style?: string; color?: string; weight?: string }) => {
          try {
            const border = range.format.borders.getItem(borderIndex)
            const styleToApply = overrides.style ?? borderStyle
            const colorToApply = overrides.color ?? borderColor
            const weightToApply = overrides.weight ?? borderWeight

            if (styleToApply) border.style = borderStyleMap[styleToApply] ?? Excel.BorderLineStyle.continuous
            if (colorToApply) border.color = colorToApply
            if (weightToApply) border.weight = borderWeightMap[weightToApply] ?? Excel.BorderWeight.thin
          } catch {
            // Some border types may not apply to single cells
          }
        }

        setBorder(Excel.BorderIndex.edgeTop, {
          style: args.borderTopStyle,
          color: args.borderTopColor,
          weight: args.borderTopWeight,
        })
        setBorder(Excel.BorderIndex.edgeBottom, {
          style: args.borderBottomStyle,
          color: args.borderBottomColor,
          weight: args.borderBottomWeight,
        })
        setBorder(Excel.BorderIndex.edgeLeft, {
          style: args.borderLeftStyle,
          color: args.borderLeftColor,
          weight: args.borderLeftWeight,
        })
        setBorder(Excel.BorderIndex.edgeRight, {
          style: args.borderRightStyle,
          color: args.borderRightColor,
          weight: args.borderRightWeight,
        })
        setBorder(Excel.BorderIndex.insideHorizontal, {
          style: args.borderInsideHorizontalStyle,
          color: args.borderInsideHorizontalColor,
          weight: args.borderInsideHorizontalWeight,
        })
        setBorder(Excel.BorderIndex.insideVertical, {
          style: args.borderInsideVerticalStyle,
          color: args.borderInsideVerticalColor,
          weight: args.borderInsideVerticalWeight,
        })

        if (borders) {
          const borderItems = [
            Excel.BorderIndex.edgeTop,
            Excel.BorderIndex.edgeBottom,
            Excel.BorderIndex.edgeLeft,
            Excel.BorderIndex.edgeRight,
            Excel.BorderIndex.insideHorizontal,
            Excel.BorderIndex.insideVertical,
          ]
          for (const border of borderItems) {
            try {
              const b = range.format.borders.getItem(border)
              b.style = Excel.BorderLineStyle.continuous
              b.color = '#000000'
            } catch {
              // Some border types may not apply to single cells
            }
          }
        }

        await context.sync()
        return `Successfully applied formatting${address ? ` to ${address}` : ' to selection'}`
      },
  },

  sortRange: {
    name: 'sortRange',
    category: 'write',
    description: 'Sort a data range by a specific column. Pass an explicit address (e.g. "A1:D50") to sort a known range, or omit it to sort the current selection.',
    inputSchema: {
      type: 'object',
      properties: {
        address: {
          type: 'string',
          description: 'Optional range address to sort (e.g. "A1:D50", "Sheet2!B2:E100"). If omitted, the current user selection is used.',
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
      const { address, columnIndex = 0, ascending = true, hasHeaders = true } = args as Record<string, any>

        const range = address
          ? context.workbook.worksheets.getActiveWorksheet().getRange(address)
          : context.workbook.getSelectedRange()
        range.load('values, rowCount, columnCount')
        await context.sync()

        const dataRange = hasHeaders ? range.getResizedRange(-1, 0).getOffsetRange(1, 0) : range

        // Manual sort as fallback-safe approach
        const values = range.values.slice()
        const headers = hasHeaders ? [values.shift()!] : []
        values.sort((a, b) => {
          const va = a[columnIndex]
          const vb = b[columnIndex]
          if (va < vb) return ascending ? -1 : 1
          if (va > vb) return ascending ? 1 : -1
          return 0
        })

        range.values = [...headers, ...values]
        await context.sync()
        return `Successfully sorted data by column ${columnIndex} (${ascending ? 'ascending' : 'descending'})`
      },
  },

  getWorksheetInfo: {
    name: 'getWorksheetInfo',
    category: 'read',
    description: 'Get information about the active worksheet including name, used range dimensions, and worksheet count.',
    inputSchema: {
      type: 'object',
      properties: {},
      required: [],
    },
    executeExcel: async (context) => {
      
        const sheet = context.workbook.worksheets.getActiveWorksheet()
        sheet.load('name, id, position')
        const usedRange = sheet.getUsedRangeOrNullObject()
        usedRange.load('address, rowCount, columnCount')

        const sheets = context.workbook.worksheets
        sheets.load('items/name')
        await context.sync()

        const sheetNames = sheets.items.map((s: any) => s.name)

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
        )
      },
  },

  getDataFromSheet: {
    name: 'getDataFromSheet',
    category: 'read',
    description: 'Read data from another worksheet without activating it.',
    inputSchema: {
      type: 'object',
      properties: {
        name: {
          type: 'string',
          description: 'Worksheet name to read from.',
        },
        address: {
          type: 'string',
          description: 'Optional range address. Uses used range if omitted.',
        },
      },
      required: ['name'],
    },
    executeExcel: async (context, args: Record<string, any>) => {
      const { name, address } = args as Record<string, any>

        const sheet = await safeGetSheet(context, name)
        const range = address ? sheet.getRange(address) : sheet.getUsedRange()
        range.load('address, values, rowCount, columnCount')
        await context.sync()

        return JSON.stringify(
          {
            worksheet: name,
            address: range.address,
            rowCount: range.rowCount,
            columnCount: range.columnCount,
            values: range.values,
          },
          null,
          2,
        )
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
          description: 'What to clear: "contents" (values only), "formats" (formatting only), or "all" (both)',
          enum: ['contents', 'formats', 'all'],
        },
      },
      required: [],
    },
    executeExcel: async (context, args: Record<string, any>) => {
      const { address, clearType = 'all' } = args
      
        const sheet = context.workbook.worksheets.getActiveWorksheet()
        const range = address ? sheet.getRange(address) : context.workbook.getSelectedRange()

        switch (clearType) {
          case 'contents':
            range.clear(Excel.ClearApplyTo.contents)
            break
          case 'formats':
            range.clear(Excel.ClearApplyTo.formats)
            break
          default:
            range.clear(Excel.ClearApplyTo.all)
        }

        await context.sync()
        return `Successfully cleared ${clearType}${address ? ` from ${address}` : ' from selection'}`
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
      const { searchText, replaceText } = args as Record<string, any>
      
        const sheet = context.workbook.worksheets.getActiveWorksheet()
        const usedRange = sheet.getUsedRange()
        usedRange.load('values, rowCount, columnCount')
        await context.sync()

        let matchCount = 0
        const newValues = usedRange.values.map((row: any[]) =>
          row.map((cell: any) => {
            const cellStr = String(cell)
            if (cellStr.includes(searchText)) {
              matchCount++
              if (replaceText !== undefined) {
                return cellStr.replace(new RegExp(searchText.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g'), replaceText)
              }
            }
            return cell
          }),
        )

        if (replaceText !== undefined && matchCount > 0) {
          usedRange.values = newValues
          await context.sync()
          return `Found and replaced ${matchCount} occurrence(s) of "${searchText}" with "${replaceText}"`
        }

        return `Found ${matchCount} occurrence(s) of "${searchText}"`
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
      const { name } = args as Record<string, any>
      
        const sheet = context.workbook.worksheets.add(name || undefined)
        sheet.activate()
        sheet.load('name')
        await context.sync()
        return `Successfully created and activated worksheet "${sheet.name}"`
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
      } = args as Record<string, any>

      
        const sheet = await safeGetSheet(context, sheetName)

        if (action === 'protect') {
          sheet.protection.protect({
            allowAutoFilter,
            allowFormatCells,
            allowInsertRows,
            selectionMode: Excel.ProtectionSelectionMode.normal,
          }, password)
        } else {
          sheet.protection.unprotect(password)
        }

        await context.sync()
        return `Successfully ${action === 'protect' ? 'protected' : 'unprotected'} worksheet "${sheetName ?? 'active'}"`
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
    executeExcel: async (context) => {
      
        const names = context.workbook.names
        names.load('items/name,items/formula,items/value')
        await context.sync()

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
        )
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
      const { name, rangeAddress } = args as Record<string, any>

        context.workbook.names.add(name, rangeAddress)
        await context.sync()
        return `Successfully set named range "${name}" = ${rangeAddress}`
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
          description: 'If true, clear existing conditional formats on the target range before applying the new rule.',
        },
        operator: {
          type: 'string',
          description: 'Operator for cellValue rule.',
          enum: ['between', 'notBetween', 'equalTo', 'notEqualTo', 'greaterThan', 'greaterThanOrEqual', 'lessThan', 'lessThanOrEqual'],
        },
        formula1: {
          type: 'string',
          description: 'First formula/value for cellValue or custom rule (examples: "100", "=A2>AVERAGE($A$2:$A$100)").',
        },
        formula2: {
          type: 'string',
          description: 'Second formula/value for between/notBetween operators.',
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
          enum: ['threeTrafficLights1', 'threeArrows', 'threeSymbols', 'fourArrows', 'fourTrafficLights', 'fiveArrows'],
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
      } = args

      
        const sheet = context.workbook.worksheets.getActiveWorksheet()
        const targetRange = sheet.getRange(address)
        const conditionalFormats: any = targetRange.conditionalFormats

        if (clearExisting) {
          conditionalFormats.clearAll()
        }

        const ruleTypeMap: Record<string, any> = {
          cellValue: Excel.ConditionalFormatType.cellValue,
          containsText: Excel.ConditionalFormatType.containsText,
          custom: Excel.ConditionalFormatType.custom,
          colorScale: Excel.ConditionalFormatType.colorScale,
          dataBar: Excel.ConditionalFormatType.dataBar,
          iconSet: Excel.ConditionalFormatType.iconSet,
        }

        const selectedType = ruleTypeMap[ruleType]
        if (!selectedType) {
          throw new Error(`Unsupported conditional formatting ruleType: ${ruleType}`)
        }

        const cf: any = conditionalFormats.add(selectedType)

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
          }
          cf.cellValue.rule = {
            formula1,
            ...(formula2 ? { formula2 } : {}),
            operator: operatorMap[operator] ?? Excel.ConditionalCellValueOperator.greaterThan,
          }
        } else if (ruleType === 'containsText') {
          const textOperatorMap: Record<string, any> = {
            contains: Excel.ConditionalTextOperator.contains,
            beginsWith: Excel.ConditionalTextOperator.beginsWith,
            endsWith: Excel.ConditionalTextOperator.endsWith,
            notContains: Excel.ConditionalTextOperator.notContains,
          }
          cf.textComparison.rule = {
            operator: textOperatorMap[textOperator] ?? Excel.ConditionalTextOperator.contains,
            text: text ?? '',
          }
        } else if (ruleType === 'custom') {
          cf.custom.rule.formula = formula1
        } else if (ruleType === 'colorScale') {
          cf.colorScale.criteria = {
            minimum: { color: colorScaleMinColor, type: Excel.ConditionalFormatColorCriterionType.lowestValue },
            midpoint: { color: colorScaleMidColor, type: Excel.ConditionalFormatColorCriterionType.percentile, formula: '50' },
            maximum: { color: colorScaleMaxColor, type: Excel.ConditionalFormatColorCriterionType.highestValue },
          }
        } else if (ruleType === 'dataBar') {
          cf.dataBar.barColor = dataBarColor
          cf.dataBar.lowerBoundRule = { type: Excel.ConditionalFormatRuleType.lowestValue }
          cf.dataBar.upperBoundRule = { type: Excel.ConditionalFormatRuleType.highestValue }
        } else if (ruleType === 'iconSet') {
          const iconSetMap: Record<string, any> = {
            threeTrafficLights1: Excel.IconSet.threeTrafficLights1,
            threeArrows: Excel.IconSet.threeArrows,
            threeSymbols: Excel.IconSet.threeSymbols,
            fourArrows: Excel.IconSet.fourArrows,
            fourTrafficLights: Excel.IconSet.fourTrafficLights,
            fiveArrows: Excel.IconSet.fiveArrows,
          }
          cf.iconSet.style = iconSetMap[iconSetStyle] ?? Excel.IconSet.threeTrafficLights1
        }

        const applyTextAndFillFormat = (format: any) => {
          if (!format) return
          if (fillColor) format.fill.color = fillColor
          if (fontColor) format.font.color = fontColor
          if (bold !== undefined) format.font.bold = bold
        }

        if (ruleType === 'cellValue') {
          applyTextAndFillFormat(cf.cellValue?.format)
        } else if (ruleType === 'containsText') {
          applyTextAndFillFormat(cf.textComparison?.format)
        } else if (ruleType === 'custom') {
          applyTextAndFillFormat(cf.custom?.format)
        }

        if (stopIfTrue !== undefined) cf.stopIfTrue = stopIfTrue

        await context.sync()
        return `Successfully applied ${ruleType} conditional formatting on ${address}`
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
      const { address } = args as Record<string, any>

        const sheet = context.workbook.worksheets.getActiveWorksheet()
        const targetRange = address ? sheet.getRange(address) : sheet.getUsedRangeOrNullObject()
        targetRange.load('address,isNullObject')
        await context.sync()

        if (targetRange.isNullObject) {
          return 'No used range found on the active worksheet, so no conditional formatting rules were read.'
        }

        const conditionalFormats = targetRange.conditionalFormats
        conditionalFormats.load('items/type,items/priority,items/stopIfTrue')
        await context.sync()

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
        )
      },
  },


  findData: {
    name: 'findData',
    category: 'read',
    description: 'Find text or values across the spreadsheet. Returns matching cells with their addresses and values. Options for regex, match case, and entire cell match. Returns up to 2000 results; when truncated, totalMatches indicates how many were found in total.',
    inputSchema: {
      type: 'object',
      properties: {
        searchTerm: { type: 'string', description: 'The text or pattern to search for' },
        matchCase: { type: 'boolean', description: 'Case sensitive. Default: false' },
        matchEntireCell: { type: 'boolean', description: 'Match entire cell content. Default: false' },
        useRegex: { type: 'boolean', description: 'Use regex pattern. Default: false' },
      },
      required: ['searchTerm'],
    },
    executeExcel: async (context, args: Record<string, any>) => {
      const { searchTerm, matchCase = false, matchEntireCell = false, useRegex = false } = args as Record<string, any>
      const MAX_RESULTS = 2000

      let pattern: RegExp | null = null
      if (useRegex) {
        try {
          pattern = new RegExp(searchTerm, matchCase ? '' : 'i')
        } catch {
          return JSON.stringify({ error: `Invalid regex pattern: "${searchTerm}". Please provide a valid regular expression.` })
        }
      }

      const sheets = context.workbook.worksheets
      sheets.load('items')
      await context.sync()

      const matches: any[] = []
      let totalMatches = 0

      for (const sheet of sheets.items) {
        sheet.load('name')
        const usedRange = sheet.getUsedRangeOrNullObject()
        usedRange.load('values,address,rowCount,columnCount')
        await context.sync()

        if (usedRange.isNullObject) continue

        const startMatch = usedRange.address.split('!')[1]?.match(/([A-Z]+)(\d+)/)
        const startCol = startMatch ? startMatch[1].split('').reduce((acc: number, c: string) => acc * 26 + c.charCodeAt(0) - 64, 0) - 1 : 0
        const startRow = startMatch ? parseInt(startMatch[2], 10) - 1 : 0
        const colLetter = (idx: number) => {
          let letter = ''; let temp = idx;
          while (temp >= 0) { letter = String.fromCharCode((temp % 26) + 65) + letter; temp = Math.floor(temp / 26) - 1; }
          return letter
        }

        for (let r = 0; r < usedRange.rowCount; r++) {
          for (let c = 0; c < usedRange.columnCount; c++) {
            const val = usedRange.values[r][c]
            const target = String(val ?? '')
            let isMatch = false
            if (pattern) {
              isMatch = pattern.test(target)
            } else {
              const compVal = matchCase ? target : target.toLowerCase()
              const compTerm = matchCase ? searchTerm : searchTerm.toLowerCase()
              isMatch = matchEntireCell ? compVal === compTerm : compVal.includes(compTerm)
            }
            if (isMatch) {
              totalMatches++
              if (matches.length < MAX_RESULTS) {
                matches.push({ sheet: sheet.name, address: `${colLetter(startCol + c)}${startRow + r + 1}`, value: val })
              }
            }
          }
        }
      }

      const truncated = totalMatches > MAX_RESULTS
      return JSON.stringify(truncated ? { matches, totalMatches, truncated: true } : matches, null, 2)
    },
  },


  getAllObjects: {
    name: 'getAllObjects',
    category: 'read',
    description: 'List all charts and pivot tables. By default scans the active sheet only. Pass allSheets: true to scan all sheets in the workbook (may be slow on large workbooks).',
    inputSchema: {
      type: 'object',
      properties: {
        allSheets: {
          type: 'boolean',
          description: 'When true, list objects from ALL sheets. When false (default), list only the active sheet.',
        },
      },
      required: [],
    },
    executeExcel: async (context, args: Record<string, any>) => {
      const allSheets = args.allSheets === true // default false

      if (!allSheets) {
        const sheet = context.workbook.worksheets.getActiveWorksheet()
        sheet.load('name')
        const charts = sheet.charts
        const pivotTables = sheet.pivotTables
        charts.load('items/name, items/id')
        pivotTables.load('items/name, items/id')
        await context.sync()

        return JSON.stringify(
          {
            charts: charts.items.map((c: any) => ({ name: c.name, id: c.id, sheetName: sheet.name })),
            pivotTables: pivotTables.items.map((p: any) => ({ name: p.name, id: p.id, sheetName: sheet.name })),
          },
          null,
          2,
        )
      }

      // Workbook-wide scan
      const worksheets = context.workbook.worksheets
      worksheets.load('items/name')
      await context.sync()

      for (const sheet of worksheets.items) {
        sheet.charts.load('items/name, items/id')
        sheet.pivotTables.load('items/name, items/id')
      }
      await context.sync()

      const allCharts: any[] = []
      const allPivots: any[] = []
      for (const sheet of worksheets.items) {
        for (const c of sheet.charts.items) {
          allCharts.push({ name: c.name, id: c.id, sheetName: sheet.name })
        }
        for (const p of sheet.pivotTables.items) {
          allPivots.push({ name: p.name, id: p.id, sheetName: sheet.name })
        }
      }

      return JSON.stringify({ charts: allCharts, pivotTables: allPivots }, null, 2)
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
5. Values MUST be 2D arrays: \`range.values = [[value]]\``,
    inputSchema: {
      type: 'object',
      properties: {
        code: {
          type: 'string',
          description: 'JavaScript code following the template. Must include load(), sync(), and try/catch.',
        },
        explanation: {
          type: 'string',
          description: 'Brief explanation of what this code does (required for audit trail).',
        },
      },
      required: ['code', 'explanation'],
    },
    executeExcel: async (context: Excel.RequestContext, args: Record<string, any>) => {
      const { code, explanation } = args

      // Validate code BEFORE execution
      const validation = validateOfficeCode(code, 'Excel')

      if (!validation.valid) {
        return JSON.stringify({
          success: false,
          error: 'Code validation failed. Fix the errors below and try again.',
          validationErrors: validation.errors,
          validationWarnings: validation.warnings,
          suggestion: 'Refer to the Office.js skill document for correct patterns. Remember: Excel values must be 2D arrays.',
          codeReceived: code.slice(0, 300) + (code.length > 300 ? '...' : ''),
        }, null, 2)
      }

      // Log warnings but proceed
      if (validation.warnings.length > 0) {
        console.warn('[eval_officejs] Validation warnings:', validation.warnings)
      }

      try {
        // Execute in sandbox with host restriction
        const result = await sandboxedEval(
          code,
          {
            context,
            Excel: typeof Excel !== 'undefined' ? Excel : undefined,
            Office: typeof Office !== 'undefined' ? Office : undefined,
          },
          'Excel'  // Restrict to Excel namespace only
        )

        return JSON.stringify({
          success: true,
          result: result ?? null,
          explanation,
          warnings: validation.warnings.length > 0 ? validation.warnings : undefined,
        }, null, 2)
      } catch (err: any) {
        return JSON.stringify({
          success: false,
          error: err.message || String(err),
          explanation,
          codeExecuted: code.slice(0, 200) + '...',
          hint: 'Check that all properties are loaded before access, and context.sync() is called.',
        }, null, 2)
      }
    },
  },
}, (def) => async (args = {}) => {
  try {
    return await runExcel(ctx => def.executeExcel(ctx, args))
  } catch (error: any) {
    return JSON.stringify({
      error: true,
      message: error.message || String(error),
      tool: def.name,
      suggestion: 'Fix the error parameters or context and try again.'
    }, null, 2)
  }
})

export function getToolDefinitions(): ToolDefinition[] {
  return Object.values(excelToolDefinitions)
}

export const getExcelToolDefinitions = getToolDefinitions

export { excelToolDefinitions }
