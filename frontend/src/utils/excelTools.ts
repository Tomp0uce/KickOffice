import type { ToolProperty, ExcelToolDefinition } from '@/types'
import { executeOfficeAction } from './officeAction'
import { localStorageKey } from './enum'
import { sandboxedEval } from './sandbox'

const runExcel = <T>(action: (context: Excel.RequestContext) => Promise<T>): Promise<T> =>
  executeOfficeAction(() => Excel.run(action))

type ExcelToolTemplate = Omit<ExcelToolDefinition, 'execute'> & {
  executeExcel: (context: Excel.RequestContext, args: Record<string, any>) => Promise<string>
}

function createExcelTools(definitions: Record<ExcelToolName, ExcelToolTemplate>): Record<ExcelToolName, ExcelToolDefinition> {
  return Object.fromEntries(
    Object.entries(definitions).map(([name, definition]) => [
      name,
      {
        ...definition,
        execute: async (args: Record<string, any> = {}) => runExcel(context => definition.executeExcel(context, args)),
      },
    ]),
  ) as unknown as Record<ExcelToolName, ExcelToolDefinition>
}

export type ExcelToolName =
  | 'getSelectedCells'
  | 'setCellValue'
  | 'getWorksheetData'
  | 'createTable'
  | 'insertFormula'
  | 'fillFormulaDown'
  | 'formatRange'
  | 'sortRange'
  | 'getWorksheetInfo'
  | 'getDataFromSheet'
  | 'clearRange'
  | 'searchAndReplace'
  | 'addWorksheet'
  | 'getNamedRanges'
  | 'applyConditionalFormatting'
  | 'batchSetCellValues'
  | 'batchProcessRange'
  | 'findData'
  | 'getAllObjects'
  | 'manageObject'
  | 'eval_officejs'


function getExcelFormulaLanguage(): 'en' | 'fr' {
  const configured = localStorage.getItem(localStorageKey.excelFormulaLanguage)
  return configured === 'fr' ? 'fr' : 'en'
}

function colToInt(col: string): number {
  let num = 0
  for (let i = 0; i < col.length; i++) {
    num = num * 26 + (col.charCodeAt(i) - 64)
  }
  return num
}

function intToCol(num: number): string {
  let col = ''
  while (num > 0) {
    const mod = (num - 1) % 26
    col = String.fromCharCode(65 + mod) + col
    num = Math.floor((num - mod) / 26)
  }
  return col
}

const excelToolDefinitions = createExcelTools({
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

  setCellValue: {
    name: 'setCellValue',
    category: 'write',
    description:
      'Set a value in a specific cell or range. Use A1-style notation for the address (e.g., "A1", "B2:D5").',
    inputSchema: {
      type: 'object',
      properties: {
        address: {
          type: 'string',
          description: 'Cell address in A1 notation (e.g., "A1", "B2:D5")',
        },
        value: {
          type: 'string',
          description: 'The value to set. For multiple cells, provide a JSON 2D array like [["a","b"],["c","d"]]',
        },
      },
      required: ['address', 'value'],
    },
    executeExcel: async (context, args: Record<string, any>) => {
      const { address, value } = args as Record<string, any>
      
        const sheet = context.workbook.worksheets.getActiveWorksheet()
        const range = sheet.getRange(address)

        // Try to parse as JSON array for multi-cell values
        try {
          const parsed = JSON.parse(value)
          if (Array.isArray(parsed)) {
            range.values = parsed
          } else {
            range.values = [[parsed]]
          }
        } catch {
          // Single value
          const num = Number(value)
          range.values = [[isNaN(num) ? value : num]]
        }

        await context.sync()
        return `Successfully set value at ${address}`
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

  insertFormula: {
    name: 'insertFormula',
    category: 'write',
    description:
      'Insert an Excel formula at a specific cell address. The formula should start with "=" (e.g., "=SUM(A1:A10)", "=AVERAGE(B2:B20)").',
    inputSchema: {
      type: 'object',
      properties: {
        address: {
          type: 'string',
          description: 'Cell address where the formula will be inserted (e.g., "A11")',
        },
        formula: {
          type: 'string',
          description: 'The Excel formula to insert (e.g., "=SUM(A1:A10)")',
        },
      },
      required: ['address', 'formula'],
    },
    executeExcel: async (context, args: Record<string, any>) => {
      const { address, formula } = args as Record<string, any>
      
        const sheet = context.workbook.worksheets.getActiveWorksheet()
        const cell = sheet.getRange(address)
        const formulaLocale = getExcelFormulaLanguage()

        if (formulaLocale === 'fr') {
          cell.formulasLocal = [[formula]]
        } else {
          cell.formulas = [[formula]]
        }

        cell.format.font.bold = true
        await context.sync()
        return `Successfully inserted ${formulaLocale === 'fr' ? 'localized French' : 'English'} formula "${formula}" at ${address}`
      },
  },

  fillFormulaDown: {
    name: 'fillFormulaDown',
    category: 'write',
    description:
      'Insert an Excel formula into the first cell of a range and fill it down to all rows in that range. This is much more efficient than calling insertFormula repeatedly for each row. The formula should reference the first row and Excel will automatically adjust relative references for each subsequent row. For example, to apply "=A2*B2" from C2 to C100, use startCell="C2", endCell="C100", formula="=A2*B2".',
    inputSchema: {
      type: 'object',
      properties: {
        startCell: {
          type: 'string',
          description: 'The first cell address where the formula starts (e.g., "C2")',
        },
        endCell: {
          type: 'string',
          description: 'The last cell address where the formula should be filled to (e.g., "C100")',
        },
        formula: {
          type: 'string',
          description: 'The Excel formula for the first row (e.g., "=A2*B2"). Relative references will auto-adjust for each row.',
        },
      },
      required: ['startCell', 'endCell', 'formula'],
    },
    executeExcel: async (context, args: Record<string, any>) => {
      const { startCell, endCell, formula } = args as Record<string, any>
      
        const sheet = context.workbook.worksheets.getActiveWorksheet()
        const formulaLocale = getExcelFormulaLanguage()

        // Set the formula in the first cell
        const firstCell = sheet.getRange(startCell)
        if (formulaLocale === 'fr') {
          firstCell.formulasLocal = [[formula]]
        } else {
          firstCell.formulas = [[formula]]
        }
        await context.sync()

        // Now select the full range and fill down from the first cell
        const fullRange = sheet.getRange(`${startCell}:${endCell}`)
        fullRange.load('rowCount')
        await context.sync()

        if (fullRange.rowCount > 1) {
          // Use the first cell as source and fill down to the rest
          const sourceRange = sheet.getRange(`${startCell}:${startCell}`)
          sourceRange.autoFill(fullRange, Excel.AutoFillType.fillDefault)
          await context.sync()
        }

        return `Successfully filled formula "${formula}" from ${startCell} to ${endCell}`
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
      const sheet = sheetName
        ? context.workbook.worksheets.getItem(sheetName)
        : context.workbook.worksheets.getActiveWorksheet()

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
          sheet.pivotTables.add(source, destRange, name)
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
    description: 'Sort the selected data range by a specific column.',
    inputSchema: {
      type: 'object',
      properties: {
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
      const { columnIndex = 0, ascending = true, hasHeaders = true } = args as Record<string, any>
      
        const range = context.workbook.getSelectedRange()
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
      
        const sheet = context.workbook.worksheets.getItem(name)
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
    executeExcel: async (context, args: Record<string, any>) => {
      const { name } = args as Record<string, any>
      
        const sheet = context.workbook.worksheets.add(name || undefined)
        sheet.activate()
        sheet.load('name')
        await context.sync()
        return `Successfully created and activated worksheet "${sheet.name}"`
      },
  },

  setRowHeight: {
    name: 'setRowHeight',
    category: 'format',
    description: 'Set the height of one or more rows.',
    inputSchema: {
      type: 'object',
      properties: {
        rowIndex: {
          type: 'number',
          description: 'The 1-based row number',
        },
        height: {
          type: 'number',
          description: 'Row height in points',
        },
      },
      required: ['rowIndex', 'height'],
    },
    executeExcel: async (context, args: Record<string, any>) => {
      const { rowIndex, height } = args as Record<string, any>
      
        const sheet = context.workbook.worksheets.getActiveWorksheet()
        const rowRange = sheet.getRange(`${rowIndex}:${rowIndex}`)
        rowRange.format.rowHeight = height
        await context.sync()
        return `Successfully set row ${rowIndex} height to ${height}`
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
    executeExcel: async (context, args: Record<string, any>) => {
      const {
        sheetName,
        action,
        password,
        allowAutoFilter = false,
        allowFormatCells = false,
        allowInsertRows = false,
      } = args as Record<string, any>

      
        const sheet = sheetName
          ? context.workbook.worksheets.getItem(sheetName)
          : context.workbook.worksheets.getActiveWorksheet()

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
    executeExcel: async (context, args: Record<string, any>) => {
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
    executeExcel: async (context, args: Record<string, any>) => {
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
            rules: conditionalFormats.items.map(rule => ({
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

  batchSetCellValues: {
    name: 'batchSetCellValues',
    category: 'write',
    description:
      'Set values for multiple individual cells in a single operation. Much more efficient than calling setCellValue repeatedly. Use this whenever you need to modify more than 2 cells (e.g., translating, cleaning, or transforming cell contents). Provide an array of {address, value} pairs. For ranges larger than 100 cells, process in chunks of 50-100 cells at a time.',
    inputSchema: {
      type: 'object',
      properties: {
        cells: {
          type: 'array',
          description: 'Array of cell updates. Each item has an "address" (A1 notation) and a "value" (the new cell content).',
          items: {
            type: 'object',
            // @ts-expect-error ToolProperty doesn't natively support nested array items schema
            properties: {
              address: { type: 'string', description: 'Cell address in A1 notation (e.g., "A1")' },
              value: { type: 'string', description: 'New value for the cell' },
            },
            required: ['address', 'value'],
          },
        },
      },
      required: ['cells'],
    },
    executeExcel: async (context, args: Record<string, any>) => {
      const { cells } = args as Record<string, any>
      if (!Array.isArray(cells) || cells.length === 0) {
        throw new Error('Error: cells array is empty or invalid')
      }
      const sheet = context.workbook.worksheets.getActiveWorksheet()
      for (const cell of cells) {
        const range = sheet.getRange(cell.address)
        const num = Number(cell.value)
        range.values = [[isNaN(num) || cell.value === '' ? cell.value : num]]
      }
      await context.sync()
      return `Successfully updated ${cells.length} cells`
    },
  },

  batchProcessRange: {
    name: 'batchProcessRange',
    category: 'write',
    description: 'BETA: Process an entire contiguous range in one go (like A1:C5). Provide a 2D array of values. Empty arrays or nulls will skip the cell update or clear it. OVERWRITE NOTE: Be careful to read the cells first if you are unsure whether they are empty, as this will overwrite any existing data in the specified range.',
    inputSchema: {
      type: 'object',
      properties: {
        address: {
          type: 'string',
          description: 'Target range address in A1 notation (e.g., "A1:A50", "B2:D10")',
        },
        values: {
          type: 'array',
          description: 'A 2D array of new values matching the range dimensions. Example for a single-column range of 3 rows: [["value1"],["value2"],["value3"]].',
          items: {
            type: 'array',
            items: { type: 'string' },
          },
        },
      },
      required: ['address', 'values'],
    },
    executeExcel: async (context, args: Record<string, any>) => {
      const { address, values } = args as Record<string, any>
      if (!Array.isArray(values) || values.length === 0) {
        throw new Error('Error: values array is empty or invalid')
      }
      const sheet = context.workbook.worksheets.getActiveWorksheet()
      const range = sheet.getRange(address)
      range.values = values
      await context.sync()
      return `Successfully updated range ${address} (${values.length} rows × ${values[0]?.length || 0} columns)`
    },
  },

  findData: {
    name: 'findData',
    category: 'read',
    description: 'Find text or values across the spreadsheet. Returns matching cells with their addresses and values. Options for regex, match case, and entire cell match.',
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
      const sheets = context.workbook.worksheets
      sheets.load('items')
      await context.sync()

      const pattern = useRegex ? new RegExp(searchTerm, matchCase ? '' : 'i') : null
      const matches: any[] = []

      for (const sheet of sheets.items) {
        if (matches.length > 200) break // limit
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
         if (matches.length > 200) break
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
              matches.push({ sheet: sheet.name, address: `${colLetter(startCol + c)}${startRow + r + 1}`, value: val })
            }
          }
        }
      }
      return JSON.stringify(matches, null, 2)
    },
  },

  duplicateWorksheet: {
    name: 'duplicateWorksheet',
    category: 'write',
    description: 'Duplicate an existing worksheet.',
    inputSchema: {
      type: 'object',
      properties: {
        sourceName: { type: 'string', description: 'Name of the sheet to duplicate' },
        newName: { type: 'string', description: 'Name for the new copied sheet' }
      },
      required: ['sourceName'],
    },
    executeExcel: async (context, args: Record<string, any>) => {
      const { sourceName, newName } = args as Record<string, any>
      const sheet = context.workbook.worksheets.getItem(sourceName)
      const copy = sheet.copy()
      if (newName) copy.name = newName
      await context.sync()
      return `Successfully duplicated worksheet ${sourceName}${newName ? ' to ' + newName : ''}`
    }
  },

  hideUnhideRowColumn: {
    name: 'hideUnhideRowColumn',
    category: 'format',
    description: 'Hide or unhide specific rows or columns.',
    inputSchema: {
      type: 'object',
      properties: {
        dimension: { type: 'string', enum: ['rows', 'columns'], description: 'Whether to modify rows or columns' },
        reference: { type: 'string', description: 'Row number(s) e.g. "5:10" or "5", or column letter(s) e.g. "C" or "C:E"' },
        action: { type: 'string', enum: ['hide', 'unhide'], description: 'Action to perform' }
      },
      required: ['dimension', 'reference', 'action']
    },
    executeExcel: async (context, args: Record<string, any>) => {
      const { dimension, reference, action } = args as Record<string, any>
      const sheet = context.workbook.worksheets.getActiveWorksheet()
      const isRow = dimension === 'rows'
      
      let refStr = String(reference)
      if (!refStr.includes(':')) refStr = `${refStr}:${refStr}`

      const range = sheet.getRange(refStr)
      if (isRow) {
        range.rowHidden = action === 'hide'
      } else {
        range.columnHidden = action === 'hide'
      }
      await context.sync()
      return `Successfully ${action}d ${dimension} ${reference}`
    }
  },

  getAllObjects: {
    name: 'getAllObjects',
    category: 'read',
    description: 'List all charts and pivot tables. By default scans the entire workbook (all sheets). Pass allSheets: false to limit to the active sheet only.',
    inputSchema: {
      type: 'object',
      properties: {
        allSheets: {
          type: 'boolean',
          description: 'When true (default), list objects from ALL sheets. When false, list only the active sheet.',
        },
      },
      required: [],
    },
    executeExcel: async (context, args: Record<string, any>) => {
      const allSheets = args.allSheets !== false // default true

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
    description: "Execute arbitrary Office.js code within an Excel.run context. Use this as an escape hatch for operations not covered by dedicated tools: sorting, autofilter, freeze panes, hyperlinks, row/column insert/delete/resize/hide, data validation, number formats, cell comments, named ranges, sheet rename/duplicate/protect/activate, autofit, conditional formatting inspection, etc. The code runs inside `Excel.run(async (context) => { ... })` with `context` (Excel.RequestContext) and `Excel` global available. Always call `await context.sync()` before returning.",
    inputSchema: {
      type: 'object',
      properties: {
        code: {
          type: 'string',
          description: "JavaScript code to execute. Has access to `context` (Excel.RequestContext) and `Excel` global. Must be valid async JS. Return a value to get it as result. Example: `const sheet = context.workbook.worksheets.getActiveWorksheet(); sheet.getRange('A1:D100').sort.apply([{key:0,ascending:true}]); await context.sync(); return 'Sorted column A ascending.';`",
        },
        explanation: {
          type: 'string',
          description: 'Brief explanation of what this code does.',
        },
      },
      required: ['code'],
    },
    executeExcel: async (context: Excel.RequestContext, args: Record<string, any>) => {
      const { code } = args as Record<string, any>
      try {
        const result = await sandboxedEval(code, { context, Excel: typeof Excel !== 'undefined' ? Excel : undefined })
        return JSON.stringify({ success: true, result: result ?? null }, null, 2)
      } catch (err: any) {
        return JSON.stringify({ success: false, error: err.message || String(err) }, null, 2)
      }
    },
  },
})

export function getExcelToolDefinitions(): ExcelToolDefinition[] {
  return Object.values(excelToolDefinitions)
}

export { excelToolDefinitions }
