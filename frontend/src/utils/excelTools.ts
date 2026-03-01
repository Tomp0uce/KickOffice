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
  | 'addDataValidation'
  | 'createTable'
  | 'copyRange'
  | 'insertFormula'
  | 'fillFormulaDown'
  | 'createChart'
  | 'formatRange'
  | 'sortRange'
  | 'applyAutoFilter'
  | 'removeAutoFilter'
  | 'getWorksheetInfo'
  | 'renameWorksheet'
  | 'deleteWorksheet'
  | 'activateWorksheet'
  | 'getDataFromSheet'
  | 'freezePanes'
  | 'addHyperlink'
  | 'addCellComment'
  | 'insertRow'
  | 'insertColumn'
  | 'deleteRow'
  | 'deleteColumn'
  | 'mergeCells'
  | 'setCellNumberFormat'
  | 'clearRange'
  | 'getCellFormula'
  | 'searchAndReplace'
  | 'autoFitColumns'
  | 'addWorksheet'
  | 'setColumnWidth'
  | 'setRowHeight'
  | 'protectWorksheet'
  | 'getNamedRanges'
  | 'setNamedRange'
  | 'applyConditionalFormatting'
  | 'getConditionalFormattingRules'
  | 'batchSetCellValues'
  | 'batchProcessRange'
  | 'findData'
  | 'duplicateWorksheet'
  | 'hideUnhideRowColumn'
  | 'getAllObjects'
  | 'modifyObject'


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
    executeExcel: async (context, args) => {
      const { address, value } = args
      
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

  addDataValidation: {
    name: 'addDataValidation',
    category: 'write',
    description:
      'Apply data validation rules to a range (dropdown list, number/date limits, text length, or custom formula).',
    inputSchema: {
      type: 'object',
      properties: {
        address: {
          type: 'string',
          description: 'Target range address (e.g., "A2:A100"). Uses selection if omitted.',
        },
        validationType: {
          type: 'string',
          description: 'Validation type to apply.',
          enum: ['list', 'wholeNumber', 'decimal', 'date', 'textLength', 'custom'],
        },
        listSource: {
          type: 'string',
          description: 'For list validation, comma-separated values or a range reference (e.g., "A,B,C" or "=$F$1:$F$10").',
        },
        operator: {
          type: 'string',
          description: 'Comparison operator for number/date/textLength validations.',
          enum: ['between', 'notBetween', 'equalTo', 'notEqualTo', 'greaterThan', 'greaterThanOrEqual', 'lessThan', 'lessThanOrEqual'],
        },
        formula1: {
          type: 'string',
          description: 'Primary formula/value for validation rule.',
        },
        formula2: {
          type: 'string',
          description: 'Secondary formula/value for between/notBetween.',
        },
      },
      required: [],
    },
    executeExcel: async (context, args) => {
      const {
        address,
        validationType = 'list',
        listSource = 'A,B,C',
        operator = 'between',
        formula1 = '0',
        formula2,
      } = args

      
        const sheet = context.workbook.worksheets.getActiveWorksheet()
        const range = address ? sheet.getRange(address) : context.workbook.getSelectedRange()

        const operatorMap: Record<string, any> = {
          between: Excel.DataValidationOperator.between,
          notBetween: Excel.DataValidationOperator.notBetween,
          equalTo: Excel.DataValidationOperator.equalTo,
          notEqualTo: Excel.DataValidationOperator.notEqualTo,
          greaterThan: Excel.DataValidationOperator.greaterThan,
          greaterThanOrEqual: Excel.DataValidationOperator.greaterThanOrEqualTo,
          lessThan: Excel.DataValidationOperator.lessThan,
          lessThanOrEqual: Excel.DataValidationOperator.lessThanOrEqualTo,
        }

        if (validationType === 'list') {
          range.dataValidation.rule = { list: { inCellDropDown: true, source: listSource } }
        } else if (validationType === 'custom') {
          range.dataValidation.rule = { custom: { formula: formula1 } }
        } else {
          const validationBody = {
            formula1,
            ...(formula2 ? { formula2 } : {}),
            operator: operatorMap[operator] ?? Excel.DataValidationOperator.between,
          }

          const rule: Record<string, any> = {}
          if (validationType === 'wholeNumber') rule.wholeNumber = validationBody
          if (validationType === 'decimal') rule.decimal = validationBody
          if (validationType === 'date') rule.date = validationBody
          if (validationType === 'textLength') rule.textLength = validationBody
          range.dataValidation.rule = rule
        }

        range.dataValidation.errorAlert = {
          title: 'Invalid value',
          message: 'The entered value does not match data validation rules.',
          style: Excel.DataValidationAlertStyle.stop,
          showAlert: true,
        }

        await context.sync()
        return `Successfully applied ${validationType} data validation${address ? ` on ${address}` : ' on selection'}`
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
    executeExcel: async (context, args) => {
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

  copyRange: {
    name: 'copyRange',
    category: 'write',
    description:
      'Copy values, formulas, and number formats from a source range to a destination range.',
    inputSchema: {
      type: 'object',
      properties: {
        sourceAddress: {
          type: 'string',
          description: 'Source range address (e.g., "A1:C20").',
        },
        destinationAddress: {
          type: 'string',
          description: 'Top-left destination address or destination range (e.g., "E1").',
        },
      },
      required: ['sourceAddress', 'destinationAddress'],
    },
    executeExcel: async (context, args) => {
      const { sourceAddress, destinationAddress } = args
      
        const sheet = context.workbook.worksheets.getActiveWorksheet()
        const sourceRange = sheet.getRange(sourceAddress)
        sourceRange.load('values, formulas, numberFormat, rowCount, columnCount')
        await context.sync()

        const destinationRange = sheet.getRange(destinationAddress).getResizedRange(sourceRange.rowCount - 1, sourceRange.columnCount - 1)
        destinationRange.values = sourceRange.values
        destinationRange.formulas = sourceRange.formulas
        destinationRange.numberFormat = sourceRange.numberFormat

        await context.sync()
        return `Successfully copied ${sourceAddress} to ${destinationAddress}`
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
    executeExcel: async (context, args) => {
      const { address, formula } = args
      
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
    executeExcel: async (context, args) => {
      const { startCell, endCell, formula } = args
      
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

  createChart: {
    name: 'createChart',
    category: 'write',
    description:
      'Create a chart from the currently selected data range. Supports various chart types.',
    inputSchema: {
      type: 'object',
      properties: {
        chartType: {
          type: 'string',
          description: 'Type of chart to create',
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
          description: 'Optional chart title',
        },
      },
      required: ['chartType'],
    },
    executeExcel: async (context, args) => {
      const { chartType, title } = args
      
        const range = context.workbook.getSelectedRange()
        range.load('address')
        await context.sync()

        const sheet = range.worksheet
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

        const excelChartType = chartTypeMap[chartType] || Excel.ChartType.columnClustered
        const chart = sheet.charts.add(excelChartType, range, Excel.ChartSeriesBy.auto)

        if (title) {
          chart.title.text = title
        }

        chart.width = 400
        chart.height = 300

        await context.sync()
        return `Successfully created ${chartType} chart${title ? ` with title "${title}"` : ''}`
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
    executeExcel: async (context, args) => {
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
      } = args
      
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
    executeExcel: async (context, args) => {
      const { columnIndex = 0, ascending = true, hasHeaders = true } = args
      
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

  applyAutoFilter: {
    name: 'applyAutoFilter',
    category: 'write',
    description: 'Apply auto filter to the selected range so users can filter data by column values.',
    inputSchema: {
      type: 'object',
      properties: {},
      required: [],
    },
    executeExcel: async (context) => {
      
        const range = context.workbook.getSelectedRange()
        range.load('address')
        await context.sync()

        try {
          const sheet = range.worksheet
          sheet.autoFilter.apply(range)
          await context.sync()
          return 'Successfully applied auto filter to the selected range'
        } catch {
          // Fallback: highlight header row
          const headerRow = range.getRow(0)
          headerRow.format.fill.color = '#0078d4'
          headerRow.format.font.color = '#FFFFFF'
          headerRow.format.font.bold = true
          await context.sync()
          return 'Applied header formatting (auto filter API not available in this context)'
        }
      },
  },

  removeAutoFilter: {
    name: 'removeAutoFilter',
    category: 'write',
    description: 'Remove auto filter from a worksheet.',
    inputSchema: {
      type: 'object',
      properties: {
        sheetName: {
          type: 'string',
          description: 'Optional worksheet name. Uses active worksheet if omitted.',
        },
      },
      required: [],
    },
    executeExcel: async (context, args) => {
      const { sheetName } = args
      
        const sheet = sheetName
          ? context.workbook.worksheets.getItem(sheetName)
          : context.workbook.worksheets.getActiveWorksheet()
        sheet.autoFilter.remove()
        await context.sync()
        return `Successfully removed auto filter from worksheet "${sheetName ?? 'active'}"`
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

  renameWorksheet: {
    name: 'renameWorksheet',
    category: 'write',
    description: 'Rename an existing worksheet.',
    inputSchema: {
      type: 'object',
      properties: {
        currentName: {
          type: 'string',
          description: 'Current worksheet name.',
        },
        newName: {
          type: 'string',
          description: 'New worksheet name.',
        },
      },
      required: ['currentName', 'newName'],
    },
    executeExcel: async (context, args) => {
      const { currentName, newName } = args
      
        const sheet = context.workbook.worksheets.getItem(currentName)
        sheet.name = newName
        await context.sync()
        return `Successfully renamed worksheet "${currentName}" to "${newName}"`
      },
  },

  deleteWorksheet: {
    name: 'deleteWorksheet',
    category: 'write',
    description: 'Delete a worksheet by name.',
    inputSchema: {
      type: 'object',
      properties: {
        name: {
          type: 'string',
          description: 'Worksheet name to delete.',
        },
      },
      required: ['name'],
    },
    executeExcel: async (context, args) => {
      const { name } = args
      
        const sheet = context.workbook.worksheets.getItem(name)
        sheet.delete()
        await context.sync()
        return `Successfully deleted worksheet "${name}"`
      },
  },

  activateWorksheet: {
    name: 'activateWorksheet',
    category: 'write',
    description: 'Activate a worksheet by name.',
    inputSchema: {
      type: 'object',
      properties: {
        name: {
          type: 'string',
          description: 'Worksheet name to activate.',
        },
      },
      required: ['name'],
    },
    executeExcel: async (context, args) => {
      const { name } = args
      
        const sheet = context.workbook.worksheets.getItem(name)
        sheet.activate()
        await context.sync()
        return `Successfully activated worksheet "${name}"`
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
    executeExcel: async (context, args) => {
      const { name, address } = args
      
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

  freezePanes: {
    name: 'freezePanes',
    category: 'write',
    description: 'Freeze or unfreeze worksheet panes by rows, columns, or anchor cell.',
    inputSchema: {
      type: 'object',
      properties: {
        mode: {
          type: 'string',
          description: 'Freeze mode to apply.',
          enum: ['rows', 'columns', 'at', 'unfreeze'],
        },
        count: {
          type: 'number',
          description: 'Row or column count for rows/columns mode.',
        },
        address: {
          type: 'string',
          description: 'Anchor range address for mode="at" (e.g., "C3").',
        },
      },
      required: ['mode'],
    },
    executeExcel: async (context, args) => {
      const { mode, count = 1, address } = args
      
        const sheet = context.workbook.worksheets.getActiveWorksheet()

        if (mode === 'rows') {
          sheet.freezePanes.freezeRows(count)
        } else if (mode === 'columns') {
          sheet.freezePanes.freezeColumns(count)
        } else if (mode === 'at') {
          if (!address) {
            throw new Error('address is required when mode is "at"')
          }
          sheet.freezePanes.freezeAt(sheet.getRange(address))
        } else {
          sheet.freezePanes.unfreeze()
        }

        await context.sync()
        return `Successfully applied freeze panes mode "${mode}"`
      },
  },

  addHyperlink: {
    name: 'addHyperlink',
    category: 'write',
    description: 'Add a clickable hyperlink to a cell or range.',
    inputSchema: {
      type: 'object',
      properties: {
        address: {
          type: 'string',
          description: 'Target cell/range address (e.g., "A1").',
        },
        hyperlinkAddress: {
          type: 'string',
          description: 'Hyperlink URL or mailto value.',
        },
        textToDisplay: {
          type: 'string',
          description: 'Optional displayed text.',
        },
        screenTip: {
          type: 'string',
          description: 'Optional tooltip text.',
        },
      },
      required: ['address', 'hyperlinkAddress'],
    },
    executeExcel: async (context, args) => {
      const { address, hyperlinkAddress, textToDisplay, screenTip } = args
      
        const sheet = context.workbook.worksheets.getActiveWorksheet()
        const range = sheet.getRange(address)
        range.hyperlink = {
          address: hyperlinkAddress,
          ...(textToDisplay ? { textToDisplay } : {}),
          ...(screenTip ? { screenTip } : {}),
        }

        await context.sync()
        return `Successfully added hyperlink to ${address}`
      },
  },

  addCellComment: {
    name: 'addCellComment',
    category: 'write',
    description: 'Add a comment (note) to a cell range.',
    inputSchema: {
      type: 'object',
      properties: {
        address: {
          type: 'string',
          description: 'Target cell/range address (e.g., "B2").',
        },
        text: {
          type: 'string',
          description: 'Comment text.',
        },
      },
      required: ['address', 'text'],
    },
    executeExcel: async (context, args) => {
      const { address, text } = args
      
        const comments = context.workbook.comments
        comments.add(address, text)
        await context.sync()
        return `Successfully added comment to ${address}`
      },
  },

  insertRow: {
    name: 'insertRow',
    category: 'write',
    description: 'Insert one or more rows at the specified position.',
    inputSchema: {
      type: 'object',
      properties: {
        rowIndex: {
          type: 'number',
          description: 'The 1-based row number where to insert (e.g., 5 inserts before row 5)',
        },
        count: {
          type: 'number',
          description: 'Number of rows to insert (default: 1)',
        },
      },
      required: ['rowIndex'],
    },
    executeExcel: async (context, args) => {
      const { rowIndex, count = 1 } = args
      
        const sheet = context.workbook.worksheets.getActiveWorksheet()
        const range = sheet.getRange(`${rowIndex}:${rowIndex + count - 1}`)
        range.insert(Excel.InsertShiftDirection.down)
        await context.sync()
        return `Successfully inserted ${count} row(s) at position ${rowIndex}`
      },
  },

  insertColumn: {
    name: 'insertColumn',
    category: 'write',
    description: 'Insert one or more columns at the specified position.',
    inputSchema: {
      type: 'object',
      properties: {
        columnLetter: {
          type: 'string',
          description: 'Column letter where to insert (e.g., "C" inserts before column C)',
        },
        count: {
          type: 'number',
          description: 'Number of columns to insert (default: 1)',
        },
      },
      required: ['columnLetter'],
    },
    executeExcel: async (context, args) => {
      const { columnLetter, count = 1 } = args
      
        const sheet = context.workbook.worksheets.getActiveWorksheet()
        const startColNum = colToInt(columnLetter.toUpperCase())
        const endCol = intToCol(startColNum + count - 1)
        const range = sheet.getRange(`${columnLetter}:${endCol}`)
        range.insert(Excel.InsertShiftDirection.right)
        await context.sync()
        return `Successfully inserted ${count} column(s) at position ${columnLetter}`
      },
  },

  deleteRow: {
    name: 'deleteRow',
    category: 'write',
    description: 'Delete one or more rows at the specified position.',
    inputSchema: {
      type: 'object',
      properties: {
        rowIndex: {
          type: 'number',
          description: 'The 1-based row number to delete',
        },
        count: {
          type: 'number',
          description: 'Number of rows to delete (default: 1)',
        },
      },
      required: ['rowIndex'],
    },
    executeExcel: async (context, args) => {
      const { rowIndex, count = 1 } = args
      
        const sheet = context.workbook.worksheets.getActiveWorksheet()
        const range = sheet.getRange(`${rowIndex}:${rowIndex + count - 1}`)
        range.delete(Excel.DeleteShiftDirection.up)
        await context.sync()
        return `Successfully deleted ${count} row(s) at position ${rowIndex}`
      },
  },

  deleteColumn: {
    name: 'deleteColumn',
    category: 'write',
    description: 'Delete one or more columns at the specified position.',
    inputSchema: {
      type: 'object',
      properties: {
        columnLetter: {
          type: 'string',
          description: 'Column letter to delete (e.g., "C")',
        },
        count: {
          type: 'number',
          description: 'Number of columns to delete (default: 1)',
        },
      },
      required: ['columnLetter'],
    },
    executeExcel: async (context, args) => {
      const { columnLetter, count = 1 } = args
      
        const sheet = context.workbook.worksheets.getActiveWorksheet()
        const startColNum = colToInt(columnLetter.toUpperCase())
        const endCol = intToCol(startColNum + count - 1)
        const range = sheet.getRange(`${columnLetter}:${endCol}`)
        range.delete(Excel.DeleteShiftDirection.left)
        await context.sync()
        return `Successfully deleted ${count} column(s) at position ${columnLetter}`
      },
  },

  mergeCells: {
    name: 'mergeCells',
    category: 'format',
    description: 'Merge or unmerge the selected cells or a specific range.',
    inputSchema: {
      type: 'object',
      properties: {
        address: {
          type: 'string',
          description: 'Optional range address (e.g., "A1:C1"). Uses selection if not provided.',
        },
        merge: {
          type: 'boolean',
          description: 'True to merge, false to unmerge (default: true)',
        },
      },
      required: [],
    },
    executeExcel: async (context, args) => {
      const { address, merge = true } = args
      
        const sheet = context.workbook.worksheets.getActiveWorksheet()
        const range = address ? sheet.getRange(address) : context.workbook.getSelectedRange()

        if (merge) {
          range.merge()
        } else {
          range.unmerge()
        }
        await context.sync()
        return `Successfully ${merge ? 'merged' : 'unmerged'} cells${address ? ` at ${address}` : ''}`
      },
  },

  setCellNumberFormat: {
    name: 'setCellNumberFormat',
    category: 'format',
    description:
      'Set the number format for a range of cells (e.g., currency, percentage, date format).',
    inputSchema: {
      type: 'object',
      properties: {
        address: {
          type: 'string',
          description: 'Optional range address. Uses selection if not provided.',
        },
        format: {
          type: 'string',
          description:
            'Number format string. Examples: "#,##0.00" (number), "$#,##0.00" (currency), "0.00%" (percentage), "yyyy-mm-dd" (date), "0" (integer)',
        },
      },
      required: ['format'],
    },
    executeExcel: async (context, args) => {
      const { address, format } = args
      
        const sheet = context.workbook.worksheets.getActiveWorksheet()
        const range = address ? sheet.getRange(address) : context.workbook.getSelectedRange()
        range.numberFormat = range.values.map((row: any[]) => row.map(() => format))
        await context.sync()
        return `Successfully set number format "${format}"${address ? ` for ${address}` : ' for selection'}`
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
    executeExcel: async (context, args) => {
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

  getCellFormula: {
    name: 'getCellFormula',
    category: 'read',
    description: 'Get the formula (if any) from a specific cell or the selected range.',
    inputSchema: {
      type: 'object',
      properties: {
        address: {
          type: 'string',
          description: 'Optional cell address (e.g., "A1"). Uses selection if not provided.',
        },
      },
      required: [],
    },
    executeExcel: async (context, args) => {
      const { address } = args
      
        const sheet = context.workbook.worksheets.getActiveWorksheet()
        const range = address ? sheet.getRange(address) : context.workbook.getSelectedRange()
        range.load('formulas, formulasLocal, values, address')
        await context.sync()

        return JSON.stringify(
          {
            address: range.address,
            formulas: range.formulas,
            formulasLocal: range.formulasLocal,
            values: range.values,
          },
          null,
          2,
        )
      },
  },

  searchAndReplace: {
    name: 'searchAndReplace',
    category: 'read',
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
    executeExcel: async (context, args) => {
      const { searchText, replaceText } = args
      
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

  autoFitColumns: {
    name: 'autoFitColumns',
    category: 'format',
    description: 'Auto-fit column widths to match their content for the selected range or the entire used range.',
    inputSchema: {
      type: 'object',
      properties: {
        address: {
          type: 'string',
          description: 'Optional range address. If omitted, auto-fits the used range columns.',
        },
      },
      required: [],
    },
    executeExcel: async (context, args) => {
      const { address } = args
      
        const sheet = context.workbook.worksheets.getActiveWorksheet()
        const range = address ? sheet.getRange(address) : sheet.getUsedRange()
        range.format.autofitColumns()
        range.format.autofitRows()
        await context.sync()
        return `Successfully auto-fitted columns and rows${address ? ` for ${address}` : ''}`
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
    executeExcel: async (context, args) => {
      const { name } = args
      
        const sheet = context.workbook.worksheets.add(name || undefined)
        sheet.activate()
        sheet.load('name')
        await context.sync()
        return `Successfully created and activated worksheet "${sheet.name}"`
      },
  },

  setColumnWidth: {
    name: 'setColumnWidth',
    category: 'format',
    description: 'Set the width of one or more columns.',
    inputSchema: {
      type: 'object',
      properties: {
        columnLetter: {
          type: 'string',
          description: 'Column letter (e.g., "A") or range (e.g., "A:D")',
        },
        width: {
          type: 'number',
          description: 'Column width in points',
        },
      },
      required: ['columnLetter', 'width'],
    },
    executeExcel: async (context, args) => {
      const { columnLetter, width } = args
      
        const sheet = context.workbook.worksheets.getActiveWorksheet()
        const colRange = columnLetter.includes(':')
          ? sheet.getRange(columnLetter)
          : sheet.getRange(`${columnLetter}:${columnLetter}`)
        colRange.format.columnWidth = width
        await context.sync()
        return `Successfully set column ${columnLetter} width to ${width}`
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
    executeExcel: async (context, args) => {
      const { rowIndex, height } = args
      
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
    executeExcel: async (context, args) => {
      const {
        sheetName,
        action,
        password,
        allowAutoFilter = false,
        allowFormatCells = false,
        allowInsertRows = false,
      } = args

      
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
    executeExcel: async (context, args) => {
      const { name, rangeAddress } = args
      
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
    executeExcel: async (context, args) => {
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
    executeExcel: async (context, args) => {
      const { address } = args
      
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
    executeExcel: async (context, args) => {
      const { cells } = args
      if (!Array.isArray(cells) || cells.length === 0) {
        return 'Error: cells array is empty or invalid'
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
    executeExcel: async (context, args) => {
      const { address, values } = args
      if (!Array.isArray(values) || values.length === 0) {
        return 'Error: values array is empty or invalid'
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
    executeExcel: async (context, args) => {
      const { searchTerm, matchCase = false, matchEntireCell = false, useRegex = false } = args
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
    executeExcel: async (context, args) => {
      const { sourceName, newName } = args
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
    executeExcel: async (context, args) => {
      const { dimension, reference, action } = args
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
    description: 'Get all charts and pivot tables on the active worksheet.',
    inputSchema: {
      type: 'object',
      properties: {},
      required: []
    },
    executeExcel: async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet()
      const charts = sheet.charts
      const pivotTables = sheet.pivotTables
      charts.load('items/name, items/id')
      pivotTables.load('items/name, items/id')
      await context.sync()

      const result = {
        charts: charts.items.map(c => ({ name: c.name, id: c.id })),
        pivotTables: pivotTables.items.map(p => ({ name: p.name, id: p.id }))
      }
      return JSON.stringify(result, null, 2)
    }
  },

  modifyObject: {
    name: 'modifyObject',
    category: 'write',
    description: 'Delete a chart or pivot table by its name or ID.',
    inputSchema: {
      type: 'object',
      properties: {
        objectType: { type: 'string', enum: ['chart', 'pivotTable'], description: 'The type of object' },
        name: { type: 'string', description: 'The name of the object to delete' }
      },
      required: ['objectType', 'name']
    },
    executeExcel: async (context, args) => {
      const { objectType, name } = args
      const sheet = context.workbook.worksheets.getActiveWorksheet()
      
      if (objectType === 'chart') {
        const chart = sheet.charts.getItem(name)
        chart.delete()
      } else {
        const pivot = sheet.pivotTables.getItem(name)
        pivot.delete()
      }
      await context.sync()
      return `Successfully deleted ${objectType} ${name}`
    }
  },
})

export function getExcelToolDefinitions(): ExcelToolDefinition[] {
  return Object.values(excelToolDefinitions)
}

export function getExcelTool(name: ExcelToolName): ExcelToolDefinition | undefined {
  return excelToolDefinitions[name]
}

export { excelToolDefinitions }
