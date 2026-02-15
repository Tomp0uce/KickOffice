import { localStorageKey } from './enum'

export type ExcelToolName =
  | 'getSelectedCells'
  | 'setCellValue'
  | 'getWorksheetData'
  | 'insertFormula'
  | 'createChart'
  | 'formatRange'
  | 'sortRange'
  | 'applyAutoFilter'
  | 'getWorksheetInfo'
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
  | 'applyConditionalFormatting'
  | 'getConditionalFormattingRules'


function getExcelFormulaLanguage(): 'en' | 'fr' {
  const configured = localStorage.getItem(localStorageKey.excelFormulaLanguage)
  return configured === 'fr' ? 'fr' : 'en'
}

const excelToolDefinitions: Record<ExcelToolName, ExcelToolDefinition> = {
  getSelectedCells: {
    name: 'getSelectedCells',
    description:
      'Get the values, address, and dimensions of the currently selected cells in Excel. Returns a JSON object with address, rowCount, columnCount, and the 2D values array.',
    inputSchema: {
      type: 'object',
      properties: {},
      required: [],
    },
    execute: async () => {
      return Excel.run(async (context) => {
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
      })
    },
  },

  setCellValue: {
    name: 'setCellValue',
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
    execute: async (args) => {
      const { address, value } = args
      return Excel.run(async (context) => {
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
      })
    },
  },

  getWorksheetData: {
    name: 'getWorksheetData',
    description:
      'Get all data from the used range of the active worksheet. Returns the values, address, row count, and column count.',
    inputSchema: {
      type: 'object',
      properties: {},
      required: [],
    },
    execute: async () => {
      return Excel.run(async (context) => {
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
      })
    },
  },

  insertFormula: {
    name: 'insertFormula',
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
    execute: async (args) => {
      const { address, formula } = args
      return Excel.run(async (context) => {
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
      })
    },
  },

  createChart: {
    name: 'createChart',
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
    execute: async (args) => {
      const { chartType, title } = args
      return Excel.run(async (context) => {
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
      })
    },
  },

  formatRange: {
    name: 'formatRange',
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
      },
      required: [],
    },
    execute: async (args) => {
      const { address, fillColor, fontColor, bold, italic, fontSize, borders, horizontalAlignment } = args
      return Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet()
        const range = address ? sheet.getRange(address) : context.workbook.getSelectedRange()

        if (fillColor) range.format.fill.color = fillColor
        if (fontColor) range.format.font.color = fontColor
        if (bold !== undefined) range.format.font.bold = bold
        if (italic !== undefined) range.format.font.italic = italic
        if (fontSize) range.format.font.size = fontSize
        if (horizontalAlignment) {
          range.format.horizontalAlignment = horizontalAlignment as Excel.HorizontalAlignment
        }
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
      })
    },
  },

  sortRange: {
    name: 'sortRange',
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
    execute: async (args) => {
      const { columnIndex = 0, ascending = true, hasHeaders = true } = args
      return Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange()
        range.load('values, rowCount, columnCount')
        await context.sync()

        const sortOn = ascending ? Excel.SortOrder.ascending : Excel.SortOrder.descending
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
      })
    },
  },

  applyAutoFilter: {
    name: 'applyAutoFilter',
    description: 'Apply auto filter to the selected range so users can filter data by column values.',
    inputSchema: {
      type: 'object',
      properties: {},
      required: [],
    },
    execute: async () => {
      return Excel.run(async (context) => {
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
      })
    },
  },

  getWorksheetInfo: {
    name: 'getWorksheetInfo',
    description: 'Get information about the active worksheet including name, used range dimensions, and worksheet count.',
    inputSchema: {
      type: 'object',
      properties: {},
      required: [],
    },
    execute: async () => {
      return Excel.run(async (context) => {
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
      })
    },
  },

  insertRow: {
    name: 'insertRow',
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
    execute: async (args) => {
      const { rowIndex, count = 1 } = args
      return Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet()
        const range = sheet.getRange(`${rowIndex}:${rowIndex + count - 1}`)
        range.insert(Excel.InsertShiftDirection.down)
        await context.sync()
        return `Successfully inserted ${count} row(s) at position ${rowIndex}`
      })
    },
  },

  insertColumn: {
    name: 'insertColumn',
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
    execute: async (args) => {
      const { columnLetter, count = 1 } = args
      return Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet()
        const endCol = String.fromCharCode(columnLetter.charCodeAt(0) + count - 1)
        const range = sheet.getRange(`${columnLetter}:${endCol}`)
        range.insert(Excel.InsertShiftDirection.right)
        await context.sync()
        return `Successfully inserted ${count} column(s) at position ${columnLetter}`
      })
    },
  },

  deleteRow: {
    name: 'deleteRow',
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
    execute: async (args) => {
      const { rowIndex, count = 1 } = args
      return Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet()
        const range = sheet.getRange(`${rowIndex}:${rowIndex + count - 1}`)
        range.delete(Excel.DeleteShiftDirection.up)
        await context.sync()
        return `Successfully deleted ${count} row(s) at position ${rowIndex}`
      })
    },
  },

  deleteColumn: {
    name: 'deleteColumn',
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
    execute: async (args) => {
      const { columnLetter, count = 1 } = args
      return Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet()
        const endCol = String.fromCharCode(columnLetter.charCodeAt(0) + count - 1)
        const range = sheet.getRange(`${columnLetter}:${endCol}`)
        range.delete(Excel.DeleteShiftDirection.left)
        await context.sync()
        return `Successfully deleted ${count} column(s) at position ${columnLetter}`
      })
    },
  },

  mergeCells: {
    name: 'mergeCells',
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
    execute: async (args) => {
      const { address, merge = true } = args
      return Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet()
        const range = address ? sheet.getRange(address) : context.workbook.getSelectedRange()

        if (merge) {
          range.merge()
        } else {
          range.unmerge()
        }
        await context.sync()
        return `Successfully ${merge ? 'merged' : 'unmerged'} cells${address ? ` at ${address}` : ''}`
      })
    },
  },

  setCellNumberFormat: {
    name: 'setCellNumberFormat',
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
    execute: async (args) => {
      const { address, format } = args
      return Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet()
        const range = address ? sheet.getRange(address) : context.workbook.getSelectedRange()
        range.numberFormat = range.values.map((row: any[]) => row.map(() => format))
        await context.sync()
        return `Successfully set number format "${format}"${address ? ` for ${address}` : ' for selection'}`
      })
    },
  },

  clearRange: {
    name: 'clearRange',
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
    execute: async (args) => {
      const { address, clearType = 'all' } = args
      return Excel.run(async (context) => {
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
      })
    },
  },

  getCellFormula: {
    name: 'getCellFormula',
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
    execute: async (args) => {
      const { address } = args
      return Excel.run(async (context) => {
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
      })
    },
  },

  searchAndReplace: {
    name: 'searchAndReplace',
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
    execute: async (args) => {
      const { searchText, replaceText } = args
      return Excel.run(async (context) => {
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
      })
    },
  },

  autoFitColumns: {
    name: 'autoFitColumns',
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
    execute: async (args) => {
      const { address } = args
      return Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet()
        const range = address ? sheet.getRange(address) : sheet.getUsedRange()
        range.format.autofitColumns()
        range.format.autofitRows()
        await context.sync()
        return `Successfully auto-fitted columns and rows${address ? ` for ${address}` : ''}`
      })
    },
  },

  addWorksheet: {
    name: 'addWorksheet',
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
    execute: async (args) => {
      const { name } = args
      return Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.add(name || undefined)
        sheet.activate()
        sheet.load('name')
        await context.sync()
        return `Successfully created and activated worksheet "${sheet.name}"`
      })
    },
  },

  setColumnWidth: {
    name: 'setColumnWidth',
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
    execute: async (args) => {
      const { columnLetter, width } = args
      return Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet()
        const colRange = columnLetter.includes(':')
          ? sheet.getRange(columnLetter)
          : sheet.getRange(`${columnLetter}:${columnLetter}`)
        colRange.format.columnWidth = width
        await context.sync()
        return `Successfully set column ${columnLetter} width to ${width}`
      })
    },
  },


  setRowHeight: {
    name: 'setRowHeight',
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
    execute: async (args) => {
      const { rowIndex, height } = args
      return Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet()
        const rowRange = sheet.getRange(`${rowIndex}:${rowIndex}`)
        rowRange.format.rowHeight = height
        await context.sync()
        return `Successfully set row ${rowIndex} height to ${height}`
      })
    },
  },


  applyConditionalFormatting: {
    name: 'applyConditionalFormatting',
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
    execute: async (args) => {
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

      return Excel.run(async (context) => {
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

        if (fillColor) cf.format.fill.color = fillColor
        if (fontColor) cf.format.font.color = fontColor
        if (bold !== undefined) cf.format.font.bold = bold
        if (stopIfTrue !== undefined) cf.stopIfTrue = stopIfTrue

        await context.sync()
        return `Successfully applied ${ruleType} conditional formatting on ${address}`
      })
    },
  },

  getConditionalFormattingRules: {
    name: 'getConditionalFormattingRules',
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
    execute: async (args) => {
      const { address } = args
      return Excel.run(async (context) => {
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
      })
    },
  },
}

export function getExcelToolDefinitions(): ExcelToolDefinition[] {
  return Object.values(excelToolDefinitions)
}

export function getExcelTool(name: ExcelToolName): ExcelToolDefinition | undefined {
  return excelToolDefinitions[name]
}

export { excelToolDefinitions }
