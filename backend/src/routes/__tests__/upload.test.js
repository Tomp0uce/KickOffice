import { describe, it, expect, vi, beforeEach } from 'vitest'
import express from 'express'
import request from 'supertest'

// ─── Mocks ────────────────────────────────────────────────────────────────────
vi.mock('file-type', () => ({
  fileTypeFromBuffer: vi.fn().mockResolvedValue(null),
}))

vi.mock('pdf-parse', () => ({
  PDFParse: vi.fn(),
}))

vi.mock('mammoth', () => ({
  default: { extractRawText: vi.fn() },
}))

vi.mock('../../services/imageStore.js', () => ({
  storeImage: vi.fn(),
}))

vi.mock('../../utils/logger.js', () => ({
  default: {
    child: () => ({
      info: vi.fn(),
      warn: vi.fn(),
      error: vi.fn(),
    }),
  },
}))

// ExcelJS is NOT mocked — we use the real library to verify the migration works.

import { uploadRouter } from '../upload.js'

// ─── Test App Setup ───────────────────────────────────────────────────────────
function createApp() {
  const app = express()
  // Fake logger middleware
  app.use((req, _res, next) => {
    req.logger = { info: vi.fn(), warn: vi.fn(), error: vi.fn() }
    next()
  })
  app.use('/api/upload', uploadRouter)
  return app
}

// ─── Helper: create a real XLSX buffer using ExcelJS ──────────────────────────
async function createXlsxBuffer(sheets) {
  const ExcelJS = (await import('exceljs')).default
  const wb = new ExcelJS.Workbook()
  for (const [name, rows] of Object.entries(sheets)) {
    const ws = wb.addWorksheet(name)
    for (const row of rows) {
      ws.addRow(row)
    }
  }
  return Buffer.from(await wb.xlsx.writeBuffer())
}

// ─────────────────────────────────────────────────────────────────────────────
describe('POST /api/upload — XLSX parsing (SEC-H2: exceljs migration)', () => {
  let app

  beforeEach(() => {
    vi.clearAllMocks()
    app = createApp()
  })

  it('extracts CSV text from a single-sheet XLSX file', async () => {
    const xlsxBuffer = await createXlsxBuffer({
      'Sales': [
        ['Product', 'Price', 'Qty'],
        ['Widget', 10, 5],
        ['Gadget', 20, 3],
      ],
    })

    const res = await request(app)
      .post('/api/upload')
      .attach('file', xlsxBuffer, 'test.xlsx')

    expect(res.status).toBe(200)
    expect(res.body.extractedText).toContain('--- Sheet: Sales ---')
    expect(res.body.extractedText).toContain('Product')
    expect(res.body.extractedText).toContain('Widget')
    expect(res.body.extractedText).toContain('10')
  })

  it('extracts CSV from multi-sheet XLSX files', async () => {
    const xlsxBuffer = await createXlsxBuffer({
      'Q1': [['Jan', 100], ['Feb', 200]],
      'Q2': [['Mar', 300], ['Apr', 400]],
    })

    const res = await request(app)
      .post('/api/upload')
      .attach('file', xlsxBuffer, 'multi.xlsx')

    expect(res.status).toBe(200)
    expect(res.body.extractedText).toContain('--- Sheet: Q1 ---')
    expect(res.body.extractedText).toContain('--- Sheet: Q2 ---')
    expect(res.body.extractedText).toContain('100')
    expect(res.body.extractedText).toContain('400')
  })

  it('handles cells with commas by quoting them', async () => {
    const xlsxBuffer = await createXlsxBuffer({
      'Data': [['Name', 'Description'], ['Item', 'red, blue, green']],
    })

    const res = await request(app)
      .post('/api/upload')
      .attach('file', xlsxBuffer, 'commas.xlsx')

    expect(res.status).toBe(200)
    // Value with commas should be quoted
    expect(res.body.extractedText).toContain('"red, blue, green"')
  })

  it('handles CSV files as plain text', async () => {
    const csvContent = 'Name,Age\nAlice,30\nBob,25'
    const csvBuffer = Buffer.from(csvContent)

    const res = await request(app)
      .post('/api/upload')
      .attach('file', csvBuffer, 'data.csv')

    expect(res.status).toBe(200)
    // CSV now handled as plain text — no spreadsheet library needed
    expect(res.body.extractedText).toContain('Alice')
    expect(res.body.extractedText).toContain('30')
  })
})
