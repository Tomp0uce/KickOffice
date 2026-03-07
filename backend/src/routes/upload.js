import { Router } from 'express'
import multer from 'multer'
import { fileTypeFromBuffer } from 'file-type'
import { PDFParse } from 'pdf-parse'
import mammoth from 'mammoth'
import * as xlsx from 'xlsx'
import { ErrorCodes } from '../config/errorCodes.js'
import { logAndRespond } from '../utils/http.js'
import { systemLog } from '../utils/logger.js'

const uploadRouter = Router()

const UPLOAD_MAX_FILE_SIZE = 10 * 1024 * 1024 // 10MB max file size
const UPLOAD_MAX_FIELD_SIZE = 1024 // 1KB per non-file field
const TEXT_MAX_CHARS = 600000 // approx 200k tokens — proportional to GPT-5.2 400k context window

// Configure multer to store files in memory
const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: UPLOAD_MAX_FILE_SIZE,
    fields: 10,
    fieldSize: UPLOAD_MAX_FIELD_SIZE,
  }
})

uploadRouter.post('/', upload.single('file'), async (req, res) => {
  if (!req.file) {
    return logAndRespond(res, 400, { code: ErrorCodes.NO_FILE_UPLOADED, error: 'No file uploaded' }, 'POST /api/upload')
  }

  const file = req.file
  const filename = file.originalname || 'unknown_file'
  let mimeType = file.mimetype
  
  try {
    const typeInfo = await fileTypeFromBuffer(file.buffer)
    if (typeInfo) {
      mimeType = typeInfo.mime
    }
  } catch (err) {
    systemLog('WARN', 'Could not determine file type from buffer', err)
  }
  
  systemLog('INFO', `POST /api/upload started parsing file: ${filename}`, {
    size: file.size,
    mimeType
  })

  try {
    let extractedText = ''

    // PDF Extraction
    if (mimeType === 'application/pdf' || filename.toLowerCase().endsWith('.pdf')) {
      try {
        const parser = new PDFParse({ data: file.buffer })
        const data = await parser.getText()
        await parser.destroy()
        if (!data || !data.text) throw new Error('Empty or unreadable PDF')
        extractedText = data.text
      } catch (pdfErr) {
        systemLog('ERROR', 'PDF extraction failed', pdfErr)
        return logAndRespond(res, 400, { code: ErrorCodes.PDF_EXTRACTION_FAILED, error: 'Failed to extract text from PDF. The file may be corrupted or encrypted.' }, 'POST /api/upload')
      }
    } 
    // DOCX Extraction
    else if (
      mimeType === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' ||
      filename.toLowerCase().endsWith('.docx')
    ) {
      try {
        const result = await mammoth.extractRawText({ buffer: file.buffer })
        extractedText = result.value
      } catch (docxErr) {
        systemLog('ERROR', 'DOCX extraction failed', docxErr)
        return logAndRespond(res, 400, { code: ErrorCodes.DOCX_EXTRACTION_FAILED, error: 'Failed to extract text from DOCX. The file may be corrupted.' }, 'POST /api/upload')
      }
    } 
    // XLSX / CSV Extraction
    else if (
      mimeType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
      mimeType === 'text/csv' ||
      filename.toLowerCase().endsWith('.xlsx') ||
      filename.toLowerCase().endsWith('.csv')
    ) {
      const workbook = xlsx.read(file.buffer, { type: 'buffer' })
      const allCsv = []
      
      for (const sheetName of workbook.SheetNames) {
        const sheet = workbook.Sheets[sheetName]
        if (!sheet) continue
        
        const csv = xlsx.utils.sheet_to_csv(sheet)
        if (csv.trim()) {
          allCsv.push(`--- Sheet: ${sheetName} ---\n${csv}`)
        }
      }
      extractedText = allCsv.join('\n\n')
    }
    // Plain text fallback
    else if (mimeType.startsWith('text/') || filename.toLowerCase().endsWith('.txt') || filename.toLowerCase().endsWith('.md')) {
       extractedText = file.buffer.toString('utf-8')
    }
    // Image types — return base64 data-URI for LLM vision
    else if (
      mimeType === 'image/png' ||
      mimeType === 'image/jpeg' ||
      mimeType === 'image/jpg' ||
      mimeType === 'image/webp' ||
      mimeType === 'image/gif' ||
      /\.(png|jpe?g|webp|gif)$/i.test(filename)
    ) {
      const b64 = file.buffer.toString('base64')
      const imageBase64 = `data:${mimeType};base64,${b64}`
      systemLog('INFO', `POST /api/upload completed image encoding`, { filename, bytes: file.size })
      return res.json({ filename, imageBase64 })
    }
    else {
      return logAndRespond(res, 400, {
        code: ErrorCodes.UNSUPPORTED_FILE_TYPE,
        error: `Unsupported file type: ${mimeType}. Please upload a PDF, DOCX, XLSX, CSV, Image (PNG/JPG/WEBP/GIF), or Text file.`
      }, 'POST /api/upload')
    }

    if (!extractedText || !extractedText.trim()) {
      return logAndRespond(res, 400, { code: ErrorCodes.FILE_EMPTY, error: 'No text could be extracted from the file. Make sure the file contains readable text.' }, 'POST /api/upload')
    }

    // Limit text size to prevent enormous context windows (approx context token limit defense)
    if (extractedText.length > TEXT_MAX_CHARS) {
        extractedText = extractedText.substring(0, TEXT_MAX_CHARS) + '\n\n... [Content truncated due to file size]'
    }

    res.json({
      filename,
      extractedText: extractedText.trim()
    })

    systemLog('INFO', `POST /api/upload completed file parsing`, { filename, charCount: extractedText.length })

  } catch (error) {
    systemLog('ERROR', `POST /api/upload error parsing file ${filename}`, error)
    return logAndRespond(res, 500, { code: ErrorCodes.INTERNAL_ERROR, error: 'Failed to process file' }, 'POST /api/upload')
  }
})

export { uploadRouter }
