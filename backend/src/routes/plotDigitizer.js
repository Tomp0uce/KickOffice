import { Router } from 'express'
import { getImage } from '../services/imageStore.js'
import { extractChartData } from '../services/plotDigitizerService.js'
import { ErrorCodes } from '../config/errorCodes.js'
import { logAndRespond } from '../utils/http.js'

const plotDigitizerRouter = Router()

plotDigitizerRouter.post('/', async (req, res) => {
  const {
    imageId,
    xAxisRange,
    yAxisRange,
    targetColor,
    chartType = 'line',
    colorTolerance = 120,
    numPoints = 40,
  } = req.body

  // --- Input validation ---
  if (!imageId || typeof imageId !== 'string') {
    return logAndRespond(res, 400, {
      code: ErrorCodes.VALIDATION_ERROR,
      error: 'imageId is required and must be a string.',
    }, 'POST /api/chart-extract')
  }

  if (!Array.isArray(xAxisRange) || xAxisRange.length !== 2 ||
      !xAxisRange.every(v => typeof v === 'number' && isFinite(v))) {
    return logAndRespond(res, 400, {
      code: ErrorCodes.VALIDATION_ERROR,
      error: 'xAxisRange must be an array of 2 finite numbers [min, max].',
    }, 'POST /api/chart-extract')
  }

  if (!Array.isArray(yAxisRange) || yAxisRange.length !== 2 ||
      !yAxisRange.every(v => typeof v === 'number' && isFinite(v))) {
    return logAndRespond(res, 400, {
      code: ErrorCodes.VALIDATION_ERROR,
      error: 'yAxisRange must be an array of 2 finite numbers [min, max].',
    }, 'POST /api/chart-extract')
  }

  if (!targetColor || typeof targetColor !== 'string' || !/^#([0-9a-fA-F]{3}|[0-9a-fA-F]{6})$/.test(targetColor)) {
    return logAndRespond(res, 400, {
      code: ErrorCodes.VALIDATION_ERROR,
      error: 'targetColor must be a hex color string (#RGB or #RRGGBB).',
    }, 'POST /api/chart-extract')
  }

  if (typeof colorTolerance !== 'number' || colorTolerance < 0 || colorTolerance > 441) {
    return logAndRespond(res, 400, {
      code: ErrorCodes.VALIDATION_ERROR,
      error: 'colorTolerance must be a number between 0 and 441.',
    }, 'POST /api/chart-extract')
  }

  if (typeof numPoints !== 'number' || numPoints < 5 || numPoints > 200) {
    return logAndRespond(res, 400, {
      code: ErrorCodes.VALIDATION_ERROR,
      error: 'numPoints must be a number between 5 and 200.',
    }, 'POST /api/chart-extract')
  }

  // --- Retrieve image from store ---
  const imageEntry = getImage(imageId)
  if (!imageEntry) {
    return logAndRespond(res, 404, {
      code: ErrorCodes.CHART_IMAGE_NOT_FOUND,
      error: 'Image not found or expired. Please re-upload the image.',
    }, 'POST /api/chart-extract')
  }

  req.logger.info('POST /api/chart-extract started', {
    imageId,
    targetColor,
    chartType,
    xAxisRange,
    yAxisRange,
    colorTolerance,
    numPoints,
  })

  try {
    const result = await extractChartData({
      imageBuffer: imageEntry.buffer,
      xAxisRange,
      yAxisRange,
      targetColor,
      chartType,
      colorTolerance,
      numPoints,
    })

    req.logger.info('POST /api/chart-extract completed', {
      imageId,
      pointsExtracted: result.points.length,
      pixelsMatched: result.pixelsMatched,
    })

    return res.json(result)
  } catch (error) {
    req.logger.error('POST /api/chart-extract error', { error, imageId })
    return logAndRespond(res, 500, {
      code: ErrorCodes.CHART_EXTRACTION_FAILED,
      error: `Chart data extraction failed: ${error.message}`,
    }, 'POST /api/chart-extract')
  }
})

export { plotDigitizerRouter }
