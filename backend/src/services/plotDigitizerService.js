/**
 * Chart data extraction service — powered by WebPlotDigitizer algorithms.
 *
 * Uses adapted core algorithms from WebPlotDigitizer (Copyright 2025 Ankit Rohatgi, AGPL-3.0)
 * for color-based pixel detection, blob averaging, and bar extraction.
 * Image loading via jimp (pure JS, no native binaries).
 *
 * See backend/src/services/wpd/ for the adapted WPD source files.
 */

import { Jimp } from 'jimp'
import { dist3d, cspline, csplineInterp } from './wpd/mathFunctions.js'
import { AveragingWindowCore } from './wpd/averagingWindowCore.js'
import { BarExtractionAlgo } from './wpd/barExtraction.js'
import { Dataset } from './wpd/dataset.js'

/**
 * Extract data points from a chart image by scanning pixels matching a target color.
 *
 * Pipeline:
 * 1. Load image via jimp → get RGBA pixel buffer.
 * 2. Color detection (WPD dist3d): build a Set<pixelIndex> of matching pixels.
 * 3. Curve extraction (WPD AveragingWindowCore or BarExtractionAlgo).
 * 4. Coordinate mapping: pixel → real values via axis ranges.
 * 5. Optional cubic spline smoothing (WPD cspline).
 *
 * @param {Object} options
 * @param {Buffer} options.imageBuffer - Raw image file buffer (PNG, JPEG, etc.)
 * @param {[number, number]} options.xAxisRange - [min, max] real values on the X axis
 * @param {[number, number]} options.yAxisRange - [min, max] real values on the Y axis
 * @param {string} options.targetColor - Hex color of the data series (e.g. "#FF0000")
 * @param {string} [options.chartType="line"] - "line", "scatter", "bar", "area"
 * @param {number} [options.colorTolerance=120] - Max Euclidean RGB distance (0–441). WPD default is 120.
 * @param {number} [options.numPoints=40] - Desired number of output points
 * @returns {Promise<Object>} { points: [{x, y}], pixelsMatched, imageSize, plotBounds }
 */
export async function extractChartData({
  imageBuffer,
  xAxisRange,
  yAxisRange,
  targetColor,
  chartType = 'line',
  colorTolerance = 120,
  numPoints = 40,
}) {
  // --- Parse target color ---
  const hex = targetColor.replace('#', '')
  let targetR, targetG, targetB
  if (hex.length === 3) {
    targetR = parseInt(hex[0] + hex[0], 16)
    targetG = parseInt(hex[1] + hex[1], 16)
    targetB = parseInt(hex[2] + hex[2], 16)
  } else {
    targetR = parseInt(hex.slice(0, 2), 16)
    targetG = parseInt(hex.slice(2, 4), 16)
    targetB = parseInt(hex.slice(4, 6), 16)
  }

  if ([targetR, targetG, targetB].some(v => isNaN(v))) {
    throw new Error(`Invalid targetColor "${targetColor}". Expected hex format #RRGGBB or #RGB.`)
  }

  // --- Load image with jimp ---
  const image = await Jimp.read(imageBuffer)
  const { width, height } = image
  const data = image.bitmap.data

  // --- Step 1: Color detection — build binary data (WPD approach) ---
  // Uses dist3d for Euclidean RGB distance, stores matching pixel indices in a Set.
  // This matches WPD's AutoDetectionData.generateBinaryDataUsingFullImage().
  const binaryData = new Set()
  for (let idx = 0; idx < width * height; idx++) {
    let ir = data[idx * 4]
    let ig = data[idx * 4 + 1]
    let ib = data[idx * 4 + 2]
    const ia = data[idx * 4 + 3]

    // WPD convention: for fully transparent pixels, assume white
    if (ia === 0) {
      ir = 255
      ig = 255
      ib = 255
    }

    const d = dist3d(ir, ig, ib, targetR, targetG, targetB)
    if (d <= colorTolerance) {
      binaryData.add(idx)
    }
  }

  if (binaryData.size === 0) {
    return {
      points: [],
      pixelsMatched: 0,
      imageSize: { width, height },
      warning: 'No pixels matched the target color within the tolerance. Try increasing colorTolerance or adjusting targetColor.',
    }
  }

  // --- Step 2: Curve extraction using WPD algorithms ---
  const dataSeries = new Dataset()
  const isBar = chartType === 'bar'

  if (isBar) {
    // WPD BarExtractionAlgo: finds top/bottom edges of bars, groups by proximity
    const delX = Math.max(10, Math.round(width / (numPoints * 2)))
    const delVal = Math.max(5, Math.round(height / 50))
    const barAlgo = new BarExtractionAlgo(binaryData, height, width, delX, delVal, dataSeries, 'vertical')
    barAlgo.run()
  } else {
    // WPD AveragingWindowCore: column-by-column blob detection + neighbor merging
    // xStep and yStep control the merging window (in pixels)
    const xStep = Math.max(3, Math.round(width / numPoints))
    const yStep = Math.max(3, Math.round(height / 50))
    const avgCore = new AveragingWindowCore(binaryData, height, width, xStep, yStep, dataSeries)
    avgCore.run()
  }

  if (dataSeries.getCount() === 0) {
    return {
      points: [],
      pixelsMatched: binaryData.size,
      imageSize: { width, height },
      warning: 'Pixels matched but no coherent data points could be extracted. Try adjusting colorTolerance.',
    }
  }

  // --- Step 3: Determine bounding box from detected pixels ---
  let pxMin = Infinity, pxMax = -Infinity
  let pyMin = Infinity, pyMax = -Infinity
  for (const idx of binaryData) {
    const px = idx % width
    const py = Math.floor(idx / width)
    if (px < pxMin) pxMin = px
    if (px > pxMax) pxMax = px
    if (py < pyMin) pyMin = py
    if (py > pyMax) pyMax = py
  }

  const pxSpan = pxMax - pxMin || 1
  const pySpan = pyMax - pyMin || 1

  // --- Step 4: Map pixel coordinates → real-world values ---
  const xMin = xAxisRange[0]
  const xMax = xAxisRange[1]
  const yMin = yAxisRange[0]
  const yMax = yAxisRange[1]

  const rawPixelPoints = dataSeries.getAllPixels()
  const realPoints = rawPixelPoints.map(({ x: px, y: py }) => ({
    x: xMin + ((px - pxMin) / pxSpan) * (xMax - xMin),
    y: yMax - ((py - pyMin) / pySpan) * (yMax - yMin), // Y inverted: top of image = yMax
  }))

  // Sort by X
  realPoints.sort((a, b) => a.x - b.x)

  // --- Step 5: Optional cubic spline smoothing (WPD cspline) ---
  let outputPoints = realPoints

  if (realPoints.length >= 3 && (chartType === 'line' || chartType === 'area') && realPoints.length > numPoints) {
    // Use WPD's cubic spline to interpolate at evenly spaced X values
    const xVals = realPoints.map(p => p.x)
    const yVals = realPoints.map(p => p.y)
    const cs = cspline(xVals, yVals)

    if (cs) {
      const interpPoints = []
      const step = (xMax - xMin) / (numPoints - 1)
      for (let i = 0; i < numPoints; i++) {
        const xi = xMin + i * step
        const yi = csplineInterp(cs, xi)
        if (yi !== null) {
          interpPoints.push({ x: round(xi, 3), y: round(yi, 3) })
        }
      }
      if (interpPoints.length > 0) {
        outputPoints = interpPoints
      }
    }
  }

  // Round the final output
  outputPoints = outputPoints.map(p => ({ x: round(p.x, 3), y: round(p.y, 3) }))

  return {
    points: outputPoints,
    pixelsMatched: binaryData.size,
    imageSize: { width, height },
    plotBounds: { pxMin, pxMax, pyMin, pyMax },
  }
}

function round(value, decimals) {
  const factor = 10 ** decimals
  return Math.round(value * factor) / factor
}
