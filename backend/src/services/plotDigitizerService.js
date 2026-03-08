/**
 * Chart data extraction service — pure JS implementation using Jimp.
 *
 * Algorithm:
 * 1. Load image with Jimp → RGBA pixel buffer.
 * 2. Color detection: scan pixels inside plotAreaBox using Euclidean RGB distance.
 * 3. Bucket matching pixels into numPoints columns along the X axis.
 * 4. Average the Y positions within each bucket.
 * 5. Map (bucketCenter_px, avgY_px) → real (x, y) values via plotAreaBox and axis ranges.
 *    Note: pixel Y axis is inverted (0 = top of image).
 */

import { Jimp } from 'jimp'

/**
 * Resolve plotAreaBox coordinates to pixel values.
 * If all four values are in [0, 1], they are treated as fractions of image dimensions.
 * Otherwise, they are treated as raw pixel values and clamped to image bounds.
 *
 * @param {{ xMinPx: number, xMaxPx: number, yMinPx: number, yMaxPx: number }} box
 * @param {number} width  - Image width in pixels
 * @param {number} height - Image height in pixels
 * @returns {{ pxMin: number, pxMax: number, pyMin: number, pyMax: number }}
 */
function resolvePlotBox(box, width, height) {
  const { xMinPx, xMaxPx, yMinPx, yMaxPx } = box

  const isFraction =
    xMinPx >= 0 && xMinPx <= 1 &&
    xMaxPx >= 0 && xMaxPx <= 1 &&
    yMinPx >= 0 && yMinPx <= 1 &&
    yMaxPx >= 0 && yMaxPx <= 1

  let pxMin, pxMax, pyMin, pyMax
  if (isFraction) {
    pxMin = Math.round(xMinPx * width)
    pxMax = Math.round(xMaxPx * width)
    pyMin = Math.round(yMinPx * height)
    pyMax = Math.round(yMaxPx * height)
  } else {
    pxMin = Math.round(xMinPx)
    pxMax = Math.round(xMaxPx)
    pyMin = Math.round(yMinPx)
    pyMax = Math.round(yMaxPx)
  }

  // Clamp to valid image bounds
  pxMin = Math.max(0, Math.min(pxMin, width - 1))
  pxMax = Math.max(pxMin + 1, Math.min(pxMax, width - 1))
  pyMin = Math.max(0, Math.min(pyMin, height - 1))
  pyMax = Math.max(pyMin + 1, Math.min(pyMax, height - 1))

  return { pxMin, pxMax, pyMin, pyMax }
}

/**
 * Extract data points from a chart image by scanning pixels matching a target color
 * within the defined plot area bounding box.
 *
 * @param {Object} options
 * @param {Buffer} options.imageBuffer    - Raw image file buffer (PNG, JPEG, etc.)
 * @param {[number, number]} options.xAxisRange - [min, max] real values on the X axis
 * @param {[number, number]} options.yAxisRange - [min, max] real values on the Y axis
 * @param {string}  options.targetColor   - Hex color of the data series (e.g. "#FF0000")
 * @param {{ xMinPx: number, xMaxPx: number, yMinPx: number, yMaxPx: number }} options.plotAreaBox
 *   Bounding box of the chart's plot area. Values in [0,1] are treated as fractions;
 *   values > 1 are treated as raw pixel coordinates.
 * @param {string}  [options.chartType="line"] - "line", "scatter", "bar", or "area"
 * @param {number}  [options.colorTolerance=120] - Max Euclidean RGB distance (0–441)
 * @param {number}  [options.numPoints=40]  - Number of X-axis buckets
 * @returns {Promise<Object>} { points: [{x, y}], pixelsMatched, imageSize, plotBounds }
 */
export async function extractChartData({
  imageBuffer,
  xAxisRange,
  yAxisRange,
  targetColor,
  plotAreaBox,
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

  // --- Load image ---
  const image = await Jimp.read(imageBuffer)
  const { width, height } = image
  const data = image.bitmap.data

  // --- Resolve plot area bounding box to pixel coordinates ---
  const { pxMin, pxMax, pyMin, pyMax } = resolvePlotBox(plotAreaBox, width, height)
  const plotWidth = pxMax - pxMin
  const plotHeight = pyMax - pyMin

  // --- Step 1: Scan pixels inside the plot area for color match ---
  const matchingPixels = [] // { px, py }
  for (let py = pyMin; py <= pyMax; py++) {
    for (let px = pxMin; px <= pxMax; px++) {
      const i = (py * width + px) * 4
      let r = data[i]
      let g = data[i + 1]
      let b = data[i + 2]
      const a = data[i + 3]

      // Fully transparent pixels are treated as white (no match expected)
      if (a === 0) { r = 255; g = 255; b = 255 }

      const d = Math.sqrt(
        (r - targetR) ** 2 +
        (g - targetG) ** 2 +
        (b - targetB) ** 2
      )
      if (d <= colorTolerance) {
        matchingPixels.push({ px, py })
      }
    }
  }

  if (matchingPixels.length === 0) {
    return {
      points: [],
      pixelsMatched: 0,
      imageSize: { width, height },
      warning: 'No pixels matched the target color within the tolerance. Try increasing colorTolerance or adjusting targetColor.',
    }
  }

  // --- Step 2: Bucket matching pixels into numPoints columns along X ---
  const buckets = Array.from({ length: numPoints }, () => /** @type {number[]} */ ([]))
  for (const { px, py } of matchingPixels) {
    const relX = (px - pxMin) / plotWidth
    const bucketIdx = Math.min(numPoints - 1, Math.floor(relX * numPoints))
    buckets[bucketIdx].push(py)
  }

  // --- Step 3: Map each non-empty bucket to real (x, y) coordinates ---
  const [xMin, xMax] = xAxisRange
  const [yMin, yMax] = yAxisRange

  const points = []
  for (let i = 0; i < numPoints; i++) {
    if (buckets[i].length === 0) continue

    // Average pixel Y for this bucket
    const avgPy = buckets[i].reduce((sum, py) => sum + py, 0) / buckets[i].length

    // Map bucket center → real X (linear interpolation within plot area)
    const relX = (i + 0.5) / numPoints
    const realX = xMin + relX * (xMax - xMin)

    // Map average pixel Y → real Y
    // Pixel Y=pyMin is the TOP of the plot area → yMax
    // Pixel Y=pyMax is the BOTTOM of the plot area → yMin
    const relY = (avgPy - pyMin) / plotHeight
    const realY = yMax - relY * (yMax - yMin)

    points.push({ x: round(realX, 3), y: round(realY, 3) })
  }

  // Sort by X
  points.sort((a, b) => a.x - b.x)

  return {
    points,
    pixelsMatched: matchingPixels.length,
    imageSize: { width, height },
    plotBounds: { pxMin, pxMax, pyMin, pyMax },
  }
}

function round(value, decimals) {
  const factor = 10 ** decimals
  return Math.round(value * factor) / factor
}
