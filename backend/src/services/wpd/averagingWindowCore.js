/*
    WebPlotDigitizer - web based chart data extraction software (and more)

    Copyright (C) 2025 Ankit Rohatgi

    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU Affero General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU Affero General Public License for more details.

    You should have received a copy of the GNU Affero General Public License
    along with this program.  If not, see <https://www.gnu.org/licenses/>

    Adapted for KickOffice (ESM) from WebPlotDigitizer source.
*/

/**
 * Core averaging window algorithm for curve/line data extraction.
 *
 * For each column of pixels:
 *   1. Scan vertically to find "blobs" (groups of matching pixels separated by yStep gap).
 *   2. Average the Y position of each blob.
 * Then merge nearby points using xStep/yStep proximity window.
 *
 * @param {Set} binaryData - Set of flat pixel indices (y * width + x) that matched the target color.
 * @param {number} imageHeight
 * @param {number} imageWidth
 * @param {number} xStep - Horizontal merge distance in pixels.
 * @param {number} yStep - Vertical merge distance in pixels.
 * @param {Dataset} dataSeries - Output container (addPixel(x, y) interface).
 */
export class AveragingWindowCore {
  constructor(binaryData, imageHeight, imageWidth, xStep, yStep, dataSeries) {
    this._binaryData = binaryData
    this._imageHeight = imageHeight
    this._imageWidth = imageWidth
    this._dx = xStep
    this._dy = yStep
    this._dataSeries = dataSeries
  }

  run() {
    const dw = this._imageWidth
    const dh = this._imageHeight
    const xStep = this._dx
    const yStep = this._dy
    const blobAvg = []
    let xPoints = []
    let xPointsPicked = 0
    let pointsPicked = 0

    this._dataSeries.clearAll()

    // Pass 1: Scan each column for vertical blobs
    for (let coli = 0; coli < dw; coli++) {
      let blobs = -1
      let firstbloby = -2.0 * yStep
      let bi = 0

      for (let rowi = 0; rowi < dh; rowi++) {
        if (this._binaryData.has(rowi * dw + coli)) {
          if (rowi > firstbloby + yStep) {
            blobs++
            bi = 1
            blobAvg[blobs] = rowi
            firstbloby = rowi
          } else {
            bi++
            blobAvg[blobs] = (blobAvg[blobs] * (bi - 1.0) + rowi) / bi
          }
        }
      }

      if (blobs >= 0) {
        const xi = coli + 0.5
        for (let blbi = 0; blbi <= blobs; blbi++) {
          const yi = blobAvg[blbi] + 0.5
          xPoints[xPointsPicked] = [xi, yi, true]
          xPointsPicked++
        }
      }
    }

    if (xPointsPicked === 0) return this._dataSeries

    // Pass 2: Merge nearby points within (xStep, yStep) window
    for (let pi = 0; pi < xPointsPicked; pi++) {
      if (xPoints[pi][2] === true) {
        let xxi = pi + 1
        let oldX = xPoints[pi][0]
        let oldY = xPoints[pi][1]
        let avgX = oldX
        let avgY = oldY
        let matches = 1
        let inRange = true

        while (inRange && xxi < xPointsPicked) {
          const newX = xPoints[xxi][0]
          const newY = xPoints[xxi][1]

          if (Math.abs(newX - oldX) <= xStep && Math.abs(newY - oldY) <= yStep && xPoints[xxi][2] === true) {
            avgX = (avgX * matches + newX) / (matches + 1.0)
            avgY = (avgY * matches + newY) / (matches + 1.0)
            matches++
            xPoints[xxi][2] = false
          }

          if (newX > oldX + 2 * xStep) {
            inRange = false
          }

          xxi++
        }

        xPoints[pi][2] = false
        pointsPicked++
        this._dataSeries.addPixel(avgX, avgY)
      }
    }

    xPoints = []
    return this._dataSeries
  }
}
