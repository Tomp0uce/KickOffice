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

class BarValue {
  constructor() {
    this.npoints = 0
    this.avgValTop = 0
    this.avgValBot = 0
    this.avgX = 0
  }

  append(x, valTop, valBot) {
    this.avgX = (this.npoints * this.avgX + x) / (this.npoints + 1.0)
    this.avgValTop = (this.npoints * this.avgValTop + valTop) / (this.npoints + 1.0)
    this.avgValBot = (this.npoints * this.avgValBot + valBot) / (this.npoints + 1.0)
    this.npoints++
  }

  isPointInGroup(x, valTop, valBot, delX, delVal) {
    if (this.npoints === 0) return true
    return (
      Math.abs(this.avgX - x) <= delX &&
      Math.abs(this.avgValTop - valTop) <= delVal &&
      Math.abs(this.avgValBot - valBot) <= delVal
    )
  }
}

/**
 * Bar chart extraction algorithm.
 * Finds top and bottom edges of bars, groups by proximity.
 *
 * @param {Set} binaryData - Set of flat pixel indices that matched the target color.
 * @param {number} imageHeight
 * @param {number} imageWidth
 * @param {number} delX - Grouping distance along the bar width axis (pixels).
 * @param {number} delVal - Grouping distance along the bar value axis (pixels).
 * @param {Dataset} dataSeries - Output container.
 * @param {'vertical'|'horizontal'} orientation - 'vertical' = bars grow up/down (Y axis), 'horizontal' = bars grow left/right (X axis).
 */
export class BarExtractionAlgo {
  constructor(binaryData, imageHeight, imageWidth, delX, delVal, dataSeries, orientation = 'vertical') {
    this._binaryData = binaryData
    this._imageHeight = imageHeight
    this._imageWidth = imageWidth
    this._delX = delX
    this._delVal = delVal
    this._dataSeries = dataSeries
    this._orientation = orientation
  }

  run() {
    const width = this._imageWidth
    const height = this._imageHeight
    const barValueColl = []
    const dataSeries = this._dataSeries

    dataSeries.clearAll()

    const appendData = (x, valTop, valBot) => {
      for (const bv of barValueColl) {
        if (bv.isPointInGroup(x, valTop, valBot, this._delX, this._delVal)) {
          bv.append(x, valTop, valBot)
          return
        }
      }
      const bv = new BarValue()
      bv.append(x, valTop, valBot)
      barValueColl.push(bv)
    }

    if (this._orientation === 'vertical') {
      // Vertical bars: scan each column (px), find top and bottom matching pixels
      for (let px = 0; px < width; px++) {
        let valTop = 0
        let valBot = height - 1
        let valCount = 0

        for (let py = 0; py < height; py++) {
          if (this._binaryData.has(py * width + px)) {
            valTop = py
            valCount++
            break
          }
        }
        for (let py = height - 1; py >= 0; py--) {
          if (this._binaryData.has(py * width + px)) {
            valBot = py
            valCount++
            break
          }
        }
        if (valCount === 2) {
          appendData(px, valTop, valBot)
        }
      }
    } else {
      // Horizontal bars: scan each row (py), find left and right matching pixels
      for (let py = 0; py < height; py++) {
        let valTop = width - 1
        let valBot = 0
        let valCount = 0

        for (let px = width - 1; px >= 0; px--) {
          if (this._binaryData.has(py * width + px)) {
            valTop = px
            valCount++
            break
          }
        }
        for (let px = 0; px < width; px++) {
          if (this._binaryData.has(py * width + px)) {
            valBot = px
            valCount++
            break
          }
        }
        if (valCount === 2) {
          appendData(py, valTop, valBot)
        }
      }
    }

    // Output the averaged center of each bar group
    for (const bv of barValueColl) {
      if (this._orientation === 'vertical') {
        // X = bar center, Y = top edge (the meaningful value)
        dataSeries.addPixel(bv.avgX + 0.5, bv.avgValTop + 0.5)
      } else {
        // Y = bar center, X = right edge (the meaningful value)
        dataSeries.addPixel(bv.avgValTop + 0.5, bv.avgX + 0.5)
      }
    }

    return dataSeries
  }
}
