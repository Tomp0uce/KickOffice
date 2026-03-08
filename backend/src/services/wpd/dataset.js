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

    Adapted for KickOffice (ESM) — minimal version of wpd.Dataset.
*/

/**
 * Minimal Dataset class compatible with WPD curve detection algorithms.
 * Stores pixel-space data points as {x, y} pairs.
 */
export class Dataset {
  constructor() {
    this._dataPoints = []
    this._pixelMetadataKeys = []
    this.dataPointsHaveLabels = false
  }

  addPixel(pxi, pyi, mdata) {
    const idx = this._dataPoints.length
    this._dataPoints.push({ x: pxi, y: pyi, metadata: mdata || null })
    return idx
  }

  getPixel(index) {
    return this._dataPoints[index]
  }

  getAllPixels() {
    return this._dataPoints
  }

  getCount() {
    return this._dataPoints.length
  }

  clearAll() {
    this._dataPoints = []
  }

  setMetadataKeys(keys) {
    this._pixelMetadataKeys = keys
  }

  getMetadataKeys() {
    return this._pixelMetadataKeys
  }
}
