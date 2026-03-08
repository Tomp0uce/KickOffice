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

export function dist3d(x1, y1, z1, x2, y2, z2) {
  return Math.sqrt((x1 - x2) * (x1 - x2) + (y1 - y2) * (y1 - y2) + (z1 - z2) * (z1 - z2))
}

export function cspline(x, y) {
  const len = x.length
  if (len < 3) return null

  const cs = { x, y, len, d: [] }
  const l = []
  const b = []

  b[0] = 2.0
  l[0] = 3.0 * (y[1] - y[0])
  for (let i = 1; i < len - 1; ++i) {
    b[i] = 4.0 - 1.0 / b[i - 1]
    l[i] = 3.0 * (y[i + 1] - y[i - 1]) - l[i - 1] / b[i - 1]
  }

  b[len - 1] = 2.0 - 1.0 / b[len - 2]
  l[len - 1] = 3.0 * (y[len - 1] - y[len - 2]) - l[len - 2] / b[len - 1]

  let i = len - 1
  cs.d[i] = l[i] / b[i]
  while (i > 0) {
    --i
    cs.d[i] = (l[i] - cs.d[i + 1]) / b[i]
  }

  return cs
}

export function csplineInterp(cs, x) {
  if (x >= cs.x[cs.len - 1] || x < cs.x[0]) return null

  let i = 0
  while (x > cs.x[i]) i++
  i = (i > 0) ? i - 1 : 0

  const t = (x - cs.x[i]) / (cs.x[i + 1] - cs.x[i])
  const a = cs.y[i]
  const b = cs.d[i]
  const c = 3.0 * (cs.y[i + 1] - cs.y[i]) - 2.0 * cs.d[i] - cs.d[i + 1]
  const d = 2.0 * (cs.y[i] - cs.y[i + 1]) + cs.d[i] + cs.d[i + 1]
  return a + b * t + c * t * t + d * t * t * t
}
