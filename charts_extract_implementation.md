# Chart Image Data Extraction — Implementation Guide

**Date**: 2026-03-08
**Feature**: Extract data from chart images and reproduce them in Excel
**Status**: Implemented

---

## Architecture Overview

```
User uploads chart image
        │
        ▼
┌─────────────────┐     ┌──────────────────────┐
│  Upload Route    │────▶│  Image Store (RAM)    │
│  POST /api/upload│     │  30-min TTL, UUID key │
└─────────────────┘     └──────────┬───────────┘
        │                          │
        ▼                          │
┌─────────────────┐                │
│  LLM (Vision)   │                │
│  Analyzes image:│                │
│  - axis ranges  │                │
│  - chart type   │                │
│  - data color   │                │
└────────┬────────┘                │
         │ calls extract_chart_data│
         ▼                         ▼
┌─────────────────┐     ┌──────────────────────┐
│  Chart Extract   │────▶│  PlotDigitizer       │
│  POST /api/      │     │  Service             │
│  chart-extract   │     │  (jimp pixel scan)   │
└─────────────────┘     └──────────┬───────────┘
                                   │
                                   ▼
                        ┌──────────────────────┐
                        │  JSON [{x, y}, ...]  │
                        │  ~40 sampled points   │
                        └──────────┬───────────┘
                                   │
                                   ▼
                        ┌──────────────────────┐
                        │  LLM writes to Excel │
                        │  setCellRange +      │
                        │  manageObject (chart) │
                        └──────────────────────┘
```

## Data Flow (Step by Step)

### 1. Image Upload
- User drags a chart image into KickOffice
- `POST /api/upload` receives the file via multer (memory storage)
- Backend detects image MIME type, stores buffer in **ImageStore** (in-memory Map with UUID key, 30-min TTL)
- Returns `{ filename, imageBase64, imageId }` to frontend

### 2. Vision Analysis (LLM)
- The image is injected into the LLM conversation as:
  - **Vision content**: `image_url` part in the message (for visual understanding)
  - **Text context**: `<uploaded_images>` block with `imageId` reference
- The LLM uses its vision capability to determine:
  - **X axis range**: e.g., `[2020, 2025]` (from axis labels)
  - **Y axis range**: e.g., `[0, 50000]` (from axis labels)
  - **Target color**: e.g., `"#0070C0"` (observed data series color)
  - **Chart type**: e.g., `"line"` (visual classification)

### 3. Pixel Extraction (Backend)
- LLM calls `extract_chart_data` tool with semantic parameters + `imageId` + `plotAreaBox`
- Frontend sends `POST /api/chart-extract` with JSON payload
- Backend retrieves image buffer from ImageStore by `imageId`
- **PlotDigitizerService** runs the algorithm:
  1. Loads image with **jimp** (pure JS, no native deps)
  2. Resolves `plotAreaBox` (fraction [0,1] or raw pixels) to pixel coordinates
  3. Scans all pixels **inside the plot area bounding box**, computes Euclidean RGB distance to target color
  4. Keeps pixels with `d <= colorTolerance` (default: 120)
  5. Buckets matching pixels by X column into `numPoints` buckets
  6. Averages Y pixel positions per bucket, maps to real (x, y) via axis ranges
- Returns `{ points: [{x, y}], pixelsMatched, imageSize, plotBounds }`

### 4. Excel Chart Creation (LLM)
- LLM receives the JSON points
- Calls `setCellRange` to write X/Y columns into the spreadsheet
- Calls `manageObject` to create the appropriate chart type

**Zero user interaction** — the entire flow is automated once the user uploads the image.

---

## Files Created

| File | Purpose |
|---|---|
| `backend/src/services/imageStore.js` | In-memory image buffer storage with UUID keys and 30-min TTL |
| `backend/src/services/plotDigitizerService.js` | Core pixel scanning algorithm — pure JS bucket-based implementation |
| `backend/src/routes/plotDigitizer.js` | Express route `POST /api/chart-extract` with input validation |

## Files Modified

| File | Change |
|---|---|
| `backend/src/routes/upload.js` | Store image buffer in ImageStore, return `imageId` in response |
| `backend/src/server.js` | Register `/api/chart-extract` route |
| `backend/src/config/errorCodes.js` | Add `CHART_IMAGE_NOT_FOUND`, `CHART_EXTRACTION_FAILED` |
| `backend/package.json` | Add `jimp` dependency (pure JS image processing) |
| `frontend/src/api/backend.ts` | Add `extractChartData()` API client + types + error codes |
| `frontend/src/utils/excelTools.ts` | Add `extract_chart_data` tool definition |
| `frontend/src/skills/excel.skill.md` | Document chart extraction workflow |
| `frontend/src/composables/useAgentPrompts.ts` | Add tool + workflow to Excel agent prompt |
| `frontend/src/composables/useAgentLoop.ts` | Inject `imageId` into `<uploaded_images>` context |

---

## Algorithm Details (PlotDigitizerService)

### Pure-JS Bucket Algorithm

The extraction engine is a fully original implementation using only **jimp** (pure JavaScript image processing, no native dependencies).

**Pipeline:**

1. **Plot area resolution** (`resolvePlotBox`):
   - Accepts `plotAreaBox` with `{xMinPx, xMaxPx, yMinPx, yMaxPx}`
   - Values in `[0, 1]` are treated as fractions of image dimensions (auto-detection)
   - Values `> 1` are treated as raw pixel coordinates
   - Clamped to valid image bounds

2. **Color detection**:
   - Euclidean RGB distance: `d = sqrt((R1-R2)^2 + (G1-G2)^2 + (B1-B2)^2)`
   - Scans only pixels **inside the plot area bounding box** (not the full image)
   - Pixels with `d <= colorTolerance` (default: 120) are collected
   - Transparent pixels (alpha=0) are treated as white (no match)

3. **X-axis bucketing**:
   - Matching pixels are assigned to one of `numPoints` equal-width columns based on their relative X position within the plot area
   - Each bucket accumulates pixel Y values

4. **Coordinate mapping**:
   ```
   relX = (i + 0.5) / numPoints                          // bucket center
   realX = xMin + relX * (xMax - xMin)

   avgPy = average(bucket[i].yValues)
   relY = (avgPy - pyMin) / (pyMax - pyMin)
   realY = yMax - relY * (yMax - yMin)                   // Y axis inverted
   ```

5. **Output**: `points` sorted by `x`, rounded to 3 decimal places.

### Key Design Decisions

- **LLM provides `plotAreaBox`**: The LLM uses vision to identify the exact plot area (axes origin, bounds), eliminating auto-detection errors from chart decorations (titles, legends, gridlines)
- **Pure JavaScript**: Uses `jimp` (no native binaries, no canvas/opencv) — safe for NAS and containerized deployment
- **LLM-driven semantics + deterministic math**: Vision AI for axis ranges, colors, and bounding box; backend for pixel-level coordinate mapping
- **License-clean**: Entirely original implementation with no dependency on external chart digitization libraries
- **Graceful degradation**: Returns warning with 0 points if no pixels match

---

## API Reference

### POST /api/chart-extract

**Headers**: `X-User-Key`, `X-User-Email` (standard auth)

**Request Body**:
```json
{
  "imageId": "uuid-from-upload",
  "xAxisRange": [0, 100],
  "yAxisRange": [0, 50],
  "targetColor": "#FF0000",
  "plotAreaBox": { "xMinPx": 0.08, "xMaxPx": 0.95, "yMinPx": 0.05, "yMaxPx": 0.88 },
  "chartType": "line",
  "colorTolerance": 120,
  "numPoints": 40
}
```

| Parameter | Type | Required | Default | Description |
|---|---|---|---|---|
| `imageId` | string | Yes | — | UUID from upload response |
| `xAxisRange` | [number, number] | Yes | — | Real X axis [min, max] |
| `yAxisRange` | [number, number] | Yes | — | Real Y axis [min, max] |
| `targetColor` | string | Yes | — | Hex color (#RGB or #RRGGBB) |
| `plotAreaBox` | object | Yes | — | Plot area bounds. Values in [0,1] = fraction of image size; values > 1 = raw pixels. Fields: `xMinPx`, `xMaxPx`, `yMinPx`, `yMaxPx` |
| `chartType` | string | No | "line" | "line", "scatter", "bar", "area" |
| `colorTolerance` | number | No | 120 | Euclidean RGB distance threshold (0-441) |
| `numPoints` | number | No | 40 | Output sample size (5-200) |

**Success Response** (200):
```json
{
  "points": [
    { "x": 0.5, "y": 12.3 },
    { "x": 2.8, "y": 15.7 }
  ],
  "pixelsMatched": 1523,
  "imageSize": { "width": 800, "height": 600 },
  "plotBounds": { "pxMin": 50, "pxMax": 750, "pyMin": 30, "pyMax": 550 }
}
```

**Error Responses**:
- `400` — Validation error (missing/invalid parameters)
- `404` — Image not found or expired (re-upload needed)
- `500` — Extraction failure (corrupt image, jimp error)

---

## Tool Definition (LLM-facing)

The `extract_chart_data` tool is registered in `excelTools.ts` and available when the host is Excel. The LLM sees:

```
extract_chart_data: Extract numerical data points from a chart/graph image
  using pixel color analysis. You MUST first analyze the image yourself
  (via vision) to determine the axis ranges, the dominant color of the
  data series, and the chart type.
```

Required parameters: `imageId`, `xAxisRange`, `yAxisRange`, `targetColor`
Optional parameters: `chartType`, `colorTolerance`, `numPoints`

---

## Testing

### Backend unit test (run from `backend/`):
```bash
node -e "
import { extractChartData } from './src/services/plotDigitizerService.js'
import { Jimp } from 'jimp'

const img = new Jimp({ width: 200, height: 100, color: 0xFFFFFFFF })
for (let x = 20; x <= 180; x++) {
  const y = Math.round(80 - (x - 20) * 60 / 160)
  const idx = (y * 200 + x) * 4
  img.bitmap.data[idx] = 255; img.bitmap.data[idx+1] = 0
  img.bitmap.data[idx+2] = 0; img.bitmap.data[idx+3] = 255
}
const buf = await img.getBuffer('image/png')

const result = await extractChartData({
  imageBuffer: buf,
  xAxisRange: [0, 100],
  yAxisRange: [0, 50],
  targetColor: '#FF0000',
  numPoints: 10,
})
console.log('Points:', result.points.length, '| Expected: 10')
console.log('Data should be linear from ~(5, 3) to ~(95, 47)')
console.log('Sample:', result.points.slice(0, 3))
"
```

### Full integration test:
```bash
node -e "
import { storeImage, getImage } from './src/services/imageStore.js'

const buf = Buffer.from('test')
const id = storeImage(buf, 'image/png')
const entry = getImage(id)
console.log('Store test:', entry ? 'PASS' : 'FAIL')
console.log('Bad ID test:', getImage('nonexistent') === null ? 'PASS' : 'FAIL')
"
```

---

## Limitations

1. **Single-color extraction**: Each call extracts ONE data series. For multi-series charts, call once per color.
2. **Bounding box auto-detection**: Works best when the data pixels define the plot area. Charts with colored backgrounds or gridlines may need higher tolerance.
3. **No axis detection**: The service relies on the LLM to read axis labels. If the LLM misreads them, the extracted values will be proportionally off.
4. **Pie/donut charts**: Not supported — pixel scanning doesn't work well for radial layouts. The LLM should read percentages visually.
5. **Image expiry**: Stored images expire after 30 minutes. The user must re-upload if the session is long.

## Dependencies

- **jimp** (v1.x): Pure JavaScript image processing. No native binaries, no GPU required. Safe for NAS deployment.
