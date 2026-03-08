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
- LLM calls `extract_chart_data` tool with semantic parameters + `imageId`
- Frontend sends `POST /api/chart-extract` with JSON payload
- Backend retrieves image buffer from ImageStore by `imageId`
- **PlotDigitizerService** runs the algorithm:
  1. Loads image with **jimp** (pure JS, no native deps)
  2. Scans all pixels, computes Euclidean RGB distance to target color
  3. Keeps pixels within `colorTolerance` (default: 80)
  4. Auto-detects plot area bounding box from matched pixels
  5. Maps pixel (x, y) → real values via axis ranges (rule of three)
  6. Buckets by X (or Y for bar charts), averages to produce ~40 points
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
| `backend/src/services/plotDigitizerService.js` | Core pixel scanning algorithm inspired by WebPlotDigitizer |
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

### Core: WebPlotDigitizer Algorithms (AGPL-3.0)

The extraction engine uses adapted algorithms from [WebPlotDigitizer](https://github.com/ankitrohatgi/WebPlotDigitizer) (Copyright 2025 Ankit Rohatgi, AGPL-3.0). The adapted source files live in `backend/src/services/wpd/`.

**Pipeline:**

1. **Color detection** (`dist3d` from `wpd/mathFunctions.js`):
   - Euclidean RGB distance: `d = sqrt((R1-R2)^2 + (G1-G2)^2 + (B1-B2)^2)`
   - Pixels with `d <= colorTolerance` (default: 120, WPD default) are stored in a `Set<pixelIndex>`
   - Transparent pixels (alpha=0) are treated as white (WPD convention)

2. **Line/scatter extraction** (`AveragingWindowCore` from `wpd/averagingWindowCore.js`):
   - Scans each image column vertically for "blobs" (groups of matching pixels separated by `yStep` gap)
   - Averages the Y position of each blob per column
   - Second pass: merges nearby points within a `(xStep, yStep)` proximity window
   - Much more accurate than naive bucket-averaging

3. **Bar chart extraction** (`BarExtractionAlgo` from `wpd/barExtraction.js`):
   - Vertical bars: for each column, finds top and bottom matching pixels
   - Horizontal bars: for each row, finds left and right matching pixels
   - Groups detected edges by proximity (`delX`, `delVal` thresholds)
   - Returns center position of each bar group

4. **Coordinate mapping** (our adapter in `plotDigitizerService.js`):
   ```
   realX = xMin + ((pixelX - pxMin) / pxSpan) * (xMax - xMin)
   realY = yMax - ((pixelY - pyMin) / pySpan) * (yMax - yMin)  // Y inverted
   ```

5. **Cubic spline smoothing** (`cspline` / `csplineInterp` from `wpd/mathFunctions.js`):
   - For line/area charts with many raw points, interpolates at `numPoints` evenly spaced X values
   - Produces clean, smooth output curves

### Adapted WPD Files

| File | WPD Source | Purpose |
|---|---|---|
| `wpd/mathFunctions.js` | `core/mathFunctions.js` | `dist3d`, `cspline`, `csplineInterp` |
| `wpd/averagingWindowCore.js` | `core/curve_detection/averagingWindowCore.js` | Column blob detection + neighbor merging |
| `wpd/barExtraction.js` | `core/curve_detection/barExtraction.js` | Bar top/bottom edge detection + grouping |
| `wpd/dataset.js` | `core/dataset.js` | Minimal `{x, y}` point container |

**Adaptations made** (structural only, algorithm logic preserved):
- Converted `var wpd = wpd || {}` global namespace → ESM `export class` / `export function`
- Removed UI-specific serialization/deserialization methods
- Removed dependencies on WPD's axes calibration (replaced with linear mapping)
- `BarExtractionAlgo`: constructor accepts orientation directly instead of axes object

### Key Design Decisions

- **Pure JavaScript**: Uses `jimp` (no native binaries like canvas/opencv) for NAS compatibility
- **LLM-driven semantics**: The LLM determines axis ranges and colors (human-level understanding of charts), the backend does pure math (pixel → coordinates)
- **Hybrid approach**: Vision AI for semantic understanding + deterministic WPD algorithm for precise extraction
- **WPD-grade accuracy**: Uses the same blob detection and spline interpolation as the industry-standard WebPlotDigitizer
- **Graceful degradation**: Returns a warning with 0 points if no pixels match, suggesting parameter adjustments

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
  "chartType": "line",
  "colorTolerance": 80,
  "numPoints": 40
}
```

| Parameter | Type | Required | Default | Description |
|---|---|---|---|---|
| `imageId` | string | Yes | — | UUID from upload response |
| `xAxisRange` | [number, number] | Yes | — | Real X axis [min, max] |
| `yAxisRange` | [number, number] | Yes | — | Real Y axis [min, max] |
| `targetColor` | string | Yes | — | Hex color (#RGB or #RRGGBB) |
| `chartType` | string | No | "line" | "line", "scatter", "bar", "area" |
| `colorTolerance` | number | No | 80 | RGB distance threshold (0-441) |
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
