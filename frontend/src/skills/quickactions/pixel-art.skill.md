# Pixel Art Quick Action Skill (Excel)

## Purpose

Convert an uploaded image into pixel art rendered directly in Excel cells using background colors.

## When to Use

- User clicks "Pixel Art" Quick Action in Excel and uploads an image
- Goal: transform the image into a colorful cell-based pixel art representation in the active worksheet

## Input Contract

- **Image**: An image uploaded by the user (available in `<uploaded_images>` context)
- **Language**: **ALWAYS respond in the UI language specified at the start of the user message as `[UI language: ...]`.**
- **Context**: Excel worksheet (active sheet receives the pixel art)

## Output Requirements

1. Pixel art rendered in Excel cells via `imageToSheet`
2. Visual verification via `screenshotRange`
3. Brief summary of what was done (dimensions, cell size)

---

## Step-by-Step Workflow

### STEP 1 — Analyze the image

Look at the uploaded image and determine:
- **Subject**: What the image depicts (useful for naming)
- **Complexity**: Whether high detail or simple shapes
- **Recommended dimensions**: Choose appropriate `maxWidth` and `maxHeight` (default 80x80 works for most images; use 40-60 for simple icons, up to 120 for detailed images)
- **Cell size**: Default `cellSize` of 3 pixels per cell works well; use 2 for more detail, 4-5 for larger cells

### STEP 2 — Generate pixel art

Call `imageToSheet` with the uploaded image data:

```json
{
  "imageBase64": "<base64 data from the uploaded image>",
  "maxWidth": 80,
  "maxHeight": 80,
  "cellSize": 3,
  "startCell": "A1"
}
```

**Tips for best results:**
- For photos/complex images: `maxWidth`/`maxHeight` 60-100, `cellSize` 3
- For logos/icons: `maxWidth`/`maxHeight` 30-50, `cellSize` 4-5
- For very detailed art: `maxWidth`/`maxHeight` 100-150, `cellSize` 2
- The tool automatically adjusts column widths and row heights to create square cells

### STEP 3 — Verify

Call `screenshotRange` on the area containing the pixel art to show the result to the user.

### STEP 4 — Report

Tell the user:
- The dimensions of the pixel art (e.g., "80x60 cells")
- The cell size used
- That they can zoom out to see the full picture

---

## Important Notes

- The `imageBase64` parameter accepts the raw base64 image data (without the `data:` prefix)
- If the image is very large or complex, prefer smaller dimensions (60x60) for faster processing
- The tool will downscale the image automatically — the original is never stretched
- Excel may need to be zoomed out (e.g., 25-50%) to see the full pixel art
- Always use `screenshotRange` after to show the user the result

---

## Edge Cases

### Image too large
If the image causes performance issues, reduce `maxWidth`/`maxHeight` to 50 and retry.

### No image data available
If the uploaded image data is not accessible, ask the user to re-upload the image.

---

## Quality Check

- The pixel art is rendered in Excel cells?
- Screenshot taken to verify the result?
- Responded in UI language?
