# Pixel Art Quick Action Skill (Excel)

## Purpose

Convert an uploaded image into pixel art rendered directly in Excel cells using background colors.

## When to Use

- User clicks "Pixel Art" Quick Action in Excel and uploads an image
- Goal: transform the image into a colorful cell-based pixel art representation in the active worksheet

## Input Contract

- **Image**: An image uploaded by the user — its VFS path is provided in `<uploaded_images>` context as `filePath="..."`
- **Language**: **ALWAYS respond in the UI language specified at the start of the user message as `[UI language: ...]`.**
- **Context**: Excel worksheet (active sheet receives the pixel art)

## Output Requirements

1. Pixel art rendered in Excel cells via `imageToSheet`
2. Visual verification via `screenshotRange`
3. Brief summary of what was done (dimensions, cell size)

---

## Step-by-Step Workflow

### STEP 1 — Get the image path

The VFS file path is provided in the user message `<uploaded_images>` section as `filePath="..."`.
Use **that exact path** as the `filePath` parameter for `imageToSheet`.

### STEP 2 — Generate pixel art

Call `imageToSheet` with the VFS file path:

```json
{
  "filePath": "/home/user/uploads/Group_352.png",
  "width": 80,
  "height": 80,
  "cellSize": 3,
  "startCell": "A1"
}
```

**Parameter reference:**
- `filePath`: path from the `<uploaded_images>` context — use EXACTLY as provided
- `width`: target width in columns (1–200). Default 80.
- `height`: target height in rows (1–200). Default 80.
- `cellSize`: cell size in points (1–50). Default 3.
- `startCell`: top-left anchor cell. Default "A1".
- `sheetName`: optional sheet name (uses active sheet if omitted)

**Tips for best results:**
- For photos/complex images: `width`/`height` 60-100, `cellSize` 3
- For logos/icons: `width`/`height` 30-50, `cellSize` 4-5
- For very detailed art: `width`/`height` 100-150, `cellSize` 2
- The tool automatically adjusts column widths and row heights to create square cells

### STEP 3 — Verify

Call `screenshotRange` on the area containing the pixel art to show the result to the user:

```json
{ "range": "A1:CL80" }
```
(Adjust the range end based on `width` value used)

### STEP 4 — Report

Tell the user:
- The dimensions of the pixel art (e.g., "80x80 cells")
- The cell size used
- That they can zoom out (e.g., 25–50%) to see the full picture

---

## Important Notes

- `filePath` is a VFS path like `/home/user/uploads/filename.png` — use it exactly as given
- The tool downscales the image automatically — the original is never stretched
- Excel may need to be zoomed out (e.g., 25-50%) to see the full pixel art
- Always use `screenshotRange` after to show the user the result

---

## Edge Cases

### Image not found at path
If `imageToSheet` returns a file-not-found error, check the exact path from `<uploaded_images>` and retry with the correct path. Do NOT ask the user to re-upload unless the path is truly missing from the context.

### Image too large
If the image causes performance issues, reduce `width`/`height` to 50 and retry.

---

## Quality Check

- Used `filePath` from the `<uploaded_images>` context (not imageBase64)?
- The pixel art is rendered in Excel cells?
- Screenshot taken to verify the result?
- Responded in UI language?
