# PowerPoint Office.js Skill

## CRITICAL POWERPOINT-SPECIFIC RULES

### Rule 1: PowerPoint HAS a host-specific API (PowerPoint.run)

Unlike older documentation suggests, modern PowerPoint DOES support `PowerPoint.run()`:

```javascript
await PowerPoint.run(async (context) => {
  const slides = context.presentation.slides;
  slides.load('items');
  await context.sync();
  // Work with slides...
});
```

### Rule 2: Slide numbers are 1-based in UI, but arrays are 0-indexed

```javascript
// User says "slide 3" → array index 2
const slides = context.presentation.slides;
slides.load('items');
await context.sync();

const slideThree = slides.items[2];  // Index 2 = slide 3
```

### Rule 3: Shape text is accessed via textFrame.textRange

```javascript
const shape = slide.shapes.getItem(shapeId);
const textRange = shape.textFrame.textRange;
textRange.load('text');
await context.sync();

console.log(textRange.text);  // The text content

// To modify:
textRange.text = 'New text';
await context.sync();
```

### Rule 4: Common API still available for basic operations

For simple text operations, the Common API works:
```javascript
Office.context.document.getSelectedDataAsync(
  Office.CoercionType.Text,
  (result) => {
    console.log(result.value);
  }
);
```

### Rule 5: Keep bullet points SHORT and FEW

**Content guidelines:**
- Max 8-10 words per bullet point
- Max 6-7 bullets per slide
- Use active voice, present tense
- No full sentences — fragments are better

**WRONG:**
```
- The implementation of the new system will require careful consideration of multiple factors including budget constraints and timeline requirements.
```

**CORRECT:**
```
- New system implementation
- Budget constraints
- Timeline requirements
```

## AVAILABLE TOOLS

### For READING:
| Tool | When to use |
|------|-------------|
| `getAllSlidesOverview` | Get text overview of all slides |
| `getSlideContent` | Read all text from specific slide |
| `getShapes` | Discover shape IDs/names on a slide |
| `getSelectedText` | Read current text selection |

### For WRITING:
| Tool | When to use |
|------|-------------|
| `insertContent` | **PREFERRED** — Update shape text with Markdown |
| `proposeShapeTextRevision` | Edit shape text with diff tracking |
| `addSlide` | Create new slide |
| `deleteSlide` | Remove a slide |

### ESCAPE HATCH:
| Tool | When to use |
|------|-------------|
| `eval_powerpointjs` | Speaker notes, images, animations |

## WORKFLOW: Always discover before modifying

```
1. Call getAllSlidesOverview → understand presentation structure
2. Call getShapes(slideNumber) → get shape IDs on target slide
3. Call insertContent or proposeShapeTextRevision → modify specific shape
```

## WORKFLOW: Adding a new slide with content

Use `addSlide` with `title` and `body` parameters to create a slide and populate its template text boxes in one step:

```json
{ "layout": "TitleAndContent", "title": "My Slide Title", "body": "- Bullet one\n- Bullet two\n- Bullet three" }
```

The tool will automatically find the title and body shapes of the layout and fill them.
**Do NOT** call `getShapes` then `insertContent` manually — use `addSlide` with content directly.

## WORKFLOW: Creating a slide from an image

When the user provides an image and asks to create a slide:

1. Call `addSlide` with an appropriate layout and a generated title/body based on what you can see in the image.
2. Call `insertImageOnSlide` with `base64Data` (the raw base64 string of the image) to place the image on the new slide.
3. **Do NOT call `getAllSlidesOverview` more than once** — it is only needed for initial discovery, not for image insertion.
4. **Never loop on `getAllSlidesOverview`** — if a slide overview returns empty shapes, skip it and proceed with the image insertion directly.

## WORKFLOW: Speaker notes

After generating speaker notes text, ALWAYS call `setSpeakerNotes` to insert them directly into the slide's notes section. Do NOT just display the notes in the chat — always persist them to the slide.

## COMMON PATTERNS

### Get all slides
```javascript
const slides = context.presentation.slides;
slides.load('items,items/id');
await context.sync();

for (const slide of slides.items) {
  console.log('Slide ID:', slide.id);
}
```

### Get shapes on a slide
```javascript
const slide = context.presentation.slides.items[0];
const shapes = slide.shapes;
shapes.load('items,items/id,items/name,items/type');
await context.sync();

for (const shape of shapes.items) {
  console.log(shape.id, shape.name, shape.type);
}
```

### Modify shape text
```javascript
const slide = context.presentation.slides.items[0];
const shape = slide.shapes.getItem('Title 1');
shape.textFrame.textRange.text = 'New Title';
await context.sync();
```

### Add speaker notes
```javascript
const slide = context.presentation.slides.items[0];
slide.notesPage.textBody.text = 'Speaker notes go here...';
await context.sync();
```

## LIMITATIONS

- **No Track Changes** — PowerPoint has no equivalent to Word's Track Changes
- **Limited formatting API** — Some rich text operations require workarounds
- **Shape positioning** — Coordinates are in points, not pixels
- **Animation API** — Very limited in Office.js
