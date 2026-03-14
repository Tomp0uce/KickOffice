/**
 * OOXML/ZIP utilities for PowerPoint slide editing.
 * Allows modifying slide XML directly via JSZip when Office.js API is insufficient.
 * Requires JSZip (install: npm install jszip in frontend/).
 */

// Dynamic import to avoid bundle issues if jszip is not installed
async function getJSZip(): Promise<any> {
  try {
    // eslint-disable-next-line @typescript-eslint/ban-ts-comment
    // @ts-ignore — jszip is an optional dependency; install with: cd frontend && npm install jszip
    const JSZip = await import('jszip');
    return JSZip.default;
  } catch {
    throw new Error('JSZip is not installed. Run: cd frontend && npm install jszip');
  }
}

/**
 * Atomically edit a slide via its OOXML ZIP:
 * 1. Export slide as base64 PPTX
 * 2. Open in JSZip
 * 3. Call callback with zip + markDirty
 * 4. If dirty, reinsert modified slide and delete original
 */
export async function withSlideZip(
  context: any,
  slideIndex: number,
  callback: (zip: any, markDirty: () => void) => Promise<any>,
): Promise<any> {
  const JSZip = await getJSZip();

  const slides = context.presentation.slides;
  slides.load('items/id');
  await context.sync();

  if (slideIndex < 0 || slideIndex >= slides.items.length) {
    throw new Error(`Slide index ${slideIndex} is out of bounds.`);
  }

  const targetSlide = slides.items[slideIndex];
  const slideId = targetSlide.id;

  // Export slide as base64
  const base64Result = (targetSlide as any).exportAsBase64();
  await context.sync();

  // Load into JSZip
  const zip = await JSZip.loadAsync(base64Result.value, { base64: true });

  let dirty = false;
  const markDirty = () => {
    dirty = true;
  };

  // Run the callback
  const result = await callback(zip, markDirty);

  if (dirty) {
    const newBase64 = await zip.generateAsync({ type: 'base64' });

    // Find preceding slide for insertion point
    slides.load('items/id');
    await context.sync();

    const prevIndex = slideIndex > 0 ? slideIndex - 1 : undefined;
    const prevSlideId = prevIndex !== undefined ? slides.items[prevIndex].id : undefined;

    context.presentation.insertSlidesFromBase64(newBase64, {
      targetSlideId: prevSlideId,
    });
    await context.sync();

    // Delete original (reload to find it)
    slides.load('items/id');
    await context.sync();

    const original = slides.items.find((s: any) => s.id === slideId);
    if (original) {
      original.delete();
      await context.sync();
    }
  }

  return result;
}

export function escapeXml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

export function sanitizeXmlAmpersands(xml: string): string {
  return xml.replace(/&(?!amp;|lt;|gt;|apos;|quot;|#\d+;|#x[0-9a-fA-F]+;)/g, '&amp;');
}
