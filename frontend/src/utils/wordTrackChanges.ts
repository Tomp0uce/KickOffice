/**
 * Word Track Changes Utilities
 *
 * Manages Track Changes state during OOXML insertion.
 * Pattern from Gemini AI for Office (MIT License):
 * https://github.com/AnsonLai/Gemini-AI-for-Office-Microsoft-Word-Add-In-for-Vibe-Drafting
 */

import { logService } from './logger';

const DEFAULT_AUTHOR = 'KickOffice AI';

export interface TrackingState {
  available: boolean;
  originalMode: any | null;
  changed: boolean;
}

/**
 * Save current Track Changes mode and set desired mode.
 * Mirrors Gemini's setChangeTrackingForAi().
 *
 * When inserting OOXML with embedded w:ins/w:del, we DISABLE native tracking
 * to prevent Word from double-tracking the inserted content.
 */
export async function setChangeTrackingForAi(
  context: Word.RequestContext,
  _redlineEnabled: boolean,
  sourceLabel: string = 'AI',
): Promise<TrackingState> {
  let originalMode = null;
  let changed = false;
  let available = false;

  try {
    const doc = context.document;
    doc.load('changeTrackingMode');
    await context.sync();

    available = true;
    originalMode = doc.changeTrackingMode;

    // Always disable native tracking during OOXML insertion.
    // - redlineEnabled=true: w:ins/w:del are already embedded → native tracking would double-record
    // - redlineEnabled=false: silent replacement → no tracking desired either
    const desiredMode = Word.ChangeTrackingMode.off;

    if (originalMode !== desiredMode) {
      doc.changeTrackingMode = desiredMode;
      await context.sync();
      changed = true;
    }
  } catch (error) {
    logService.warn(`[ChangeTracking] ${sourceLabel}: unavailable`, error instanceof Error ? error : new Error(String(error)));
  }

  return { available, originalMode, changed };
}

/**
 * Restore Track Changes mode to its original state.
 * Mirrors Gemini's restoreChangeTracking().
 *
 * MUST be called in a finally block after setChangeTrackingForAi().
 */
export async function restoreChangeTracking(
  context: Word.RequestContext,
  trackingState: TrackingState,
  sourceLabel: string = 'AI',
): Promise<void> {
  if (
    !trackingState ||
    !trackingState.available ||
    !trackingState.changed ||
    trackingState.originalMode === null
  ) {
    return;
  }

  try {
    context.document.changeTrackingMode = trackingState.originalMode;
    await context.sync();
  } catch (error) {
    logService.warn(`[ChangeTracking] ${sourceLabel}: restore failed`, error instanceof Error ? error : new Error(String(error)));
  }
}

/**
 * Load redline enabled setting from localStorage.
 * Default: true (Track Changes enabled).
 */
export function loadRedlineSetting(): boolean {
  const storedSetting = localStorage.getItem('redlineEnabled');
  return storedSetting !== null ? storedSetting === 'true' : true;
}

/**
 * Load the redline author name from localStorage.
 * Default: "KickOffice AI".
 */
export function loadRedlineAuthor(): string {
  const storedAuthor = localStorage.getItem('redlineAuthor');
  if (storedAuthor && storedAuthor.trim() !== '') {
    return storedAuthor;
  }
  return DEFAULT_AUTHOR;
}
