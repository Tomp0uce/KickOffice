import { describe, it, expect, beforeEach } from 'vitest';
import { useLoopDetection } from '../useLoopDetection';

describe('useLoopDetection', () => {
  describe('default params (windowSize=5, repeatThreshold=2)', () => {
    let detector: ReturnType<typeof useLoopDetection>;

    beforeEach(() => {
      detector = useLoopDetection();
    });

    it('returns false for the first occurrence of a signature', () => {
      expect(detector.addSignatureAndCheckLoop('tool_A{}')).toBe(false);
    });

    it('returns true when the same signature appears repeatThreshold times', () => {
      detector.addSignatureAndCheckLoop('tool_A{}');
      expect(detector.addSignatureAndCheckLoop('tool_A{}')).toBe(true);
    });

    it('does not trigger for different signatures', () => {
      detector.addSignatureAndCheckLoop('tool_A{"cell":"A1"}');
      expect(detector.addSignatureAndCheckLoop('tool_A{"cell":"B2"}')).toBe(false);
    });

    it('returns false for an empty signature', () => {
      expect(detector.addSignatureAndCheckLoop('')).toBe(false);
      expect(detector.addSignatureAndCheckLoop('')).toBe(false);
    });

    it('sliding window evicts old signatures beyond windowSize', () => {
      // Fill window with 5 different signatures, then re-add the first one
      detector.addSignatureAndCheckLoop('sig_A');
      detector.addSignatureAndCheckLoop('sig_B');
      detector.addSignatureAndCheckLoop('sig_C');
      detector.addSignatureAndCheckLoop('sig_D');
      detector.addSignatureAndCheckLoop('sig_E');
      // sig_A is now evicted (window full at 5)
      expect(detector.addSignatureAndCheckLoop('sig_A')).toBe(false);
    });

    it('clearSignatures resets the window so subsequent checks start fresh', () => {
      detector.addSignatureAndCheckLoop('loop_sig');
      detector.clearSignatures();
      // After clear, two new occurrences are needed to trigger again
      expect(detector.addSignatureAndCheckLoop('loop_sig')).toBe(false);
      expect(detector.addSignatureAndCheckLoop('loop_sig')).toBe(true);
    });
  });

  describe('custom params', () => {
    it('respects a repeatThreshold of 3', () => {
      const detector = useLoopDetection(10, 3);
      detector.addSignatureAndCheckLoop('sig');
      detector.addSignatureAndCheckLoop('sig');
      expect(detector.addSignatureAndCheckLoop('sig')).toBe(true);
    });

    it('does not trigger on 2 occurrences when threshold is 3', () => {
      const detector = useLoopDetection(10, 3);
      detector.addSignatureAndCheckLoop('sig');
      expect(detector.addSignatureAndCheckLoop('sig')).toBe(false);
    });

    it('respects a windowSize of 2 (evicts quickly)', () => {
      const detector = useLoopDetection(2, 2);
      detector.addSignatureAndCheckLoop('sig_A'); // window: [sig_A]
      detector.addSignatureAndCheckLoop('sig_B'); // window: [sig_A, sig_B]
      // Adding sig_C evicts sig_A: window becomes [sig_B, sig_C]
      detector.addSignatureAndCheckLoop('sig_C');
      // sig_A is gone — adding it now should return false
      expect(detector.addSignatureAndCheckLoop('sig_A')).toBe(false);
    });
  });

  describe('interleaved tool calls', () => {
    it('counts non-consecutive occurrences within the window', () => {
      const detector = useLoopDetection();
      detector.addSignatureAndCheckLoop('tool_A{}');
      detector.addSignatureAndCheckLoop('tool_B{}');
      // tool_A appears a second time — still within the window
      expect(detector.addSignatureAndCheckLoop('tool_A{}')).toBe(true);
    });
  });
});
