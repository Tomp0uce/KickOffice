/// <reference types="vite/client" />

// Fallback: augment ImportMeta so import.meta.env resolves without vite/client being installed.
interface ImportMeta {
  readonly env: Record<string, string | boolean | undefined>;
}

declare const __APP_VERSION__: string;

// Minimal Office.js global declarations so tsc resolves Office globals without @types/office-js.
// The actual runtime types come from the Office.js script loaded by the add-in manifest.
// We declare them as `any` to avoid re-declaring every sub-member; real types live in @types/office-js.
/* eslint-disable @typescript-eslint/no-explicit-any */
declare const Excel: any;
declare const Word: any;
declare const PowerPoint: any;
// eslint-disable-next-line @typescript-eslint/no-namespace
declare namespace Office {
  const context: any;
  function onReady(callback: () => Promise<void> | void): void;
  const HostType: any;
  const EventType: any;
  const AsyncResultStatus: any;
  const CoercionType: any;
  type AsyncResult<T = any> = { value: T; status: string; error?: any };
}
/* eslint-enable @typescript-eslint/no-explicit-any */

interface Window {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  Office: any;
}

declare module '*.md?raw' {
  const content: string;
  export default content;
}
