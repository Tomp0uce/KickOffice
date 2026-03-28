/// <reference types="vite/client" />

// Fallback: augment ImportMeta so import.meta.env resolves without vite/client being installed.
interface ImportMeta {
  readonly env: Record<string, string | boolean | undefined>;
}

declare const __APP_VERSION__: string;

// Minimal Office.js global declarations so tsc resolves Office globals without @types/office-js.
// The actual runtime types come from the Office.js script loaded by the add-in manifest.
// `declare const X: any` provides the runtime value; `declare namespace X` provides type-position
// members (e.g. `context: Excel.RequestContext`, `as Word.Alignment`).
// All members are typed as `any` — real types live in @types/office-js.

declare const Excel: any;
declare namespace Excel {
  type RequestContext = any;
  type Worksheet = any;
  type Range = any;
  type HorizontalAlignment = any;
  type VerticalAlignment = any;
  type BorderLineStyle = any;
  type BorderIndex = any;
  type BorderWeight = any;
  type ChartType = any;
  type ChartSeriesBy = any;
  type ConditionalCellValueOperator = any;
  type ConditionalFormatColorCriterionType = any;
  type ConditionalFormatRuleType = any;
  type ConditionalFormatType = any;
  type ConditionalTextOperator = any;
  type IconSet = any;
  type ProtectionSelectionMode = any;
  type ConditionalFormat = any;
  type ConditionalRangeFormat = any;
  type RangeAreas = any;
  type NamedItem = any;
  type Chart = any;
  type PivotTable = any;
}

declare const Word: any;
declare namespace Word {
  type RequestContext = any;
  type Range = any;
  type Alignment = any;
  type BuiltInStyleName = any;
  type InsertLocation = any;
}

declare const PowerPoint: any;
declare namespace PowerPoint {
  type RequestContext = any;
}

declare namespace Office {
  const context: any;
  function onReady(callback: () => Promise<void> | void): void;
  const HostType: any;
  const EventType: any;
  const AsyncResultStatus: any;
  const CoercionType: any;
  type AsyncResult<T = any> = { value: T; status: string; error?: any };
}

interface Window {
  Office: any;
}

declare module '*.md?raw' {
  const content: string;
  export default content;
}
