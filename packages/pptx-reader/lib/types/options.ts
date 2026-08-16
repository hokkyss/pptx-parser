/** Options for configuring the PPTX parsing behavior. */
export interface PptxParseOptions {
  /** Whether to extract media binary data. @default true */
  includeMedia?: boolean;
  /** If true, media data is loaded lazily via getter functions. @default false */
  lazyMedia?: boolean;
  /** Whether to parse animation timelines. @default true */
  parseAnimations?: boolean;
  /** Whether to parse slide transitions. @default true */
  parseTransitions?: boolean;
  /** Whether to preserve custom XML data parts. @default false */
  customXml?: boolean;
}
