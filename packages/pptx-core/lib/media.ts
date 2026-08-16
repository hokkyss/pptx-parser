/** Represents an embedded media asset (image, video, etc.) */
export interface PptxMediaAsset {
  /** Binary data (null if lazy loaded) */
  data: null | Uint8Array;
  /** Filename in PPTX package (e.g. 'image1.png') */
  filename: string;
  /** Alias for filename */
  fileName?: string;
  /** Unique ID */
  id: string;
  /** Function to lazily load binary data */
  lazyGetter?: () => Promise<null | Uint8Array>;
  /** MIME type (e.g. 'image/png') */
  mimeType: string;
  /** Internal path within the ZIP archive (e.g. 'ppt/media/image1.png') */
  path: string;
}
