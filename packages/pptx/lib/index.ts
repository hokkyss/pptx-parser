export {
  type AddSlideOptions,
  type CreatePresentationOptions,
  Presentation,
} from './presentation';

export type {
  ThemeColorInput,
  ThemeFontInput,
} from '@hokkyss/pptx-core';

export {
  type AddConnectorOptions,
  type AddGroupOptions,
  type AddImageOptions,
  type AddTextOptions,
  GroupBuilder,
  Slide,
} from './slide';

export { SlideMaster } from './slide-master';

export {
  type AddTableOptions,
  type CellConfig,
  type RowConfig,
  TableBuilder,
  type TableMatrix,
} from './builders/table-builder';

export {
  type AddChartOptions,
  buildChartElement,
  type ChartSeriesConfig,
} from './builders/chart-builder';

export {
  type AddShapeOptions,
  buildShapeElement,
  normalizeFill,
} from './builders/shape-builder';

export {
  buildTextBody,
  buildTextRun,
  type BulletInput,
  normalizeBullet,
  type ParagraphConfig,
  type TextOptions,
  type TextRunConfig,
} from './builders/text-builder';

export * from './units';
