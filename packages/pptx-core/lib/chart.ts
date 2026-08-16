import type { PptxFill } from './color';
import type { Points } from './units';

/** Strongly-typed chart axis styling */
export interface PptxChartAxis {
  /** Color of the axis line (hex without #) */
  axisColor?: string;
  /** Font color for axis tick labels (hex without #) */
  color?: string;
  /** Font size in points */
  fontSize?: Points;
  /** Major gridlines color (hex without #) */
  gridlineColor?: string;
  /** Show major gridlines */
  showGridlines?: boolean;
}

/** Strongly-typed chart legend */
export interface PptxChartLegend {
  /** Font color for legend text (hex without #) */
  color?: string;
  /** Font size in points */
  fontSize?: Points;
  /** Legend overlay setting. OpenXML: `<c:legend><c:overlay @_val>` */
  overlay?: boolean;
  /** Legend position ('bottom' | 'left' | 'right' | 'top' | 'topRight'). OpenXML: `<c:legend><c:legendPos @_val>` */
  position?: 'bottom' | 'left' | 'right' | 'top' | 'topRight';
}

/** Data label display options for charts */
export interface PptxChartDataLabels {
  /** Show category name */
  showCatName?: boolean;
  /** Show percentage (pie/doughnut) */
  showPercent?: boolean;
  /** Show series name */
  showSerName?: boolean;
  /** Show numerical value */
  showVal?: boolean;
}

/** Represents a data series in a chart */
export interface PptxChartSeries {
  /** Individual slice / data point colors for Pie/Doughnut charts. OpenXML: `<c:dPt>` */
  dataPointColors?: string[];
  /** Fill properties for the series. OpenXML: `<c:spPr><a:solidFill>` */
  fill?: PptxFill;
  /** Series index. OpenXML: `<c:idx @_val>` */
  index: number;
  /** Series name. OpenXML: `<c:tx>` */
  name: string;
  /** Series order. OpenXML: `<c:order @_val>` */
  order: number;
  /** Data values. OpenXML: `<c:val><c:numLit>` */
  values: number[];
}

export type PptxChartType
  = | 'area'
    | 'areaChart'
    | 'bar'
    | 'barChart'
    | 'bubble'
    | 'bubbleChart'
    | 'column'
    | 'donut'
    | 'doughnut'
    | 'doughnutChart'
    | 'horizontalBar'
    | 'line'
    | 'lineChart'
    | 'pie'
    | 'pieChart'
    | 'radar'
    | 'radarChart'
    | 'scatter'
    | 'scatterChart'
    | 'stackedBar'
    | 'stackedColumn';

/** Represents a chart */
export interface PptxChart {
  /** Category (X) axis configuration */
  catAxis?: PptxChartAxis;
  /** Categories (labels) for the data. OpenXML: `<c:cat>` */
  categories: string[];
  /** Type of the chart (e.g. 'barChart', 'lineChart', 'pieChart', 'doughnutChart', 'radarChart', 'areaChart'). OpenXML: `<c:plotArea><c:*Chart>` */
  chartType: ({} & string) | PptxChartType;
  /** Palette colors for multi-colored charts like Pie or Doughnut */
  colors?: string[];
  /** Data label configuration. OpenXML: `<c:dLbls>` */
  dataLabels?: PptxChartDataLabels;
  /** Series grouping style ('clustered' | 'stacked' | 'percentStacked' | 'standard') */
  grouping?: 'clustered' | 'percentStacked' | 'stacked' | 'standard';
  /** Hole size percentage for doughnut charts (default 50) */
  holeSize?: number;
  /** Legend information. OpenXML: `<c:legend>` */
  legend?: PptxChartLegend;
  /** Data series. OpenXML: `<c:ser>` */
  series: PptxChartSeries[];
  /** Spline smoothing for line and area charts */
  smooth?: boolean;
  /** Chart title text. OpenXML: `<c:title>` */
  title?: string;
  /** Value (Y) axis configuration */
  valAxis?: PptxChartAxis;
}
