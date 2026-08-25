import type {
  PptxChart,
  PptxChartAxis,
  PptxChartDataLabels,
  PptxChartLegend,
  PptxChartSeries,
  PptxChartType,
  PptxElement,
  PptxFill,
  Points,
} from '@hokkyss/pptx-core';
import {
  emu,
  emuDegree,
  type Inches,
  inchesToEmu,
} from '@hokkyss/pptx-core';
import { normalizeFill } from './shape-builder';

export interface ChartSeriesConfig {
  color?: string;
  colors?: string[];
  dataPointColors?: string[];
  fill?: PptxFill | string;
  name?: string;
  values: number[];
}

export interface AddChartOptions {
  axisColor?: string;
  catAxis?: PptxChartAxis;
  categories?: string[];
  chartType?: ({} & string) | PptxChartType;
  colors?: string[];
  dataLabels?: PptxChartDataLabels;
  gridColor?: string;
  grouping?: 'clustered' | 'percentStacked' | 'stacked' | 'standard';
  h?: Inches;
  holeSize?: number;
  /**
   * Unique element identifier.
   *
   * **ID Scoping**: Scoped per slide
   * - Must be unique among all elements on the SAME slide.
   * - Identical IDs may be reused on different slides without collision.
   * - Used connector endpoint attachment (`slide.addConnector({ from: { shapeId } })`).
   */
  id?: string;
  legend?: {
    color?: string;
    fontSize?: Points;
    overlay?: boolean;
    position?: 'bottom' | 'left' | 'right' | 'top' | 'topRight';
  };
  name?: string;
  series?: ChartSeriesConfig[];
  showGridlines?: boolean;
  smooth?: boolean;
  textColor?: string;
  title?: string;
  valAxis?: PptxChartAxis;
  w?: Inches;
  x?: Inches;
  y?: Inches;
  zIndex?: number;
}

/**
 * Builds a strongly-typed `PptxElement` representing a chart graphic frame.
 */
export function buildChartElement(
  options: AddChartOptions,
  counter: number | string = 1,
): PptxElement {
  const id = options.id || String(counter);
  const name = options.name || `Chart ${id}`;

  const widthEmu = options.w ? inchesToEmu(options.w) : emu(9144000);
  const heightEmu = options.h ? inchesToEmu(options.h) : emu(4572000);
  const xEmu = options.x ? inchesToEmu(options.x) : emu(914400);
  const yEmu = options.y ? inchesToEmu(options.y) : emu(1828800);

  const rawChartType = String(options.chartType || '').toLowerCase();
  const isRadialOrPie = rawChartType.includes('pie') || rawChartType.includes('doughnut') || rawChartType.includes('donut');

  const series: PptxChartSeries[] = (options.series || []).map((s, idx) => {
    const dataPointColors = s.dataPointColors || s.colors || (idx === 0 ? options.colors : undefined);

    let fill: PptxFill | undefined;
    if (!isRadialOrPie) {
      if (s.fill) {
        fill = normalizeFill(s.fill);
      } else if (s.color) {
        fill = {
          solidColor: {
            type: 'srgb',
            value: s.color.replace(/^#/, ''),
          },
          type: 'solid',
        };
      }
    }

    return {
      dataPointColors,
      fill,
      index: idx,
      name: s.name || `Series ${idx + 1}`,
      order: idx,
      values: s.values,
    };
  });

  const legend: PptxChartLegend | undefined = options.legend
    ? {
        ...options.legend,
        color: options.legend.color || options.textColor,
      }
    : options.textColor
      ? { color: options.textColor, position: 'right' }
      : undefined;

  const catAxis: PptxChartAxis | undefined = options.catAxis || options.textColor || options.axisColor
    ? {
        axisColor: options.catAxis?.axisColor || options.axisColor,
        color: options.catAxis?.color || options.textColor,
        fontSize: options.catAxis?.fontSize,
        gridlineColor: options.catAxis?.gridlineColor,
        showGridlines: options.catAxis?.showGridlines,
      }
    : undefined;

  const valAxis: PptxChartAxis | undefined = options.valAxis || options.textColor || options.axisColor || options.gridColor || options.showGridlines
    ? {
        axisColor: options.valAxis?.axisColor || options.axisColor,
        color: options.valAxis?.color || options.textColor,
        fontSize: options.valAxis?.fontSize,
        gridlineColor: options.valAxis?.gridlineColor || options.gridColor,
        showGridlines: options.valAxis?.showGridlines ?? options.showGridlines,
      }
    : undefined;

  const chartAst: PptxChart = {
    catAxis,
    categories: options.categories || [],
    chartType: options.chartType || 'barChart',
    dataLabels: options.dataLabels,
    grouping: options.grouping,
    holeSize: options.holeSize,
    legend,
    series,
    smooth: options.smooth,
    title: options.title,
    valAxis,
  };

  return {
    chart: chartAst,
    elementType: 'chart',
    id,
    isVisible: true,
    name,
    position: {
      cx: widthEmu,
      cy: heightEmu,
      x: xEmu,
      y: yEmu,
    },
    rotation: emuDegree(0),
    type: 'graphicFrame',
    zIndex: options.zIndex ?? 0,
  };
}
