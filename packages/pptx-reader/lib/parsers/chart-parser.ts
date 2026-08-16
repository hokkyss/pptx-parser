import { PptxChart, PptxChartSeries, PptxChartLegend } from '../types/ast';
import { defaultXmlParser, XmlParser } from '../xml/xml-parser';

/**
 * Parses OpenXML chart files (`ppt/charts/chart1.xml`) embedded within graphic frames into structured `PptxChart` AST elements.
 *
 * Extracts chart types (`barChart`, `lineChart`, `pieChart`, `areaChart`), series data points, categories, and legend configuration.
 * @param chartXml Raw XML string or object node representing a chart file (`<c:chartSpace>`).
 * @param parser Optional custom `XmlParser` instance.
 * @returns Parsed `PptxChart` object or `undefined` if chart space is invalid.
 * @example
 * ```ts
 * const chart = parseChart(chartXmlString);
 * console.log(`Chart type: ${chart?.chartType}, Series count: ${chart?.series.length}`);
 * ```
 */
export function parseChart(chartXml: Record<string, unknown> | string, parser: XmlParser = defaultXmlParser): PptxChart | undefined {
  let parsed: Record<string, unknown>;
  if (typeof chartXml === 'string') {
    parsed = parser.parse<Record<string, unknown>>(chartXml);
  } else {
    parsed = chartXml;
  }

  const chartSpace = (parsed['c:chartSpace'] || parsed['chartSpace'] || parsed) as Record<string, unknown>;
  const chartNode = (chartSpace['c:chart'] || chartSpace['chart']) as Record<string, unknown> | undefined;
  if (!chartNode) return undefined;

  const plotArea = (chartNode['c:plotArea'] || chartNode['plotArea']) as Record<string, unknown> | undefined;
  if (!plotArea) return undefined;

  // Identify chart type (e.g. c:barChart, c:lineChart, c:pieChart, c:areaChart)
  const chartTypeKey = Object.keys(plotArea).find((k) => k.endsWith('Chart'));
  const chartType = chartTypeKey ? chartTypeKey.replace('c:', '') : 'unknownChart';

  const chartTypeNode = chartTypeKey ? (plotArea[chartTypeKey] as Record<string, unknown>) : {};

  // Extract series (<c:ser>)
  let serList = chartTypeNode['c:ser'] || chartTypeNode['ser'];
  const series: PptxChartSeries[] = [];
  let categories: string[] = [];

  if (serList) {
    if (!Array.isArray(serList)) serList = [serList];
    for (const serNode of serList as Record<string, unknown>[]) {
      const idx = serNode['c:idx'] ? Number((serNode['c:idx'] as Record<string, unknown>)['@_val']) : 0;
      const order = serNode['c:order'] ? Number((serNode['c:order'] as Record<string, unknown>)['@_val']) : 0;

      // Extract Series Name
      const txNode = (serNode['c:tx'] || serNode['tx']) as Record<string, unknown> | undefined;
      let name = `Series ${idx + 1}`;
      if (txNode) {
        const vNode = txNode['c:v'] || txNode['v'];
        if (vNode !== undefined) {
          name = typeof vNode === 'object' && '#text' in (vNode as Record<string, unknown>) ? String((vNode as Record<string, unknown>)['#text']) : String(vNode);
        }
      }

      // Extract Values
      const valNode = (serNode['c:val'] || serNode['val']) as Record<string, unknown> | undefined;
      const values = extractNumLitValues(valNode);

      // Extract Categories
      const catNode = (serNode['c:cat'] || serNode['cat']) as Record<string, unknown> | undefined;
      const parsedCats = extractStrLitValues(catNode);
      if (parsedCats.length > categories.length) {
        categories = parsedCats;
      }

      series.push({
        index: idx,
        order,
        name,
        values,
      });
    }
  }

  // Extract Legend (<c:legend>)
  const legendNode = (chartNode['c:legend'] || chartNode['legend']) as Record<string, unknown> | undefined;
  let legend: PptxChartLegend | undefined;

  if (legendNode) {
    const legendPosNode = (legendNode['c:legendPos'] || legendNode['legendPos']) as Record<string, unknown> | undefined;
    const posRaw = legendPosNode && legendPosNode['@_val'] ? String(legendPosNode['@_val']) : 'b';
    const posMap: Record<string, 'bottom' | 'left' | 'right' | 'top' | 'topRight'> = {
      b: 'bottom',
      l: 'left',
      r: 'right',
      t: 'top',
      tr: 'topRight',
    };
    const overlayNode = (legendNode['c:overlay'] || legendNode['overlay']) as Record<string, unknown> | undefined;
    const overlay = overlayNode && overlayNode['@_val'] === '1';

    legend = {
      position: posMap[posRaw] || 'bottom',
      overlay,
    };
  }

  return {
    chartType,
    series,
    categories,
    legend,
  };
}

/**
 * Helper extracting numeric value points array from chart value node (`<c:val>`).
 * @param valNode Chart value node (`<c:val>`).
 * @returns Array of parsed numeric data points.
 */
function extractNumLitValues(valNode?: Record<string, unknown>): number[] {
  if (!valNode) return [];
  const numLit = (valNode['c:numLit'] || valNode['numLit'] || valNode['c:numRef'] || valNode['numRef']) as Record<string, unknown> | undefined;
  if (!numLit) return [];

  let ptList = numLit['c:pt'] || numLit['pt'];
  if (!ptList) return [];
  if (!Array.isArray(ptList)) ptList = [ptList];

  const values: number[] = [];
  for (const pt of ptList as Record<string, unknown>[]) {
    const vNode = pt['c:v'] || pt['v'];
    if (vNode !== undefined) {
      const numVal = typeof vNode === 'object' && '#text' in (vNode as Record<string, unknown>) ? Number((vNode as Record<string, unknown>)['#text']) : Number(vNode);
      values.push(isNaN(numVal) ? 0 : numVal);
    }
  }
  return values;
}

/**
 * Helper extracting category string labels array from chart category node (`<c:cat>`).
 * @param catNode Chart category node (`<c:cat>`).
 * @returns Array of parsed category label strings.
 */
function extractStrLitValues(catNode?: Record<string, unknown>): string[] {
  if (!catNode) return [];
  const strLit = (catNode['c:strLit'] || catNode['strLit'] || catNode['c:strRef'] || catNode['strRef']) as Record<string, unknown> | undefined;
  if (!strLit) return [];

  let ptList = strLit['c:pt'] || strLit['pt'];
  if (!ptList) return [];
  if (!Array.isArray(ptList)) ptList = [ptList];

  const categories: string[] = [];
  for (const pt of ptList as Record<string, unknown>[]) {
    const vNode = pt['c:v'] || pt['v'];
    if (vNode !== undefined) {
      const strVal = typeof vNode === 'object' && '#text' in (vNode as Record<string, unknown>) ? String((vNode as Record<string, unknown>)['#text']) : String(vNode);
      categories.push(strVal);
    }
  }
  return categories;
}
