import type {
  PptxChart,
  PptxChartAxis,
  PptxChartDataLabels,
  PptxChartLegend,
  PptxChartSeries,
} from '@hokkyss/pptx-core';

export type NormalizedChartType
  = | 'areaChart'
    | 'barChart'
    | 'doughnutChart'
    | 'lineChart'
    | 'pieChart'
    | 'radarChart'
    | 'scatterChart';

/**
 * Normalizes user/AST chart type string to standard OpenXML chart element name.
 */
export function normalizeChartType(chartType: string): NormalizedChartType {
  const lower = chartType.toLowerCase().replace(/chart$/, '');
  if (lower === 'donut' || lower === 'doughnut') return 'doughnutChart';
  if (lower === 'pie') return 'pieChart';
  if (lower === 'line') return 'lineChart';
  if (lower === 'area') return 'areaChart';
  if (lower === 'radar' || lower === 'spider') return 'radarChart';
  if (lower === 'scatter') return 'scatterChart';
  if (
    lower === 'bar'
    || lower === 'column'
    || lower === 'col'
    || lower === 'horizontalbar'
    || lower === 'stackedbar'
    || lower === 'stackedcolumn'
  ) {
    return 'barChart';
  }
  return 'barChart';
}

/**
 * Normalizes legend position string to OpenXML `<c:legendPos @_val>`.
 */
function normalizeLegendPos(pos?: string): string {
  switch (pos) {
    case 'bottom':
      return 'b';
    case 'left':
      return 'l';
    case 'top':
      return 't';
    case 'topRight':
      return 'tr';
    case 'right':
    default:
      return 'r';
  }
}

/**
 * Serializes data labels configuration `<c:dLbls>`.
 */
function serializeDataLabels(dLbls?: PptxChartDataLabels): string {
  const showVal = dLbls?.showVal ? '1' : '0';
  const showCatName = dLbls?.showCatName ? '1' : '0';
  const showSerName = dLbls?.showSerName ? '1' : '0';
  const showPercent = dLbls?.showPercent ? '1' : '0';

  return `<c:dLbls><c:showLegendKey val="0"/><c:showVal val="${showVal}"/><c:showCatName val="${showCatName}"/><c:showSerName val="${showSerName}"/><c:showPercent val="${showPercent}"/><c:showBubbleSize val="0"/></c:dLbls>`;
}

/**
 * Serializes text properties `<c:txPr>` for legend or axis labels.
 */
function serializeTextProperties(color?: string, fontSize?: number): string {
  if (!color && !fontSize) return '';
  const szAttr = fontSize ? ` sz="${Math.round(fontSize * 100)}"` : '';
  let fillXml = '';
  if (color) {
    const hex = color.replace(/^#/, '');
    fillXml = `<a:solidFill><a:srgbClr val="${hex}"/></a:solidFill>`;
  }
  return `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr${szAttr}>${fillXml}</a:defRPr></a:pPr><a:endParaRPr/></a:p></c:txPr>`;
}

/**
 * Serializes shape line properties `<c:spPr>`.
 */
function serializeLineProperties(lineColor?: string): string {
  if (!lineColor) return '';
  const hex = lineColor.replace(/^#/, '');
  return `<c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="${hex}"/></a:solidFill></a:ln></c:spPr>`;
}

/**
 * Serializes chart legend.
 */
function serializeLegend(legend?: PptxChartLegend): string {
  const pos = normalizeLegendPos(legend?.position);
  const overlay = legend?.overlay ? '1' : '0';
  const txPrXml = serializeTextProperties(legend?.color, legend?.fontSize);
  return `<c:legend><c:legendPos val="${pos}"/><c:overlay val="${overlay}"/>${txPrXml}</c:legend>`;
}

/**
 * Serializes Category (X) Axis `<c:catAx>`.
 */
function serializeCategoryAxis(catAxId: string, valAxId: string, axis?: PptxChartAxis): string {
  const spPrXml = serializeLineProperties(axis?.axisColor);
  const txPrXml = serializeTextProperties(axis?.color, axis?.fontSize);
  const majorGridlinesXml = (axis?.showGridlines || axis?.gridlineColor)
    ? `<c:majorGridlines>${serializeLineProperties(axis?.gridlineColor || '334155')}</c:majorGridlines>`
    : '';

  return `<c:catAx><c:axId val="${catAxId}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/>${majorGridlinesXml}<c:numFmt formatCode="General" sourceLinked="0"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>${spPrXml}${txPrXml}<c:crossAx val="${valAxId}"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="1"/></c:catAx>`;
}

/**
 * Serializes Value (Y) Axis `<c:valAx>`.
 */
function serializeValueAxis(valAxId: string, catAxId: string, axis?: PptxChartAxis): string {
  const spPrXml = serializeLineProperties(axis?.axisColor);
  const txPrXml = serializeTextProperties(axis?.color, axis?.fontSize);
  const majorGridlinesXml = (axis?.showGridlines || axis?.gridlineColor)
    ? `<c:majorGridlines>${serializeLineProperties(axis?.gridlineColor || '334155')}</c:majorGridlines>`
    : '';

  return `<c:valAx><c:axId val="${valAxId}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/>${majorGridlinesXml}<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>${spPrXml}${txPrXml}<c:crossAx val="${catAxId}"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>`;
}

/**
 * Serializes a single chart series `<c:ser>`.
 */
function serializeSeries(
  ser: PptxChartSeries,
  idx: number,
  categories: string[],
  chartType: NormalizedChartType,
  smooth?: boolean,
): string {
  const serIndex = ser.index !== undefined ? ser.index : idx;
  const serOrder = ser.order !== undefined ? ser.order : idx;
  const serName = ser.name || `Series ${idx + 1}`;

  let spPrXml = '';
  if (ser.fill && chartType !== 'pieChart' && chartType !== 'doughnutChart') {
    if (ser.fill.type === 'solid' && ser.fill.solidColor?.value) {
      const hex = ser.fill.solidColor.value.replace(/^#/, '');
      if (chartType === 'lineChart' || chartType === 'radarChart' || chartType === 'scatterChart') {
        spPrXml = `<c:spPr><a:ln w="25400"><a:solidFill><a:srgbClr val="${hex}"/></a:solidFill></a:ln></c:spPr>`;
      } else {
        spPrXml = `<c:spPr><a:solidFill><a:srgbClr val="${hex}"/></a:solidFill></c:spPr>`;
      }
    }
  }

  // Categories (<c:cat> or <c:xVal> for scatter)
  let catXml = '';
  if (categories && categories.length > 0) {
    const pts = categories
      .map((cat, i) => `<c:pt idx="${i}"><c:v>${escapeXml(cat)}</c:v></c:pt>`)
      .join('');
    if (chartType === 'scatterChart') {
      catXml = `<c:xVal><c:strLit><c:ptCount val="${categories.length}"/>${pts}</c:strLit></c:xVal>`;
    } else {
      catXml = `<c:cat><c:strLit><c:ptCount val="${categories.length}"/>${pts}</c:strLit></c:cat>`;
    }
  }

  // Numerical Values (<c:val> or <c:yVal> for scatter)
  const values = ser.values || [];
  const valPts = values
    .map((val, i) => `<c:pt idx="${i}"><c:v>${val}</c:v></c:pt>`)
    .join('');
  let valXml = '';
  if (chartType === 'scatterChart') {
    valXml = `<c:yVal><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>${valPts}</c:numLit></c:yVal>`;
  } else {
    valXml = `<c:val><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>${valPts}</c:numLit></c:val>`;
  }

  // Individual Data Point / Slice Colors (<c:dPt>)
  let dPtXml = '';
  const dPtColors = ser.dataPointColors;

  if (dPtColors && dPtColors.length > 0) {
    dPtXml = (ser.values || [])
      .map((_, i) => {
        const color = dPtColors[i % dPtColors.length].replace(/^#/, '');
        return `<c:dPt><c:idx val="${i}"/><c:spPr><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></c:spPr></c:dPt>`;
      })
      .join('');
  }

  const smoothXml = (smooth && (chartType === 'lineChart' || chartType === 'scatterChart'))
    ? '<c:smooth val="1"/>'
    : '';

  return `<c:ser><c:idx val="${serIndex}"/><c:order val="${serOrder}"/><c:tx><c:v>${escapeXml(serName)}</c:v></c:tx>${spPrXml}<c:invertIfNegative val="0"/>${dPtXml}${catXml}${valXml}${smoothXml}</c:ser>`;
}

/**
 *
 */
function escapeXml(unsafe: string): string {
  return String(unsafe)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

/**
 * Serializes a strongly-typed `PptxChart` AST element into ECMA-376 OpenXML `<c:chartSpace>` format.
 */
export function serializeChart(chart: PptxChart): string {
  const normType = normalizeChartType(chart.chartType || 'barChart');
  const rawType = String(chart.chartType || '').toLowerCase();
  const categories = chart.categories || [];
  const series = chart.series || [];

  const serXmlList = series
    .map((s, i) => serializeSeries(s, i, categories, normType, chart.smooth))
    .join('');

  const catAxId = '111111111';
  const valAxId = '222222222';
  const dLblsXml = serializeDataLabels(chart.dataLabels);

  let chartBody = '';

  if (normType === 'barChart') {
    const isHorizontal = rawType.includes('horizontal') || rawType === 'bar';
    const barDir = isHorizontal ? 'bar' : 'col';

    let grouping = chart.grouping || 'clustered';
    if (rawType.includes('stacked') || rawType.includes('percent')) {
      grouping = rawType.includes('percent') ? 'percentStacked' : 'stacked';
    }

    chartBody = `<c:barChart><c:barDir val="${barDir}"/><c:grouping val="${grouping}"/><c:varyColors val="0"/>${serXmlList}${dLblsXml}<c:gapWidth val="150"/><c:axId val="${catAxId}"/><c:axId val="${valAxId}"/></c:barChart>`;
  } else if (normType === 'lineChart') {
    const grouping = chart.grouping || (rawType.includes('stacked') ? (rawType.includes('percent') ? 'percentStacked' : 'stacked') : 'standard');
    chartBody = `<c:lineChart><c:grouping val="${grouping}"/><c:varyColors val="0"/>${serXmlList}${dLblsXml}<c:axId val="${catAxId}"/><c:axId val="${valAxId}"/></c:lineChart>`;
  } else if (normType === 'pieChart') {
    chartBody = `<c:pieChart><c:varyColors val="1"/>${serXmlList}${dLblsXml}</c:pieChart>`;
  } else if (normType === 'doughnutChart') {
    const holeSize = chart.holeSize !== undefined ? chart.holeSize : 50;
    chartBody = `<c:doughnutChart><c:varyColors val="1"/>${serXmlList}${dLblsXml}<c:holeSize val="${holeSize}"/></c:doughnutChart>`;
  } else if (normType === 'areaChart') {
    const grouping = chart.grouping || (rawType.includes('stacked') ? (rawType.includes('percent') ? 'percentStacked' : 'stacked') : 'standard');
    chartBody = `<c:areaChart><c:grouping val="${grouping}"/><c:varyColors val="0"/>${serXmlList}${dLblsXml}<c:axId val="${catAxId}"/><c:axId val="${valAxId}"/></c:areaChart>`;
  } else if (normType === 'radarChart') {
    chartBody = `<c:radarChart><c:radarStyle val="standard"/><c:varyColors val="0"/>${serXmlList}${dLblsXml}<c:axId val="${catAxId}"/><c:axId val="${valAxId}"/></c:radarChart>`;
  } else if (normType === 'scatterChart') {
    chartBody = `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${serXmlList}${dLblsXml}<c:axId val="${catAxId}"/><c:axId val="${valAxId}"/></c:scatterChart>`;
  }

  let axesXml = '';
  const isPolarOrRadial = normType === 'pieChart' || normType === 'doughnutChart';
  if (!isPolarOrRadial) {
    const catAxXml = serializeCategoryAxis(catAxId, valAxId, chart.catAxis);
    const valAxXml = serializeValueAxis(valAxId, catAxId, chart.valAxis);
    axesXml = `${catAxXml}${valAxXml}`;
  }

  let titleXml = '';
  if (chart.title) {
    titleXml = `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr b="1"/><a:t>${escapeXml(chart.title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`;
  }

  const legendXml = serializeLegend(chart.legend);
  const autoTitle = chart.title ? '' : '<c:autoTitleDeleted val="1"/>';

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:roundedCorners val="0"/><c:chart>${titleXml}${autoTitle}<c:plotArea><c:layout/>${chartBody}${axesXml}</c:plotArea>${legendXml}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/><c:showDLblsOverMax val="0"/></c:chart></c:chartSpace>`;
}
