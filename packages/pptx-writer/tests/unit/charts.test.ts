import { describe, expect, it } from 'vitest';
import type { PptxChart } from '@hokkyss/pptx-core';
import { serializeChart } from '../../lib/serializers/chart-serializer';

describe('Chart Serializer (Unit Tests)', () => {
  it('serializes a column bar chart correctly', () => {
    const chart: PptxChart = {
      categories: ['Q1', 'Q2', 'Q3', 'Q4'],
      chartType: 'barChart',
      legend: { position: 'right' },
      series: [
        {
          fill: { solidColor: { type: 'srgb', value: '0284C7' }, type: 'solid' },
          index: 0,
          name: 'Revenue ($M)',
          order: 0,
          values: [120, 150, 180, 240],
        },
      ],
    };

    const xml = serializeChart(chart);
    expect(xml).toContain('<c:barChart>');
    expect(xml).toContain('<c:barDir val="col"/>');
    expect(xml).toContain('<c:v>Revenue ($M)</c:v>');
    expect(xml).toContain('<c:v>Q1</c:v>');
    expect(xml).toContain('<c:v>120</c:v>');
    expect(xml).toContain('<c:legendPos val="r"/>');
    expect(xml).toContain('<c:catAx>');
    expect(xml).toContain('<c:valAx>');
  });

  it('serializes a line chart with series stroke styling', () => {
    const chart: PptxChart = {
      categories: ['Jan', 'Feb', 'Mar'],
      chartType: 'lineChart',
      legend: { position: 'bottom' },
      series: [
        {
          fill: { solidColor: { type: 'srgb', value: '10B981' }, type: 'solid' },
          index: 0,
          name: 'Active Users',
          order: 0,
          values: [500, 750, 1200],
        },
      ],
    };

    const xml = serializeChart(chart);
    expect(xml).toContain('<c:lineChart>');
    expect(xml).toContain('<c:grouping val="standard"/>');
    expect(xml).toContain('<a:srgbClr val="10B981"/>');
    expect(xml).toContain('<c:legendPos val="b"/>');
  });

  it('serializes a pie chart without Cartesian axes', () => {
    const chart: PptxChart = {
      categories: ['Direct', 'Referral', 'Organic'],
      chartType: 'pieChart',
      series: [
        {
          index: 0,
          name: 'Traffic Sources',
          order: 0,
          values: [45, 25, 30],
        },
      ],
    };

    const xml = serializeChart(chart);
    expect(xml).toContain('<c:pieChart>');
    expect(xml).toContain('<c:varyColors val="1"/>');
    expect(xml).not.toContain('<c:dPt>');
    expect(xml).not.toContain('<c:catAx>');
    expect(xml).not.toContain('<c:valAx>');
  });

  it('serializes a pie chart with custom slice dataPointColors', () => {
    const chart: PptxChart = {
      categories: ['Direct', 'Referral', 'Organic'],
      chartType: 'pieChart',
      series: [
        {
          dataPointColors: ['38BDF8', '818CF8', '34D399'],
          index: 0,
          name: 'Traffic Sources',
          order: 0,
          values: [45, 25, 30],
        },
      ],
    };

    const xml = serializeChart(chart);
    expect(xml).toContain('<c:pieChart>');
    expect(xml).toContain('<c:dPt><c:idx val="0"/><c:spPr><a:solidFill><a:srgbClr val="38BDF8"/></a:solidFill></c:spPr></c:dPt>');
    expect(xml).toContain('<c:dPt><c:idx val="1"/><c:spPr><a:solidFill><a:srgbClr val="818CF8"/></a:solidFill></c:spPr></c:dPt>');
    expect(xml).toContain('<c:dPt><c:idx val="2"/><c:spPr><a:solidFill><a:srgbClr val="34D399"/></a:solidFill></c:spPr></c:dPt>');
  });

  it('serializes a doughnut chart with custom holeSize and dataLabels', () => {
    const chart: PptxChart = {
      categories: ['TypeScript', 'Rust', 'Go'],
      chartType: 'doughnut',
      dataLabels: { showPercent: true, showVal: true },
      holeSize: 65,
      series: [
        {
          index: 0,
          name: 'Language Share',
          order: 0,
          values: [70, 20, 10],
        },
      ],
    };

    const xml = serializeChart(chart);
    expect(xml).toContain('<c:doughnutChart>');
    expect(xml).toContain('<c:holeSize val="65"/>');
    expect(xml).toContain('<c:dLbls>');
    expect(xml).toContain('<c:showPercent val="1"/>');
    expect(xml).not.toContain('<c:catAx>');
  });

  it('serializes an area chart and radar chart', () => {
    const areaChart: PptxChart = {
      categories: ['W1', 'W2', 'W3'],
      chartType: 'area',
      series: [
        {
          fill: { solidColor: { type: 'srgb', value: '6366F1' }, type: 'solid' },
          index: 0,
          name: 'Bandwidth',
          order: 0,
          values: [100, 250, 400],
        },
      ],
    };
    const areaXml = serializeChart(areaChart);
    expect(areaXml).toContain('<c:areaChart>');
    expect(areaXml).toContain('<c:catAx>');

    const radarChart: PptxChart = {
      categories: ['Speed', 'Fidelity', 'Footprint', 'Type Safety'],
      chartType: 'radar',
      series: [
        {
          fill: { solidColor: { type: 'srgb', value: '10B981' }, type: 'solid' },
          index: 0,
          name: 'Benchmark Score',
          order: 0,
          values: [98, 99, 95, 100],
        },
      ],
    };
    const radarXml = serializeChart(radarChart);
    expect(radarXml).toContain('<c:radarChart>');
  });
});

describe('Chart Serializer scatter chart, titles, and legend variants', () => {
  it('serializes scatter chart and custom title', () => {
    const chart: PptxChart = {
      categories: ['1', '2', '3'],
      chartType: 'scatter',
      title: 'Performance Benchmark',
      legend: { position: 'left' },
      series: [
        {
          index: 0,
          name: 'Latency',
          order: 0,
          values: [10, 15, 8],
        },
      ],
    };
    const xml = serializeChart(chart);
    expect(xml).toContain('<c:scatterChart>');
    expect(xml).toContain('<c:title>');
    expect(xml).toContain('Performance Benchmark');
    expect(xml).toContain('<c:legendPos val="l"/>');
  });

  it('serializes top and topRight legend positions', () => {
    const chartTop: PptxChart = {
      categories: ['A'],
      chartType: 'bar',
      legend: { position: 'top' },
      series: [{ index: 0, name: 'S', order: 0, values: [1] }],
    };
    expect(serializeChart(chartTop)).toContain('<c:legendPos val="t"/>');

    const chartTR: PptxChart = {
      categories: ['A'],
      chartType: 'bar',
      legend: { position: 'topRight' },
      series: [{ index: 0, name: 'S', order: 0, values: [1] }],
    };
    expect(serializeChart(chartTR)).toContain('<c:legendPos val="tr"/>');
  });
});

describe('Chart Serializer gridlines, smooth lines, and stacked bars', () => {
  it('serializes value axis gridlines, smooth lines, and stacked bar groupings', () => {
    const chart: PptxChart = {
      categories: ['A', 'B'],
      chartType: 'barStacked',
      smooth: true,
      valAxis: { showGridlines: true, gridlineColor: 'E2E8F0' },
      catAxis: { showGridlines: true, gridlineColor: 'E2E8F0' },
      series: [
        { index: 0, name: 'S1', order: 0, values: [10, 20] },
      ],
    };
    const xml = serializeChart(chart);
    expect(xml).toContain('<c:grouping val="stacked"/>');
    expect(xml).toContain('<c:majorGridlines>');
    expect(xml).toContain('E2E8F0');

    const percentChart: PptxChart = {
      categories: ['A'],
      chartType: 'barPercentStacked',
      series: [{ index: 0, name: 'S1', order: 0, values: [100] }],
    };
    expect(serializeChart(percentChart)).toContain('<c:grouping val="percentStacked"/>');

    const smoothLine: PptxChart = {
      categories: ['1', '2'],
      chartType: 'line',
      smooth: true,
      series: [{ index: 0, name: 'L', order: 0, values: [1, 2] }],
    };
    expect(serializeChart(smoothLine)).toContain('<c:smooth val="1"/>');
  });
});

describe('Chart Serializer axis and legend text properties', () => {
  it('serializes custom color and font size on category axis and legend', () => {
    const chart: PptxChart = {
      categories: ['X'],
      chartType: 'bar',
      catAxis: { color: '#64748B', fontSize: 12, axisColor: '#CBD5E1' },
      legend: { color: '#475569', fontSize: 10, position: 'right' },
      series: [{ index: 0, name: 'S', order: 0, values: [5] }],
    };
    const xml = serializeChart(chart);
    expect(xml).toContain('val="64748B"');
    expect(xml).toContain('val="475569"');
    expect(xml).toContain('sz="1200"');
    expect(xml).toContain('sz="1000"');
  });
});
