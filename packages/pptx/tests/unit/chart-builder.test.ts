import { describe, expect, it } from 'vitest';
import { inches, points } from '@hokkyss/pptx-core';
import { Presentation } from '../../lib/presentation';

describe('ChartBuilder (Unit Tests)', () => {
  it('adds a column bar chart to slide and builds graphicFrame element', () => {
    const pres = Presentation.create();
    const slide = pres.addSlide();

    slide.addChart({
      categories: ['Workers V8', 'Node.js', 'Deno'],
      chartType: 'column',
      h: inches(4.5),
      legend: { position: 'right' },
      series: [
        {
          color: '0284C7',
          name: 'Parse Time (ms)',
          values: [14.2, 28.5, 19.8],
        },
        {
          color: '6366F1',
          name: 'Serialization (ms)',
          values: [8.1, 16.2, 11.4],
        },
      ],
      w: inches(9.0),
      x: inches(1.0),
      y: inches(1.5),
    });

    const elements = slide.getElements();
    expect(elements.length).toBe(1);

    const chartEl = elements[0];
    expect(chartEl.elementType).toBe('chart');
    expect(chartEl.type).toBe('graphicFrame');
    expect(chartEl.chart).toBeDefined();
    expect(chartEl.chart?.categories).toEqual(['Workers V8', 'Node.js', 'Deno']);
    expect(chartEl.chart?.series.length).toBe(2);
    expect(chartEl.chart?.series[0].name).toBe('Parse Time (ms)');
    expect(chartEl.chart?.series[0].values).toEqual([14.2, 28.5, 19.8]);
  });

  it('adds a doughnut chart with dataLabels to slide', () => {
    const pres = Presentation.create();
    const slide = pres.addSlide();

    slide.addChart({
      categories: ['Core', 'Reader', 'Writer', 'Fluent'],
      chartType: 'doughnut',
      dataLabels: { showPercent: true, showVal: false },
      h: inches(4.0),
      holeSize: 60,
      series: [
        {
          color: '0284C7',
          name: 'Codebase Lines',
          values: [1200, 3400, 3100, 2200],
        },
      ],
      w: inches(5.0),
      x: inches(1.0),
      y: inches(1.0),
    });

    const chartEl = slide.getElements()[0];
    expect(chartEl.elementType).toBe('chart');
    expect(chartEl.chart?.chartType).toBe('doughnut');
    expect(chartEl.chart?.holeSize).toBe(60);
    expect(chartEl.chart?.dataLabels?.showPercent).toBe(true);
    // When no slice colors are specified on radial chart, fill is omitted for theme inheritance
    expect(chartEl.chart?.series[0].fill).toBeUndefined();
    expect(chartEl.chart?.series[0].dataPointColors).toBeUndefined();
  });

  it('supports explicit colors array for pie chart slices', () => {
    const pres = Presentation.create();
    const slide = pres.addSlide();

    slide.addChart({
      categories: ['React', 'Vue', 'Svelte', 'Solid'],
      chartType: 'pie',
      colors: ['0284C7', '10B981', 'F59E0B', '8B5CF6'],
      series: [
        {
          name: 'Frameworks',
          values: [50, 25, 15, 10],
        },
      ],
    });

    const chartEl = slide.getElements()[0];
    expect(chartEl.chart?.series[0].dataPointColors).toEqual(['0284C7', '10B981', 'F59E0B', '8B5CF6']);
  });
});

it('supports series fill object and axis customizations (catAxis, valAxis, textColor, gridColor)', () => {
  const pres = Presentation.create();
  const slide = pres.addSlide();

  slide.addChart({
    chartType: 'line',
    categories: ['Jan', 'Feb'],
    textColor: '333333',
    axisColor: '666666',
    gridColor: 'CCCCCC',
    showGridlines: true,
    series: [
      {
        name: 'Series 1',
        values: [10, 20],
        fill: { type: 'solid', solidColor: { type: 'srgb', value: '112233' } },
      },
    ],
  });

  const chartEl = slide.getElements()[0];
  expect(chartEl.chart?.series[0].fill).toBeDefined();
  expect(chartEl.chart?.legend?.color).toBe('333333');
  expect(chartEl.chart?.catAxis?.axisColor).toBe('666666');
  expect(chartEl.chart?.valAxis?.gridlineColor).toBe('CCCCCC');
  expect(chartEl.chart?.valAxis?.showGridlines).toBe(true);
});

describe('ChartBuilder edge cases', () => {
  it('covers optional axes, empty series options and series without names', () => {
    const pres = Presentation.create();
    const slide = pres.addSlide();

    slide.addChart({
      chartType: 'column',
      series: [{ values: [10, 20] }],
    });
    const chart1 = slide.getElements()[0];
    expect(chart1.chart?.series[0].name).toBe('Series 1');
    expect(chart1.chart?.categories).toEqual([]);

    slide.addChart({
      catAxis: {
        axisColor: 'FF0000',
        color: '333333',
        fontSize: points(12),
        gridlineColor: 'CCCCCC',
        showGridlines: true,
      },
      valAxis: {
        axisColor: '00FF00',
        color: '666666',
        fontSize: points(14),
        gridlineColor: 'E5E5E5',
        showGridlines: false,
      },
      series: [],
    });
    const chart2 = slide.getElements()[1];
    expect(chart2.chart?.catAxis?.axisColor).toBe('FF0000');
    expect(chart2.chart?.valAxis?.showGridlines).toBe(false);

    // Empty series options
    slide.addChart({});
    const chart3 = slide.getElements()[2];
    expect(chart3.chart?.series).toEqual([]);
  });
});
