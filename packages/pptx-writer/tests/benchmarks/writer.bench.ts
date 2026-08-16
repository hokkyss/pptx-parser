import { bench, describe } from 'vitest';
import type {
  PptxChart,
  PptxDocument,
  PptxShapeElement,
  PptxTable,
  PptxTableElement,
} from '@hokkyss/pptx-core';
import {
  emu,
  emuDegree,
  hundredthsPoint,
} from '@hokkyss/pptx-core';
import { serializeChart } from '../../lib/serializers/chart-serializer';
import { serializeShape } from '../../lib/serializers/shape-serializer';
import { serializeTable } from '../../lib/serializers/table-serializer';
import { writePptx } from '../../lib/writer';

describe('PPTX Writer Serializer Benchmarks (Vitest)', () => {
  const shapeEl: PptxShapeElement = {
    elementType: 'shape',
    fill: { solidColor: { type: 'srgb', value: '0284C7' }, type: 'solid' },
    geometry: { presetGeometry: 'roundRect' },
    id: '2',
    isVisible: true,
    name: 'Shape 1',
    position: { cx: emu(2000000), cy: emu(1000000), x: emu(1000000), y: emu(1000000) },
    rotation: emuDegree(0),
    textBody: {
      bodyProperties: {},
      paragraphs: [
        {
          properties: {},
          runs: [{ properties: { fontSize: hundredthsPoint(1400) }, text: 'Benchmark Shape' }],
        },
      ],
    },
    type: 'shape',
    zIndex: 0,
  };

  const table: PptxTable = {
    columnWidths: [emu(1000000), emu(1000000)],
    rows: [
      {
        cells: [
          {
            properties: { fill: { solidColor: { type: 'srgb', value: '0284C7' }, type: 'solid' } },
            textBody: { bodyProperties: {}, paragraphs: [{ properties: {}, runs: [{ properties: {}, text: 'A1' }] }] },
          },
          {
            properties: { fill: { solidColor: { type: 'srgb', value: '0284C7' }, type: 'solid' } },
            textBody: { bodyProperties: {}, paragraphs: [{ properties: {}, runs: [{ properties: {}, text: 'B1' }] }] },
          },
        ],
        height: emu(500000),
      },
      {
        cells: [
          { textBody: { bodyProperties: {}, paragraphs: [{ properties: {}, runs: [{ properties: {}, text: 'A2' }] }] } },
          { textBody: { bodyProperties: {}, paragraphs: [{ properties: {}, runs: [{ properties: {}, text: 'B2' }] }] } },
        ],
        height: emu(500000),
      },
    ],
  };

  const tableEl: PptxTableElement = {
    elementType: 'table',
    id: '3',
    isVisible: true,
    name: 'Table 1',
    position: { cx: emu(2000000), cy: emu(1000000), x: emu(1000000), y: emu(1000000) },
    rotation: emuDegree(0),
    table,
    type: 'graphicFrame',
    zIndex: 1,
  };

  const chart: PptxChart = {
    categories: ['Q1', 'Q2', 'Q3', 'Q4'],
    chartType: 'column',
    series: [{ index: 0, name: 'Revenue', order: 0, values: [100, 200, 300, 400] }],
  };

  bench('Serialize Shape to DrawingML XML', () => {
    serializeShape(shapeEl);
  });

  bench('Serialize Table to DrawingML XML', () => {
    serializeTable(tableEl);
  });

  bench('Serialize Chart to DrawingML XML', () => {
    serializeChart(chart);
  });

  const minimalDoc: PptxDocument = {
    customXml: [],
    media: [],
    metadata: {
      created: new Date(),
      modified: new Date(),
      revision: 1,
      slideCount: 1,
      slideHeight: emu(6858000),
      slideWidth: emu(12192000),
    },
    slideLayouts: [],
    slideMasters: [],
    slides: [
      {
        animations: [],
        elements: [shapeEl, tableEl],
        shapes: [],
        slideId: 'slide1',
        slideNumber: 1,
      },
    ],
    themes: [],
  };

  bench('Full writePptx ZIP Assembly', async () => {
    await writePptx(minimalDoc);
  });
});
