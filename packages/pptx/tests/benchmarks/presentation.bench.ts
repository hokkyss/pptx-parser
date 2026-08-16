import { bench, describe } from 'vitest';
import { inches, points } from '@hokkyss/pptx-core';
import { Presentation } from '../../lib/presentation';

describe('Presentation Generation Benchmarks (Vitest)', () => {
  bench('Single Slide with Shape & Text', async () => {
    const pres = Presentation.create();
    const slide = pres.addSlide();
    slide.setBackground('0F172A');
    slide.addText('Vitest Benchmark', { bold: true, fontSize: points(24) });
    slide.addShape('roundRect', { fill: '0284C7', h: inches(2), w: inches(4), x: inches(1), y: inches(2) });
    await pres.toBuffer();
  });

  bench('Enterprise Data Table (10 rows x 5 cols)', async () => {
    const pres = Presentation.create();
    const slide = pres.addSlide();
    slide.addTable({
      columns: [{ w: inches(2) }, { w: inches(2) }, { w: inches(2) }, { w: inches(2) }, { w: inches(2) }],
      data: [
        ['Metric', 'Target', 'Actual', 'Variance', 'Status'],
        ...Array.from({ length: 9 }, (_, i) => [
          `Service ${i + 1}`,
          '99.9%',
          '99.95%',
          '+0.05%',
          'OPTIMAL',
        ]),
      ],
      header: { fill: '0284C7', textOptions: { bold: true, color: 'FFFFFF' } },
      x: inches(1),
      y: inches(1),
    });
    await pres.toBuffer();
  });

  bench('DrawingML Column Chart Generation', async () => {
    const pres = Presentation.create();
    const slide = pres.addSlide();
    slide.addChart({
      categories: ['Q1', 'Q2', 'Q3', 'Q4'],
      chartType: 'column',
      series: [
        { color: '0284C7', name: 'Revenue', values: [100, 150, 200, 280] },
        { color: '10B981', name: 'Profit', values: [30, 45, 70, 95] },
      ],
      showGridlines: true,
    });
    await pres.toBuffer();
  });

  bench('10-Slide Full Enterprise Deck', async () => {
    const pres = Presentation.create();
    for (let s = 0; s < 10; s++) {
      const slide = pres.addSlide();
      slide.setBackground(s % 2 === 0 ? '0F172A' : 'F8FAFC');
      slide.addText(`Enterprise Slide #${s + 1}`, { bold: true, fontSize: points(20), x: inches(1), y: inches(0.8) });
      slide.addShape('roundRect', { fill: '1E293B', h: inches(3), w: inches(5), x: inches(1), y: inches(1.5) });
      if (s % 3 === 0) {
        slide.addChart({
          categories: ['A', 'B', 'C', 'D'],
          chartType: 'column',
          series: [{ color: '38BDF8', name: 'Ops', values: [10, 20, 30, 40] }],
        });
      }
      slide.setNotes(`Speaker notes for slide ${s + 1}`);
    }
    await pres.toBuffer();
  });

  bench('50-Slide Batch Scale', async () => {
    const pres = Presentation.create();
    for (let s = 0; s < 50; s++) {
      const slide = pres.addSlide();
      slide.addText(`Slide ${s + 1}`, { fontSize: points(16), x: inches(0.5), y: inches(0.5) });
      slide.addShape('rect', { fill: '0284C7', h: inches(2), w: inches(3), x: inches(1), y: inches(1) });
    }
    await pres.toBuffer();
  });
});
