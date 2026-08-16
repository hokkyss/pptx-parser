import { describe, expect, it } from 'vitest';
import { inches, points } from '@hokkyss/pptx-core';
import { Presentation } from '../../lib/presentation';

describe('Presentation SDK E2E Integration (Synthetic Multi-Slide Suite)', () => {
  it('builds a comprehensive 4-slide enterprise presentation from scratch and verifies full roundtrip', async () => {
    const pres = Presentation.create({
      author: 'Antigravity Test Suite',
      company: 'Open Source Cloud',
      title: 'Enterprise Architecture & Cloud Platform Overview',
      width: inches(13.333),
      height: inches(7.5),
    });

    // 🎨 Configure custom theme
    pres
      .setThemeName('Enterprise Modern Theme')
      .setThemeColors({
        accent1: '#0284C7', // Sky blue
        accent2: '#6366F1', // Indigo
        accent3: '#10B981', // Emerald
        accent4: '#F59E0B', // Amber
        accent5: '#EF4444', // Red
        accent6: '#8B5CF6', // Purple
        dk1: '#0F172A', // Slate 900
        lt1: '#FFFFFF',
        dk2: '#1E293B',
        lt2: '#F8FAFC',
        hlink: '#2563EB',
        folHlink: '#7C3AED',
      })
      .setThemeFonts({
        major: 'Inter',
        minor: 'Roboto',
        name: 'Enterprise Typography Scheme',
      });

    // ============================================================
    // SLIDE 1: Cover Slide
    // ============================================================
    const slide1 = pres.addSlide();
    slide1.setBackground('0F172A');

    // Title
    slide1.addText('Enterprise Cloud Architecture Deck', {
      align: 'left',
      bold: true,
      color: '38BDF8',
      fontSize: points(36),
      h: inches(1.5),
      w: inches(11),
      x: inches(1),
      y: inches(1.8),
    });

    // Subtitle
    slide1.addText('High-Performance Distributed Systems & Edge Compute', {
      align: 'left',
      color: '94A3B8',
      fontSize: points(20),
      h: inches(0.8),
      w: inches(11),
      x: inches(1),
      y: inches(3.3),
    });

    // Accent shape
    slide1.addShape('roundRect', {
      fill: '0284C7',
      h: inches(0.1),
      w: inches(3),
      x: inches(1),
      y: inches(4.3),
    });

    // Metadata footnote
    slide1.addText('Confidential • Automated E2E Test Suite • 2026', {
      color: '64748B',
      fontSize: points(12),
      h: inches(0.5),
      w: inches(10),
      x: inches(1),
      y: inches(6.2),
    });

    // ============================================================
    // SLIDE 2: Multilevel Hierarchy with All Text Modifiers
    // ============================================================
    const slide2 = pres.addSlide();
    slide2.setBackground('FFFFFF');

    slide2.addText('Core Architecture & Protocol Breakdown', {
      bold: true,
      color: '0F172A',
      fontSize: points(28),
      h: inches(0.8),
      w: inches(11.33),
      x: inches(1),
      y: inches(0.8),
    });

    slide2.addText([
      { level: 0, text: '1. Event-Driven Edge Runtime Engine' },
      { level: 1, text: '1.1 Zero-Cold-Start Microsecond Execution' },
      { level: 1, text: '1.2 Distributed Memory Caching Tier' },
      { level: 2, text: '1.2.1 Multi-Region Cache Synchronization' },
      { level: 3, text: '1.2.1.1 Consensus Protocol Specification' },
      // Modifiers:
      {
        italic: true,
        level: 4,
        text: '1.2.1.1.1 Sub-millisecond latency profile in all zones',
      },
      {
        level: 4,
        runs: [
          { strikethrough: true, text: '1.2' },
          { text: '.1.1.2 Legacy HTTP/1.1 polling fallback mode' },
        ],
      },
      {
        level: 5,
        runs: [
          { text: '1.2', underline: true },
          { text: '.1.1.2.1 Deprecated keep-alive connection pool' },
        ],
      },
      {
        level: 4,
        runs: [
          { text: '1.2.1.1.3 Algorithmic Time Complexity O(n' },
          { superscript: true, text: '2' },
          { text: ')' },
        ],
      },
      {
        level: 3,
        runs: [
          { text: '1.2.1.2 Molecular Thermal Model H' },
          { subscript: true, text: '2' },
          { text: 'O cooling system integration' },
        ],
      },
      { level: 0, text: '2. Observability & OpenTelemetry Metrics' },
      { level: 1, text: '2.1 Real-Time Distributed Tracing' },
    ], {
      h: inches(5.2),
      w: inches(11.33),
      x: inches(1),
      y: inches(1.8),
    });

    // ============================================================
    // SLIDE 3: Complex Multi-Column Financial & Benchmark Table
    // ============================================================
    const slide3 = pres.addSlide();
    slide3.setBackground('F8FAFC');

    slide3.addText('System Performance & Benchmark Metrics', {
      bold: true,
      color: '0F172A',
      fontSize: points(28),
      h: inches(0.8),
      w: inches(11.33),
      x: inches(1),
      y: inches(0.8),
    });

    slide3.addTable([
      ['Benchmark Metric', 'Legacy Baseline', 'Target SLA', 'Achieved Status'],
      ['P99 Query Latency', '142 ms', '< 50 ms', '34 ms (Pass)'],
      ['Peak Throughput', '8,500 RPS', '25,000 RPS', '32,100 RPS (Pass)'],
      ['Error Rate (HTTP 5xx)', '0.12%', '< 0.01%', '0.002% (Pass)'],
      ['Memory Overhead', '1.2 GB', '< 500 MB', '248 MB (Pass)'],
      ['Cold Start Time', '850 ms', '< 50 ms', '12 ms (Pass)'],
    ], {
      colWidths: [inches(3.5), inches(2.5), inches(2.5), inches(2.83)],
      h: inches(3.8),
      w: inches(11.33),
      x: inches(1),
      y: inches(1.8),
    });

    // ============================================================
    // SLIDE 4: Geometric Shape & Component Matrix
    // ============================================================
    const slide4 = pres.addSlide();
    slide4.setBackground('FFFFFF');

    slide4.addText('Microservice Component Architecture', {
      bold: true,
      color: '0F172A',
      fontSize: points(28),
      h: inches(0.8),
      w: inches(11.33),
      x: inches(1),
      y: inches(0.8),
    });

    const components = [
      { fill: '0284C7', name: 'API Gateway', x: inches(1), y: inches(2) },
      { fill: '6366F1', name: 'Auth Engine', x: inches(4), y: inches(2) },
      { fill: '10B981', name: 'RAG Pipeline', x: inches(7), y: inches(2) },
      { fill: '8B5CF6', name: 'Analytics Core', x: inches(10), y: inches(2) },
    ];

    for (const comp of components) {
      slide4.addShape('roundRect', {
        fill: comp.fill,
        h: inches(2),
        text: comp.name,
        w: inches(2.5),
        x: comp.x,
        y: comp.y,
      });
    }

    expect(pres.slides.length).toBe(4);

    // ============================================================
    // Full Roundtrip Serialization & Verification
    // ============================================================
    const buffer = await pres.toBuffer();
    expect(buffer).toBeInstanceOf(Uint8Array);
    expect(buffer.length).toBeGreaterThan(5000);

    const reloaded = await Presentation.load(buffer);
    expect(reloaded.slides.length).toBe(4);
    expect(reloaded.metadata.title).toBe('Enterprise Architecture & Cloud Platform Overview');
    expect(reloaded.metadata.creator).toBe('Antigravity Test Suite');

    // Verify Theme persistence
    const theme = reloaded.ast.themes[0];
    expect(theme.name).toBe('Enterprise Modern Theme');
    expect(theme.colorScheme.accent1).toBe('0284C7');
    expect(theme.colorScheme.accent2).toBe('6366F1');
    expect(theme.fontScheme.majorFont).toBe('Inter');
    expect(theme.fontScheme.minorFont).toBe('Roboto');
    expect(theme.fontScheme.name).toBe('Enterprise Typography Scheme');

    // Verify Slide 1 Elements
    const s1Elements = reloaded.getSlide(1)?.getElements();
    expect(s1Elements?.length).toBe(4);

    // Verify Slide 2 Multilevel runs & modifiers
    const s2Body = reloaded.getSlide(2)?.getElements().find((e) => e.elementType === 'shape' && e.textBody?.paragraphs && e.textBody.paragraphs.length > 5);
    expect(s2Body).toBeDefined();
    if (s2Body && s2Body.elementType === 'shape') {
      const paragraphs = s2Body.textBody?.paragraphs || [];
      expect(paragraphs.length).toBe(12);
      expect(paragraphs[0].properties.level).toBe(0);
      expect(paragraphs[1].properties.level).toBe(1);
      expect(paragraphs[5].runs[0].properties.italic).toBe(true);
      expect(paragraphs[6].runs[0].properties.strikethrough).toBeTruthy();
      expect(paragraphs[7].runs[0].properties.underline).toBeTruthy();
      expect(paragraphs[8].runs[1].properties.superscript).toBe(true);
      expect(paragraphs[9].runs[1].properties.subscript).toBe(true);
    }

    // Verify Slide 3 Table dimensions
    const s3Table = reloaded.getSlide(3)?.getElements().find((e) => e.elementType === 'table');
    expect(s3Table).toBeDefined();
    if (s3Table && s3Table.elementType === 'table') {
      expect(s3Table.table?.rows.length).toBe(6);
      expect(s3Table.table?.columnWidths.length).toBe(4);
    }

    // Verify Slide 4 Shapes
    const s4Shapes = reloaded.getSlide(4)?.getElements().filter((e) => e.elementType === 'shape');
    expect(s4Shapes?.length).toBe(5); // 1 title + 4 component boxes
  });

  it('performs slide CRUD mutations (add, duplicate, move, remove) cleanly on multi-slide presentation', async () => {
    const pres = Presentation.create({ title: 'CRUD Deck' });

    const s1 = pres.addSlide();
    s1.addText('Slide 1 Content');

    const s2 = pres.addSlide();
    s2.addText('Slide 2 Content');

    const s3 = pres.addSlide();
    s3.addText('Slide 3 Content');

    expect(pres.slides.length).toBe(3);

    // Duplicate Slide 2 -> now 4 slides
    const duplicated = pres.duplicateSlide(2);
    expect(pres.slides.length).toBe(4);
    expect(duplicated.slideNumber).toBe(4);

    // Move duplicated slide to position 2
    pres.moveSlide(4, 2);
    expect(pres.slides[1]).toBe(duplicated);
    expect(pres.slides[1].slideNumber).toBe(2);

    // Remove Slide 1
    const removed = pres.removeSlide(1);
    expect(removed).toBe(true);
    expect(pres.slides.length).toBe(3);
    expect(pres.slides[0].slideNumber).toBe(1);
    expect(pres.slides[1].slideNumber).toBe(2);
    expect(pres.slides[2].slideNumber).toBe(3);

    // Roundtrip verification
    const buffer = await pres.toBuffer();
    const verified = await Presentation.load(buffer);
    expect(verified.slides.length).toBe(3);
  });
});
