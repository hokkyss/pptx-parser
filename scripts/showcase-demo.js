import { readFileSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';
import {
  degrees,
  inches,
  points,
  Presentation,
} from '../packages/pptx/dist/index.js';

async function runShowcase() {
  console.log('==================================================================');
  console.log('       ✨ ULTIMATE SHOWCASE: @hokkyss/pptx CAPABILITIES SUITE');
  console.log('==================================================================\n');

  const t0 = performance.now();

  // 1. Initialize Presentation (16:9 Widescreen)
  const pres = Presentation.create({
    author: 'Hokkyss Core Team',
    company: 'Hokkyss Open-Source Ecosystem',
    title: 'Hokkyss PPTX — The Modern Presentation Engine',
  });

  // 2. Configure Global Presentation Theme
  pres
    .setThemeName('Enterprise Cobalt Theme')
    .setThemeColors({
      accent1: '#0284C7', // Sky 600
      accent2: '#6366F1', // Indigo 500
      accent3: '#10B981', // Emerald 500
      accent4: '#F59E0B', // Amber 500
      accent5: '#EF4444', // Red 500
      accent6: '#8B5CF6', // Purple 500
      dk1: '#0F172A',     // Slate 900
      dk2: '#1E293B',     // Slate 800
      lt1: '#FFFFFF',     // Pure White
      lt2: '#F8FAFC',     // Slate 50
    })
    .setThemeFonts({
      major: 'Inter',
      minor: 'Roboto',
      name: 'Modern Clean Sans',
    });

  const addDisclaimer = (slide, isDark = false) => {
    slide.addText(
      '* Disclaimer: All benchmark metrics, performance figures, and simulated architectures in this deck are synthetic demonstration values generated solely to showcase the authoring, formatting, and rendering capabilities of @hokkyss/pptx.',
      {
        align: 'center',
        color: isDark ? '64748B' : '94A3B8',
        fontSize: points(8),
        h: inches(0.25),
        italic: true,
        w: inches(11.73),
        x: inches(0.8),
        y: inches(7.12),
      },
    );
  };

  // ==================================================================
  // SLIDE 1: Executive Cover Slide with Metric Cards & Rich Notes
  // ==================================================================
  console.log('🎨 Generating Slide 1: Executive Dark Theme Cover with Shadows...');
  const slide1 = pres.addSlide();
  slide1.setTransition('fade', { durationMs: 600, throughBlack: true });
  slide1.setBackground('0F172A');
  slide1.setNotes([
    {
      runs: [
        { text: 'Opening Keynote: ', bold: true },
        { text: 'Modern OpenXML PPTX Parsing & Compilation Engine', bold: true, underline: true },
      ],
    },
    {
      runs: [
        { text: 'Key architectural pillars for the presenter:' },
      ],
    },
    {
      bullet: true,
      level: 0,
      runs: [
        { text: 'Emphasize ', italic: true },
        { text: '100% isomorphic execution', bold: true, underline: true },
        { text: ' across modern edge runtimes.' },
      ],
    },
    {
      bullet: true,
      level: 1,
      runs: [{ text: 'Cloudflare Workers (V8 isolates with zero Node.js native bindings)' }],
    },
    {
      bullet: true,
      level: 1,
      runs: [{ text: 'Browser environments (client-side export & live preview)' }],
    },
    {
      bullet: true,
      level: 0,
      runs: [
        { text: 'Zero native dependencies: ' },
        { text: 'no python-pptx, no libreoffice', strikethrough: true },
        { text: ' — 100% pure TypeScript codecs.' },
      ],
    },
    {
      bullet: 'number',
      level: 0,
      runs: [{ text: 'Deep OpenXML fidelity & layer composition', bold: true }],
    },
    {
      bullet: 'number',
      level: 0,
      runs: [{ text: 'Strongly typed units (Inches, Points, Degrees, EMU)', bold: true }],
    },
  ]);

  // Ambient Header Accent Shape with Shadow
  slide1.addShape('roundRect', {
    fill: '1E293B',
    h: inches(0.5),
    line: { color: '38BDF8', width: inches(0.015) },
    shadow: {
      blur: inches(0.2),
      color: '38BDF8',
      direction: degrees(90),
      distance: inches(0.02),
      opacity: 0.35,
    },
    text: '⚡ NEXT-GENERATION ISOMORPHIC PRESENTATION COMPILER',
    textOptions: {
      bold: true,
      color: '38BDF8',
      fontSize: points(10),
      align: 'center',
    },
    w: inches(4.8),
    x: inches(4.26),
    y: inches(1.5),
  });

  // Main Hero Title
  slide1.addText('@hokkyss/pptx', {
    align: 'center',
    bold: true,
    color: 'FFFFFF',
    font: '+mj-lt',
    fontSize: points(44),
    h: inches(1.0),
    hyperlink: { tooltip: 'GitHub Repository: @hokkyss/pptx', url: 'https://github.com/hokkyss/pptx-parser' },
    w: inches(11.33),
    x: inches(1.0),
    y: inches(2.3),
  });

  // Subtitle
  slide1.addText('Isomorphic OpenXML Presentation Parser, Serializer & Fluent Authoring SDK', {
    color: '94A3B8',
    fontSize: points(16),
    h: inches(0.6),
    w: inches(11.33),
    x: inches(1.0),
    y: inches(3.3),
    align: 'center',
  });

  // 3 Metric Badges
  const badges = [
    { label: '🚀 100% Isomorphic', desc: 'Node, Browser, Cloudflare, Deno & Bun', x: 1.0 },
    { label: '🌲 Zero-Dependency', desc: 'Pure TypeScript codecs, zero native binaries', x: 4.8 },
    { label: '🛡️ Type-Safe Units', desc: 'Branded EMU, Points, Inches & Degrees', x: 8.6 },
  ];

  for (const b of badges) {
    slide1.addShape('roundRect', {
      fill: '1E293B',
      h: inches(1.8),
      line: { color: '334155', width: inches(0.01) },
      shadow: {
        blur: inches(0.15),
        color: '000000',
        direction: degrees(90),
        distance: inches(0.05),
        opacity: 0.35,
      },
      text: [
        { text: `${b.label}\n\n`, bold: true, color: '38BDF8', fontSize: points(14) },
        { text: b.desc, color: 'CBD5E1', fontSize: points(12) },
      ],
      w: inches(3.6),
      x: inches(b.x),
      y: inches(4.8),
    });
  }

  addDisclaimer(slide1, true);

  // ==================================================================
  // SLIDE 2: Modular Architecture & Shape Gallery
  // -------------------------------------------------------------
  console.log('📐 Generating Slide 2: Modular Architecture & Geometry Gallery...');
  const slide2 = pres.addSlide();
  slide2.setTransition('wipe', { direction: 'right', speed: 'fast' });
  slide2.setBackground('F8FAFC');

  slide2.addText('Modular Architecture & Shape Presets', {
    bold: true,
    color: '0F172A',
    font: '+mj-lt',
    fontSize: points(24),
    h: inches(0.6),
    w: inches(11.33),
    x: inches(1),
    y: inches(0.8),
  });

  slide2.addText('Four decoupled packages paired with DrawingML vector presets, outer drop shadows, and border outlines', {
    color: '64748B',
    fontSize: points(14),
    h: inches(0.4),
    w: inches(11.33),
    x: inches(1),
    y: inches(1.4),
  });

  // Package Card 1: Core
  slide2.addShape('roundRect', {
    fill: 'FFFFFF',
    h: inches(2.0),
    hyperlink: { tooltip: 'View @hokkyss/pptx-core on GitHub', url: 'https://github.com/hokkyss/pptx-parser/tree/main/packages/pptx-core' },
    line: { color: '0284C7', width: inches(0.02) },
    shadow: { blur: inches(0.15), color: '0284C7', direction: degrees(90), distance: inches(0.06), opacity: 0.25 },
    text: [
      { text: '@hokkyss/pptx-core\n', bold: true, color: '0284C7', fontSize: points(13) },
      { text: 'Universal AST definitions, branded unit types, and ECMA-376 OpenXML schema interfaces', color: '475569', fontSize: points(10) },
    ],
    w: inches(2.6),
    x: inches(1.0),
    y: inches(2.1),
  });

  // Package Card 2: Reader
  slide2.addShape('roundRect', {
    fill: '6366F1',
    h: inches(2.0),
    hyperlink: { tooltip: 'View @hokkyss/pptx-reader on GitHub', url: 'https://github.com/hokkyss/pptx-parser/tree/main/packages/pptx-reader' },
    shadow: { blur: inches(0.12), color: '000000', direction: degrees(90), distance: inches(0.04), opacity: 0.2 },
    text: [
      { text: '@hokkyss/pptx-reader\n', bold: true, color: 'FFFFFF', fontSize: points(13) },
      { text: 'Zero-copy streaming ZIP decompression, fast XML DOM parsing, and master cascade', color: 'E0E7FF', fontSize: points(10) },
    ],
    w: inches(2.6),
    x: inches(3.9),
    y: inches(2.1),
  });

  // Package Card 3: Writer
  slide2.addShape('roundRect', {
    fill: '10B981',
    h: inches(2.0),
    hyperlink: { tooltip: 'View @hokkyss/pptx-writer on GitHub', url: 'https://github.com/hokkyss/pptx-parser/tree/main/packages/pptx-writer' },
    shadow: { blur: inches(0.1), color: '000000', direction: degrees(90), distance: inches(0.04), opacity: 0.2 },
    text: [
      { text: '@hokkyss/pptx-writer\n', bold: true, color: 'FFFFFF', fontSize: points(13) },
      { text: 'Deterministic XML serializer, DrawingML emitter, and bit-exact package bundler', color: 'D1FAE5', fontSize: points(10) },
    ],
    w: inches(2.6),
    x: inches(6.8),
    y: inches(2.1),
  });

  // Package Card 4: SDK
  slide2.addShape('roundRect', {
    fill: 'F59E0B',
    h: inches(2.0),
    hyperlink: { tooltip: 'View @hokkyss/pptx on GitHub', url: 'https://github.com/hokkyss/pptx-parser/tree/main/packages/pptx' },
    shadow: { blur: inches(0.2), color: 'D97706', direction: degrees(90), distance: inches(0.05), opacity: 0.35 },
    text: [
      { text: '@hokkyss/pptx\n', bold: true, color: 'FFFFFF', fontSize: points(13) },
      { text: 'Unified high-level fluent authoring SDK with table builders and shape helpers', color: 'FEF3C7', fontSize: points(10) },
    ],
    w: inches(2.6),
    x: inches(9.7),
    y: inches(2.1),
  });

  // Row 2 Geometric Presets
  // 5. Right Arrow
  slide2.addShape('rightArrow', {
    fill: '0284C7',
    h: inches(0.8),
    text: 'PIPELINE',
    textOptions: { bold: true, color: 'FFFFFF', fontSize: points(10), align: 'center' },
    w: inches(2.6),
    x: inches(1.0),
    y: inches(4.6),
  });

  // 6. Diamond
  slide2.addShape('diamond', {
    fill: 'F1F5F9',
    h: inches(1.5),
    line: { color: '6366F1', width: inches(0.015) },
    text: 'DECISION',
    textOptions: { bold: true, color: '6366F1', fontSize: points(9), align: 'center' },
    w: inches(2.6),
    x: inches(3.9),
    y: inches(4.4),
  });

  // 7. Cloud Callout
  slide2.addShape('wedgeRoundRect', {
    fill: 'F8FAFC',
    h: inches(1.5),
    line: { color: '0284C7', dashStyle: 'dash', width: inches(0.015) },
    text: 'Cloud Native Isolate',
    textOptions: { color: '0284C7', fontSize: points(11), align: 'center' },
    w: inches(2.6),
    x: inches(6.8),
    y: inches(4.4),
  });

  // 8. Cylinder / Storage
  slide2.addShape('cylinder', {
    fill: 'F1F5F9',
    h: inches(1.5),
    line: { color: '10B981', width: inches(0.015) },
    text: 'R2 Blob Store',
    textOptions: { bold: true, color: '10B981', fontSize: points(11), align: 'center' },
    w: inches(2.6),
    x: inches(9.7),
    y: inches(4.4),
  });

  addDisclaimer(slide2, false);

  // ==================================================================
  // SLIDE 3: Multi-Style Enterprise Data Table
  // -------------------------------------------------------------
  console.log('📊 Generating Slide 3: Multi-Style Enterprise Data Table...');
  const slide3 = pres.addSlide();
  slide3.setTransition('push', { direction: 'up', speed: 'med' });
  slide3.setBackground('FFFFFF');

  slide3.addText('Runtime Performance & Memory Benchmarks', {
    bold: true,
    color: '0F172A',
    font: '+mj-lt',
    fontSize: points(24),
    h: inches(0.6),
    w: inches(11.33),
    x: inches(1.0),
    y: inches(0.8),
  });

  slide3.addText('100-slide roundtrip parse & compilation benchmarks across modern JavaScript engines', {
    color: '64748B',
    fontSize: points(14),
    h: inches(0.4),
    w: inches(11.33),
    x: inches(1),
    y: inches(1.4),
  });

  slide3.addTable(
    (tbl) => {
      // Header Row
      tbl.addRow({
        cells: [
          { bold: true, color: 'FFFFFF', fill: '0F172A', text: 'Target Runtime / Engine' },
          { bold: true, color: 'FFFFFF', fill: '0F172A', text: 'Parse Time (100 Slides)', align: 'right' },
          { bold: true, color: 'FFFFFF', fill: '0F172A', text: 'Compile Time (100 Slides)', align: 'right' },
          { bold: true, color: 'FFFFFF', fill: '0F172A', text: 'Memory Ceiling', align: 'right' },
          { bold: true, color: 'FFFFFF', fill: '0F172A', text: 'Edge Compatibility', align: 'center' },
        ],
        h: inches(0.55),
      });

      // Row 1: Node.js
      tbl.addRow({
        cells: [
          { bold: true, color: '0F172A', fill: 'F8FAFC', text: 'Node.js 22 LTS (V8)' },
          { color: '0284C7', fill: 'F8FAFC', text: '14.2 ms', align: 'right', bold: true },
          { color: '10B981', fill: 'F8FAFC', text: '18.5 ms', align: 'right', bold: true },
          { color: '475569', fill: 'F8FAFC', text: '< 12 MB', align: 'right' },
          { bold: true, color: '10B981', fill: 'ECFDF5', text: '✅ Native Fast', align: 'center' },
        ],
        h: inches(0.48),
      });

      // Row 2: Cloudflare Workers
      tbl.addRow({
        cells: [
          { bold: true, color: '0F172A', fill: 'FFFFFF', text: 'Cloudflare Workers (Edge)' },
          { color: '0284C7', fill: 'FFFFFF', text: '16.8 ms', align: 'right', bold: true },
          { color: '10B981', fill: 'FFFFFF', text: '21.0 ms', align: 'right', bold: true },
          { color: '475569', fill: 'FFFFFF', text: '< 8 MB', align: 'right' },
          { bold: true, color: '10B981', fill: 'ECFDF5', text: '✅ Native Isolate', align: 'center' },
        ],
        h: inches(0.48),
      });

      // Row 3: Web Browsers
      tbl.addRow({
        cells: [
          { bold: true, color: '0F172A', fill: 'F8FAFC', text: 'Web Browser (Chromium / Safari)' },
          { color: '0284C7', fill: 'F8FAFC', text: '19.4 ms', align: 'right' },
          { color: '10B981', fill: 'F8FAFC', text: '24.5 ms', align: 'right' },
          { color: '475569', fill: 'F8FAFC', text: '< 15 MB', align: 'right' },
          { bold: true, color: '10B981', fill: 'ECFDF5', text: '✅ Client-Side Export', align: 'center' },
        ],
        h: inches(0.48),
      });

      // Row 4: Bun & Deno
      tbl.addRow({
        cells: [
          { bold: true, color: '0F172A', fill: 'FFFFFF', text: 'Bun 1.2+ / Deno 2+' },
          { color: '0284C7', fill: 'FFFFFF', text: '11.5 ms', align: 'right', bold: true },
          { color: '10B981', fill: 'FFFFFF', text: '15.2 ms', align: 'right', bold: true },
          { color: '475569', fill: 'FFFFFF', text: '< 10 MB', align: 'right' },
          { bold: true, color: '10B981', fill: 'ECFDF5', text: '✅ Full Support', align: 'center' },
        ],
        h: inches(0.48),
      });

      // Row 5: Legacy Native (python-pptx)
      tbl.addRow({
        cells: [
          { bold: true, color: '94A3B8', fill: 'F8FAFC', text: 'Legacy python-pptx / COM' },
          { color: 'EF4444', fill: 'F8FAFC', text: '340.0 ms', align: 'right' },
          { color: 'EF4444', fill: 'F8FAFC', text: '480.0 ms', align: 'right' },
          { color: 'EF4444', fill: 'F8FAFC', text: '> 180 MB', align: 'right' },
          { bold: true, color: 'EF4444', fill: 'FEF2F2', text: '❌ Server Native Only', align: 'center' },
        ],
        h: inches(0.48),
      });
    },
    {
      colWidths: [inches(3.5), inches(2.0), inches(2.0), inches(1.8), inches(2.03)],
      header: {
        fill: '0284C7',
        color: 'FFFFFF',
        fontSize: points(12),
      },
      w: inches(11.33),
      x: inches(1.0),
      y: inches(2.1),
    },
  );

  addDisclaimer(slide3, false);

  // ==================================================================
  // SLIDE 4: Type-Safe Units & Formula Matrix
  // ==================================================================
  console.log('📑 Generating Slide 4: Deep Multilevel Hierarchies & Inline Modifiers...');
  const slide4 = pres.addSlide();
  slide4.setTransition('split', { direction: 'in', speed: 'fast' });
  slide4.setBackground('0F172A'); // Dark theme

  slide4.addText('Type-Safe Nominal Units & Coordinate System', {
    bold: true,
    color: '38BDF8',
    font: '+mj-lt',
    fontSize: points(24),
    h: inches(0.6),
    w: inches(11.33),
    x: inches(1),
    y: inches(0.8),
  });

  slide4.addText([
    { level: 0, text: '1. Nominal Units & Compile-Time Safety' },
    { level: 1, text: '1.1 Zero-runtime-overhead branded types for dimensional coordinates' },
    {
      level: 2,
      runs: [
        { text: '1.1.1 English Metric Units (EMU): ' },
        { text: '1 inch = 914,400 EMU', bold: true, color: '38BDF8' },
        { text: ' — ' },
        { text: '1 pt = 12,700 EMU', bold: true, color: '10B981' },
      ],
    },
    {
      level: 2,
      runs: [
        { text: '1.1.2 Rotational Angles: ' },
        { text: 'degrees(θ)', bold: true, color: 'F59E0B' },
        { text: ' compiles to ' },
        { text: 'θ × 60,000 EmuDegrees', bold: true, color: 'F59E0B' },
      ],
    },
    {
      level: 2,
      runs: [
        { text: '1.1.3 Maximum memory ceiling: ' },
        { text: '2', bold: true, color: 'F59E0B' },
        { text: '32', superscript: true, bold: true, color: 'F59E0B' },
        { text: ' bytes (4 GB buffer limit)' },
      ],
    },
    { level: 0, text: '2. High-Performance Text Run Modifiers' },
    { level: 1, text: '2.1 Rich inline text runs with mixed typographic styles' },
    {
      level: 2,
      runs: [
        { text: 'Formatting: ' },
        { text: 'Bold', bold: true },
        { text: ' | ' },
        { text: 'Italic', italic: true },
        { text: ' | ' },
        { text: 'Underlined', underline: true },
        { text: ' | ' },
        { text: 'Strikethrough', strikethrough: true },
        { text: ' | ' },
        { text: 'Colored #6366F1', color: '6366F1', bold: true },
      ],
    },
  ], {
    color: 'E2E8F0',
    fontSize: points(12),
    h: inches(4.5),
    w: inches(11.33),
    x: inches(1),
    y: inches(1.6),
  });

  addDisclaimer(slide4, true);

  // ==================================================================
  // SLIDE 5: Decoupled Three-Tier Architecture Cards
  // ==================================================================
  console.log('🏗️ Generating Slide 5: Floating Microservice Architecture Cards...');
  const slide5 = pres.addSlide();
  slide5.setTransition('cover', { direction: 'left', speed: 'fast' });
  slide5.setBackground('F8FAFC');

  slide5.addText('Three-Tier Decoupled Platform Architecture', {
    bold: true,
    color: '0F172A',
    font: '+mj-lt',
    fontSize: points(24),
    h: inches(0.6),
    w: inches(11.33),
    x: inches(1),
    y: inches(0.8),
  });

  slide5.addText('Ingress routing, stateless isolate compilation, and distributed object persistence', {
    color: '64748B',
    fontSize: points(14),
    h: inches(0.4),
    w: inches(11.33),
    x: inches(1),
    y: inches(1.4),
  });

  const cards = [
    {
      badge: 'TIER 1 • INGRESS',
      badgeColor: '0284C7',
      bullets: ['Anycast DNS routing (330+ locations)', 'TLS 1.3 0-RTT termination', 'DDoS protection & Rate limiting'],
      title: 'Global Edge Router',
      x: 1.0,
    },
    {
      badge: 'TIER 2 • COMPUTE',
      badgeColor: '6366F1',
      bullets: ['V8 Isolate sandboxed execution', 'Instant <5ms cold starts', 'Streaming SSR & WebSockets'],
      title: 'Stateless Worker Fleet',
      x: 4.8,
    },
    {
      badge: 'TIER 3 • STORAGE',
      badgeColor: '10B981',
      bullets: ['Globally replicated R2 object store', 'Low-latency D1 serverless SQLite', 'Workers KV key-value cache'],
      title: 'Distributed State Store',
      x: 8.6,
    },
  ];

  for (const c of cards) {
    slide5.addShape('roundRect', {
      fill: 'FFFFFF',
      h: inches(4.5),
      line: { color: 'E2E8F0', width: inches(0.015) },
      shadow: {
        blur: inches(0.15),
        color: '0F172A',
        direction: degrees(90),
        distance: inches(0.05),
        opacity: 0.12,
      },
      w: inches(3.6),
      x: inches(c.x),
      y: inches(2.0),
    });

    slide5.addShape('roundRect', {
      fill: 'F8FAFC',
      h: inches(0.35),
      line: { color: c.badgeColor, width: inches(0.01) },
      text: c.badge,
      textOptions: { bold: true, color: c.badgeColor, fontSize: points(9), align: 'center' },
      w: inches(2.0),
      x: inches(c.x + 0.3),
      y: inches(2.3),
    });

    slide5.addText(c.title, {
      bold: true,
      color: '0F172A',
      fontSize: points(16),
      h: inches(0.5),
      w: inches(3.0),
      x: inches(c.x + 0.3),
      y: inches(2.8),
    });

    slide5.addText(c.bullets.map((b) => ({ level: 0, text: b })), {
      color: '475569',
      fontSize: points(11),
      h: inches(2.5),
      w: inches(3.0),
      x: inches(c.x + 0.3),
      y: inches(3.4),
    });
  }

  addDisclaimer(slide5, false);

  // ==================================================================
  // SLIDE 6: Isomorphic Media & High-Fidelity Image Embedding
  // ==================================================================
  console.log('🖼️ Generating Slide 6: High-Fidelity Image Embedding & Media Gallery...');
  const slide6 = pres.addSlide();
  slide6.setTransition('fade', { speed: 'med' });
  slide6.setBackground('F8FAFC');

  slide6.addText('Isomorphic Media & High-Fidelity Image Embedding', {
    bold: true,
    color: '0F172A',
    font: '+mj-lt',
    fontSize: points(24),
    h: inches(0.6),
    w: inches(11.33),
    x: inches(1.0),
    y: inches(0.8),
  });

  slide6.addText('Embedded raster graphics (PNG, JPEG), transparent alpha channels, aspect ratio scaling, and card layouts', {
    color: '64748B',
    fontSize: points(14),
    h: inches(0.4),
    w: inches(11.33),
    x: inches(1),
    y: inches(1.4),
  });

  // Left Card: Main Hero Image Showcase
  slide6.addShape('roundRect', {
    fill: 'FFFFFF',
    h: inches(4.5),
    line: { color: 'E2E8F0', width: inches(0.015) },
    shadow: {
      blur: inches(0.15),
      color: '0F172A',
      direction: degrees(90),
      distance: inches(0.05),
      opacity: 0.1,
    },
    w: inches(5.4),
    x: inches(1.0),
    y: inches(2.0),
  });

  const heroImageBytes = readFileSync(resolve(process.cwd(), 'assets/showcase-hero.jpg'));
  slide6.addImage(heroImageBytes, {
    fileName: 'showcase-hero.jpg',
    h: inches(2.7),
    w: inches(4.8),
    x: inches(1.3),
    y: inches(2.3),
  });

  slide6.addText('Rich Visual Asset Integration', {
    bold: true,
    color: '0F172A',
    fontSize: points(15),
    h: inches(0.4),
    w: inches(4.8),
    x: inches(1.3),
    y: inches(5.2),
  });

  slide6.addText('Automatic MIME-type detection, relationship mapping, and lossless OpenXML zip packaging.', {
    color: '64748B',
    fontSize: points(11),
    h: inches(0.6),
    w: inches(4.8),
    x: inches(1.3),
    y: inches(5.6),
  });

  // Right Card: Logo, Alpha Transparency & Specs
  slide6.addShape('roundRect', {
    fill: 'FFFFFF',
    h: inches(4.5),
    line: { color: 'E2E8F0', width: inches(0.015) },
    shadow: {
      blur: inches(0.15),
      color: '0F172A',
      direction: degrees(90),
      distance: inches(0.05),
      opacity: 0.1,
    },
    w: inches(5.4),
    x: inches(6.93),
    y: inches(2.0),
  });

  // Logo Badge with Dark Accent Background
  slide6.addShape('roundRect', {
    fill: '0F172A',
    h: inches(1.2),
    w: inches(4.8),
    x: inches(7.23),
    y: inches(2.3),
  });

  const logoBytes = readFileSync(resolve(process.cwd(), 'assets/logo.png'));
  slide6.addImage(logoBytes, {
    fileName: 'logo.png',
    h: inches(0.8),
    w: inches(2.0),
    x: inches(8.63),
    y: inches(2.5),
  });

  slide6.addText('Supported Image Codecs & Features', {
    bold: true,
    color: '0F172A',
    fontSize: points(15),
    h: inches(0.4),
    w: inches(4.8),
    x: inches(7.23),
    y: inches(3.7),
  });

  slide6.addText([
    { bullet: true, level: 0, text: 'PNG with 32-bit RGBA alpha transparency' },
    { bullet: true, level: 0, text: 'JPEG / JPG progressive and baseline encoding' },
    { bullet: true, level: 0, text: 'Zero-copy ArrayBuffer & Uint8Array input' },
    { bullet: true, level: 0, text: 'Placeholder resolution into slide layout picture frames' },
  ], {
    color: '475569',
    fontSize: points(11),
    h: inches(2.2),
    w: inches(4.8),
    x: inches(7.23),
    y: inches(4.2),
  });

  slide6.setNotes([
    {
      runs: [
        { text: 'Media Integration Keynote: ', bold: true },
        { text: 'Deterministic Binary Packaging', underline: true },
      ],
    },
    { bullet: true, level: 0, text: 'Binary streams are compressed with fflate without intermediate disk I/O.' },
    { bullet: true, level: 0, text: 'All media elements automatically link to ppt/media/ and content types.' },
  ]);

  addDisclaimer(slide6, false);

  // ==================================================================
  // SLIDE 7: Deep Multilevel Bullet Matrix & Nested Outlines
  // ==================================================================
  console.log('📑 Generating Slide 7: Deep Multilevel Bullet Matrix...');
  const slide7 = pres.addSlide();
  slide7.setTransition('push', { direction: 'left', speed: 'fast' });
  slide7.setBackground('0F172A'); // Dark theme

  slide7.addText('Complex Multi-Level Outlining & Hierarchy Matrix', {
    bold: true,
    color: '38BDF8',
    font: '+mj-lt',
    fontSize: points(24),
    h: inches(0.6),
    w: inches(11.33),
    x: inches(1.0),
    y: inches(0.8),
  });

  slide7.addText('Demonstrating 4 levels of hierarchical indentation, mixed bullet glyphs, auto-numbering, and paragraph spacing', {
    color: '94A3B8',
    fontSize: points(14),
    h: inches(0.4),
    w: inches(11.33),
    x: inches(1),
    y: inches(1.4),
  });

  // Left Column Card (Dark Card with Blue Glow Outline)
  slide7.addShape('roundRect', {
    fill: '1E293B',
    h: inches(4.6),
    line: { color: '38BDF8', width: inches(0.015) },
    shadow: {
      blur: inches(0.2),
      color: '000000',
      direction: degrees(90),
      distance: inches(0.05),
      opacity: 0.4,
    },
    w: inches(5.4),
    x: inches(1.0),
    y: inches(2.0),
  });

  slide7.addText('System Topology Outlining', {
    bold: true,
    color: '38BDF8',
    fontSize: points(15),
    h: inches(0.4),
    w: inches(4.8),
    x: inches(1.3),
    y: inches(2.2),
  });

  slide7.addText([
    { bullet: true, level: 0, text: 'Core Infrastructure Fleet' },
    { bullet: true, level: 1, text: 'Edge Routing Layer (Anycast DNS)' },
    { bullet: true, level: 2, text: 'TLS 1.3 Handshake Termination' },
    { bullet: true, level: 3, text: 'Zero-RTT session resumption' },
    { bullet: true, level: 3, text: 'DDoS mitigation heuristic filters' },
    { bullet: true, level: 2, text: 'Geo-Distributed Compute Mesh' },
    { bullet: true, level: 3, text: 'Sub-millisecond V8 isolate provisioning' },
    { bullet: true, level: 1, text: 'Global State & Cache Synchronization' },
    { bullet: true, level: 2, text: 'D1 Relational Database SQLite Replicas' },
    { bullet: true, level: 2, text: 'KV Cache (Read-heavy eventual consistency)' },
  ], {
    color: 'CBD5E1',
    fontSize: points(11),
    h: inches(3.8),
    w: inches(4.8),
    x: inches(1.3),
    y: inches(2.6),
  });

  // Right Column Card (Dark Card with Indigo Accent)
  slide7.addShape('roundRect', {
    fill: '1E293B',
    h: inches(4.6),
    line: { color: '6366F1', width: inches(0.015) },
    shadow: {
      blur: inches(0.2),
      color: '000000',
      direction: degrees(90),
      distance: inches(0.05),
      opacity: 0.4,
    },
    w: inches(5.4),
    x: inches(6.93),
    y: inches(2.0),
  });

  slide7.addText('Compilation Pipeline Lifecycle', {
    bold: true,
    color: '818CF8',
    fontSize: points(15),
    h: inches(0.4),
    w: inches(4.8),
    x: inches(7.23),
    y: inches(2.2),
  });

  slide7.addText([
    { bullet: 'number', level: 0, text: 'Phase 1: Ingestion & Parsing' },
    { bullet: true, level: 1, text: 'Decompress OpenXML container with fflate' },
    { bullet: true, level: 2, text: 'Stream XML parts directly to DOM parser' },
    { bullet: true, level: 2, text: 'Extract theme, slideMaster, and slideLayouts' },
    { bullet: 'number', level: 0, text: 'Phase 2: AST Mutation & Modeling' },
    { bullet: true, level: 1, text: 'Apply shape geometries, tables, and drop shadows' },
    { bullet: true, level: 2, text: 'Calculate type-safe EMU, points, and inches' },
    { bullet: 'number', level: 0, text: 'Phase 3: Serialization & Packaging' },
    { bullet: true, level: 1, text: 'Emit deterministic XML parts and relationship files' },
    { bullet: true, level: 1, text: 'Generate zero-copy ArrayBuffer presentation binary' },
  ], {
    color: 'CBD5E1',
    fontSize: points(11),
    h: inches(3.8),
    w: inches(4.8),
    x: inches(7.23),
    y: inches(2.6),
  });

  slide7.setNotes([
    {
      runs: [
        { text: 'Presenter Notes — Multi-Level Hierarchy Details: ', bold: true, underline: true },
      ],
    },
    { bullet: true, level: 0, text: 'Explain how levels 0 through 3 compute distinct left margins and hanging indents.' },
    { bullet: true, level: 1, text: 'Level 0: 0.166" bullet gap, 0" margin' },
    { bullet: true, level: 1, text: 'Level 1: 0.166" bullet gap, 0.25" margin' },
    { bullet: true, level: 2, text: 'Level 2: 0.166" bullet gap, 0.50" margin' },
    { bullet: true, level: 3, text: 'Level 3: 0.166" bullet gap, 0.75" margin' },
    { bullet: 'number', level: 0, text: 'Emphasize that number lists automatically use 0.222" clearance for digits.' },
  ]);

  addDisclaimer(slide7, true);

  // ==================================================================
  // SLIDE 8: Connectors, Process Flow Pipelines & Composite Groups
  // ==================================================================
  console.log('🔄 Generating Slide 8: Connectors & Composite Group Pipelines...');
  const slide8 = pres.addSlide();
  slide8.setTransition('pull', { direction: 'right', speed: 'fast' });
  slide8.setBackground('F8FAFC');

  slide8.addText('Flow Connectors & Composite Group Architecture', {
    bold: true,
    color: '0F172A',
    font: '+mj-lt',
    fontSize: points(24),
    h: inches(0.6),
    w: inches(11.33),
    x: inches(1.0),
    y: inches(0.8),
  });

  slide8.addText('Demonstrating vector connectors (solid, dashed), coordinates, and composite group containers (addGroup)', {
    color: '64748B',
    fontSize: points(14),
    h: inches(0.4),
    w: inches(11.33),
    x: inches(1),
    y: inches(1.4),
  });

  // Stage 1 Group: Ingress & Decompress
  slide8.addGroup({
    h: inches(3.5),
    id: 'stage-1',
    name: 'Stage 1 Ingress Container',
    w: inches(3.2),
    x: inches(1.0),
    y: inches(2.2),
  }, (g) => {
    g.addShape('roundRect', {
      fill: 'FFFFFF',
      h: inches(3.5),
      line: { color: '0284C7', width: inches(0.015) },
      shadow: { blur: inches(0.12), color: '0F172A', opacity: 0.1 },
      w: inches(3.2),
      x: inches(1.0),
      y: inches(2.2),
    });
    g.addShape('roundRect', {
      fill: '0284C7',
      h: inches(0.4),
      text: 'STEP 1: INGESTION',
      textOptions: { bold: true, color: 'FFFFFF', fontSize: points(10), align: 'center' },
      w: inches(2.6),
      x: inches(1.3),
      y: inches(2.5),
    });
    g.addText([
      { bullet: true, level: 0, text: 'Streaming ZIP decompression' },
      { bullet: true, level: 0, text: 'XML DOM part resolution' },
      { bullet: true, level: 0, text: 'Extract relationships map' },
    ], {
      color: '475569',
      fontSize: points(11),
      h: inches(2.0),
      w: inches(2.6),
      x: inches(1.3),
      y: inches(3.2),
    });
  });

  // Stage 2 Group: AST Modeling & Mutator
  slide8.addGroup({
    h: inches(3.5),
    id: 'stage-2',
    name: 'Stage 2 AST Container',
    w: inches(3.2),
    x: inches(5.0),
    y: inches(2.2),
  }, (g) => {
    g.addShape('roundRect', {
      fill: 'FFFFFF',
      h: inches(3.5),
      line: { color: '6366F1', width: inches(0.015) },
      shadow: { blur: inches(0.12), color: '0F172A', opacity: 0.1 },
      w: inches(3.2),
      x: inches(5.0),
      y: inches(2.2),
    });
    g.addShape('roundRect', {
      fill: '6366F1',
      h: inches(0.4),
      text: 'STEP 2: MODELING',
      textOptions: { bold: true, color: 'FFFFFF', fontSize: points(10), align: 'center' },
      w: inches(2.6),
      x: inches(5.3),
      y: inches(2.5),
    });
    g.addText([
      { bullet: true, level: 0, text: 'Type-safe AST generation' },
      { bullet: true, level: 0, text: 'Layer & placeholder cascade' },
      { bullet: true, level: 0, text: 'Fluent mutations & styling' },
    ], {
      color: '475569',
      fontSize: points(11),
      h: inches(2.0),
      w: inches(2.6),
      x: inches(5.3),
      y: inches(3.2),
    });
  });

  // Stage 3 Group: Serialization & Packaging
  slide8.addGroup({
    h: inches(3.5),
    id: 'stage-3',
    name: 'Stage 3 Packaging Container',
    w: inches(3.2),
    x: inches(9.0),
    y: inches(2.2),
  }, (g) => {
    g.addShape('roundRect', {
      fill: 'FFFFFF',
      h: inches(3.5),
      line: { color: '10B981', width: inches(0.015) },
      shadow: { blur: inches(0.12), color: '0F172A', opacity: 0.1 },
      w: inches(3.2),
      x: inches(9.0),
      y: inches(2.2),
    });
    g.addShape('roundRect', {
      fill: '10B981',
      h: inches(0.4),
      text: 'STEP 3: COMPILATION',
      textOptions: { bold: true, color: 'FFFFFF', fontSize: points(10), align: 'center' },
      w: inches(2.6),
      x: inches(9.3),
      y: inches(2.5),
    });
    g.addText([
      { bullet: true, level: 0, text: 'Deterministic XML emission' },
      { bullet: true, level: 0, text: 'Zero-copy fflate compression' },
      { bullet: true, level: 0, text: 'ECMA-376 compliant PPTX' },
    ], {
      color: '475569',
      fontSize: points(11),
      h: inches(2.0),
      w: inches(2.6),
      x: inches(9.3),
      y: inches(3.2),
    });
  });

  // Connector 1 -> 2 (glued to stage-1 right side and stage-2 left side)
  slide8.addConnector({
    color: '0284C7',
    dashStyle: 'solid',
    endArrow: { length: 'lg', type: 'triangle', width: 'lg' },
    from: { position: 'right', shapeId: 'stage-1' },
    to: { position: 'left', shapeId: 'stage-2' },
    width: inches(0.03),
  });

  // Connector 2 -> 3 (glued to stage-2 right side and stage-3 left side)
  slide8.addConnector({
    color: '6366F1',
    dashStyle: 'solid',
    endArrow: { length: 'lg', type: 'stealth', width: 'lg' },
    from: { position: 'right', shapeId: 'stage-2' },
    to: { position: 'left', shapeId: 'stage-3' },
    width: inches(0.03),
  });

  // Connector feedback loop (dashed return line with arrows on both ends)
  slide8.addConnector({
    color: '94A3B8',
    dashStyle: 'dash',
    endArrow: 'triangle',
    from: { x: inches(1.0), y: inches(6.2) },
    startArrow: 'oval',
    to: { x: inches(12.2), y: inches(6.2) },
    width: inches(0.02),
  });

  slide8.addText('Isomorphic Pipeline Feedback Loop: 100% Roundtrip Bit-Exact Precision', {
    color: '64748B',
    fontSize: points(11),
    h: inches(0.4),
    w: inches(11.2),
    x: inches(1.0),
    y: inches(6.35),
    align: 'center',
  });

  slide8.setNotes([
    {
      runs: [
        { text: 'Architecture Deep Dive: ', bold: true },
        { text: 'Composite Groups & Vector Connectors', underline: true },
      ],
    },
    { bullet: true, level: 0, text: 'addGroup wraps child elements inside <p:grpSp> with coordinated bounding transforms.' },
    { bullet: true, level: 0, text: 'addConnector emits <p:cxnSp> connecting arbitrary coordinates across slide elements.' },
  ]);

  addDisclaimer(slide8, false);

  // ==================================================================
  // SLIDE 9: Native OpenXML Charts & Visual Analytics
  // ==================================================================
  console.log('📊 Generating Slide 9: Native OpenXML Charts & Visual Analytics...');
  const slide9 = pres.addSlide();
  slide9.setTransition('wheel', { spokes: 4, speed: 'med' });
  slide9.setBackground('0F172A'); // Dark theme

  slide9.addText('Native OpenXML Charts & Visual Analytics', {
    bold: true,
    color: '38BDF8',
    font: '+mj-lt',
    fontSize: points(24),
    h: inches(0.6),
    w: inches(11.33),
    x: inches(1.0),
    y: inches(0.8),
  });

  slide9.addText('Zero-dependency generation of native PowerPoint <c:chartSpace> data visualizations and multi-series metrics', {
    color: '94A3B8',
    fontSize: points(14),
    h: inches(0.4),
    w: inches(11.33),
    x: inches(1.0),
    y: inches(1.4),
  });

  // Chart Container Card
  slide9.addShape('roundRect', {
    fill: '1E293B',
    h: inches(4.5),
    line: { color: '334155', width: inches(0.015) },
    shadow: { blur: inches(0.15), color: '000000', opacity: 0.3 },
    w: inches(7.0),
    x: inches(1.0),
    y: inches(2.0),
  });

  // Embed Native Column Chart with Dark Theme High-Contrast Styling
  slide9.addChart({
    axisColor: '475569',
    categories: ['Cloudflare Workers', 'Bun 1.1', 'Node.js 22', 'Deno 2.0'],
    chartType: 'column',
    gridColor: '334155',
    h: inches(4.1),
    legend: { color: 'F1F5F9', position: 'top' },
    series: [
      {
        color: '38BDF8',
        name: 'Throughput (MB/s)',
        values: [145, 138, 120, 110],
      },
      {
        color: '34D399',
        name: 'Parse Ops/sec',
        values: [850, 780, 640, 610],
      },
      {
        color: 'A78BFA',
        name: 'Compression Ratio (%)',
        values: [78, 76, 75, 74],
      },
    ],
    showGridlines: true,
    textColor: 'CBD5E1',
    w: inches(6.6),
    x: inches(1.2),
    y: inches(2.2),
  });

  // Analytics Insight Cards on the right
  const rightX = inches(8.3);
  slide9.addShape('roundRect', {
    fill: '1E293B',
    h: inches(1.35),
    line: { color: '0284C7', width: inches(0.015) },
    text: '⚡ Zero Native Bindings',
    textOptions: { bold: true, color: '38BDF8', fontSize: points(13), align: 'left' },
    w: inches(4.0),
    x: rightX,
    y: inches(2.0),
  });

  slide9.addText('Generates ECMA-376 compliant DrawingML charts directly from pure TypeScript arrays without spawning Python, Java, or Excel subprocesses.', {
    color: '94A3B8',
    fontSize: points(10),
    h: inches(0.8),
    w: inches(3.6),
    x: inches(8.5),
    y: inches(2.45),
  });

  slide9.addShape('roundRect', {
    fill: '1E293B',
    h: inches(1.35),
    line: { color: '10B981', width: inches(0.015) },
    text: '📈 Fully Editable & Scalable',
    textOptions: { bold: true, color: '34D399', fontSize: points(13), align: 'left' },
    w: inches(4.0),
    x: rightX,
    y: inches(3.55),
  });

  slide9.addText('Embedded as vector OpenXML chart spaces. Users can edit values, change themes, or modify series formatting natively inside Microsoft PowerPoint.', {
    color: '94A3B8',
    fontSize: points(10),
    h: inches(0.8),
    w: inches(3.6),
    x: inches(8.5),
    y: inches(4.0),
  });

  slide9.addShape('roundRect', {
    fill: '1E293B',
    h: inches(1.35),
    line: { color: '6366F1', width: inches(0.015) },
    text: '🎨 Multi-Type Chart Engine',
    textOptions: { bold: true, color: '818CF8', fontSize: points(13), align: 'left' },
    w: inches(4.0),
    x: rightX,
    y: inches(5.1),
  });

  slide9.addText('Supports column, bar, line, area, and pie chart topologies with custom series coloring, categorical axes, and configurable legend alignments.', {
    color: '94A3B8',
    fontSize: points(10),
    h: inches(0.8),
    w: inches(3.6),
    x: inches(8.5),
    y: inches(5.55),
  });

  slide9.setNotes([
    {
      runs: [
        { text: 'Chart Keynote: ', bold: true },
        { text: 'Native OpenXML DrawingML Charts', underline: true },
      ],
    },
    { bullet: true, level: 0, text: 'Emits standard <c:chartSpace> parts under ppt/charts/chartN.xml.' },
    { bullet: true, level: 0, text: 'Fully editable in PowerPoint without external Excel binary dependencies.' },
    { bullet: true, level: 0, text: 'Custom series color fills and categorical string/number axis mapping.' },
  ]);

  addDisclaimer(slide9, true);

  // ==================================================================
  // SLIDE 10: Multi-Topology Chart Gallery
  // -------------------------------------------------------------
  console.log('📈 Generating Slide 10: Multi-Topology Chart Gallery...');
  const slide10 = pres.addSlide();
  slide10.setTransition('blinds', { direction: 'horz', speed: 'fast' });
  slide10.setBackground('F8FAFC');

  slide10.addText('Multi-Topology Chart Gallery', {
    bold: true,
    color: '0F172A',
    font: '+mj-lt',
    fontSize: points(24),
    h: inches(0.6),
    w: inches(11.33),
    x: inches(1.0),
    y: inches(0.8),
  });

  slide10.addText('Smooth spline line trends, radial doughnut distributions, and horizontal bar comparisons', {
    color: '64748B',
    fontSize: points(14),
    h: inches(0.4),
    w: inches(11.33),
    x: inches(1.0),
    y: inches(1.4),
  });

  // Card 1: Line Chart (Smooth Spline)
  const cardW = inches(3.6);
  const cardH = inches(4.5);
  const cardY = inches(2.0);

  slide10.addShape('roundRect', {
    fill: 'FFFFFF',
    h: cardH,
    line: { color: 'E2E8F0', width: inches(0.015) },
    shadow: { blur: inches(0.12), color: '0F172A', opacity: 0.08 },
    w: cardW,
    x: inches(1.0),
    y: cardY,
  });

  slide10.addText('📈 Latency Curve vs Payload (ms)', {
    bold: true,
    color: '0F172A',
    fontSize: points(12),
    h: inches(0.35),
    w: inches(3.2),
    x: inches(1.2),
    y: inches(2.2),
  });

  slide10.addChart({
    categories: ['100K', '500K', '1MB', '5MB', '10MB'],
    chartType: 'line',
    h: inches(3.6),
    legend: { position: 'bottom' },
    series: [
      {
        color: '0284C7',
        name: '@hokkyss/pptx',
        values: [2.1, 4.8, 8.5, 32.0, 64.2],
      },
      {
        color: 'EF4444',
        name: 'Legacy Parser',
        values: [28.0, 65.0, 140.0, 580.0, 1250.0],
      },
    ],
    smooth: true,
    w: inches(3.2),
    x: inches(1.2),
    y: inches(2.6),
  });

  // Card 2: Doughnut Chart (Distribution)
  slide10.addShape('roundRect', {
    fill: 'FFFFFF',
    h: cardH,
    line: { color: 'E2E8F0', width: inches(0.015) },
    shadow: { blur: inches(0.12), color: '0F172A', opacity: 0.08 },
    w: cardW,
    x: inches(4.86),
    y: cardY,
  });

  slide10.addText('🍩 Runtime Execution Share', {
    bold: true,
    color: '0F172A',
    fontSize: points(12),
    h: inches(0.35),
    w: inches(3.2),
    x: inches(5.06),
    y: inches(2.2),
  });

  slide10.addChart({
    categories: ['Cloudflare Workers', 'Client Browser', 'Edge V8', 'Node.js'],
    chartType: 'doughnut',
    dataLabels: { showPercent: true, showVal: false },
    h: inches(3.6),
    holeSize: 55,
    legend: { position: 'bottom' },
    series: [
      {
        color: '6366F1',
        name: 'Runtime Deployments',
        values: [45, 30, 15, 10],
      },
    ],
    w: inches(3.2),
    x: inches(5.06),
    y: inches(2.6),
  });

  // Card 3: Horizontal Bar Chart (Format Comparison)
  slide10.addShape('roundRect', {
    fill: 'FFFFFF',
    h: cardH,
    line: { color: 'E2E8F0', width: inches(0.015) },
    shadow: { blur: inches(0.12), color: '0F172A', opacity: 0.08 },
    w: cardW,
    x: inches(8.73),
    y: cardY,
  });

  slide10.addText('📊 Ecosystem Compression (x)', {
    bold: true,
    color: '0F172A',
    fontSize: points(12),
    h: inches(0.35),
    w: inches(3.2),
    x: inches(8.93),
    y: inches(2.2),
  });

  slide10.addChart({
    categories: ['Raw XML', 'Uncompressed ZIP', 'Standard Deflate', 'fflate Pure TS'],
    chartType: 'horizontalBar',
    h: inches(3.6),
    legend: { position: 'bottom' },
    series: [
      {
        color: '10B981',
        name: 'Ratio Efficiency',
        values: [1.0, 1.8, 3.4, 4.2],
      },
    ],
    w: inches(3.2),
    x: inches(8.93),
    y: inches(2.6),
  });

  slide10.setNotes([
    {
      runs: [
        { text: 'Chart Suite Overview: ', bold: true },
        { text: 'Line, Doughnut & Horizontal Bar Support', underline: true },
      ],
    },
    { bullet: true, level: 0, text: 'Smooth spline curves for continuous time-series metrics.' },
    { bullet: true, level: 0, text: 'Customizable doughnut hole radii and automated data label percentages.' },
    { bullet: true, level: 0, text: 'Horizontal bar orientation for ranked ordinal comparisons.' },
  ]);

  addDisclaimer(slide10, false);

  // ==================================================================
  // SLIDE 11: Specialized Analytical Topologies (Area, Pie, Radar)
  // ==================================================================
  console.log('🎯 Generating Slide 11: Specialized Analytical Topologies (Area, Pie, Radar)...');
  const slide11 = pres.addSlide();
  slide11.setTransition('checker', { direction: 'vert', speed: 'fast' });
  slide11.setBackground('0F172A'); // Dark Executive Theme

  slide11.addText('Specialized Analytical Topologies', {
    bold: true,
    color: '38BDF8',
    font: '+mj-lt',
    fontSize: points(24),
    h: inches(0.6),
    w: inches(11.33),
    x: inches(1.0),
    y: inches(0.8),
  });

  slide11.addText('Multi-series stacked area trends, radial composition pie distributions, and multi-axis radar scorecards', {
    color: '94A3B8',
    fontSize: points(14),
    h: inches(0.4),
    w: inches(11.33),
    x: inches(1.0),
    y: inches(1.4),
  });

  const card11W = inches(3.6);
  const card11H = inches(4.5);
  const card11Y = inches(2.0);

  // Card 1: Stacked Area Chart (Dark Theme)
  slide11.addShape('roundRect', {
    fill: '1E293B',
    h: card11H,
    line: { color: '334155', width: inches(0.015) },
    shadow: { blur: inches(0.12), color: '000000', opacity: 0.3 },
    w: card11W,
    x: inches(1.0),
    y: card11Y,
  });

  slide11.addText('🌊 Stacked Area: Memory Footprint (MB)', {
    bold: true,
    color: '38BDF8',
    fontSize: points(12),
    h: inches(0.35),
    w: inches(3.2),
    x: inches(1.2),
    y: inches(2.2),
  });

  slide11.addChart({
    axisColor: '475569',
    categories: ['10K', '50K', '250K', '1M', '5M'],
    chartType: 'area',
    gridColor: '334155',
    grouping: 'stacked',
    h: inches(3.6),
    legend: { color: 'F1F5F9', position: 'bottom' },
    series: [
      {
        color: '38BDF8',
        name: 'V8 Heap',
        values: [1.2, 2.5, 5.8, 14.2, 28.0],
      },
      {
        color: '818CF8',
        name: 'AST Buffer',
        values: [0.8, 1.6, 3.9, 9.4, 18.5],
      },
      {
        color: '34D399',
        name: 'Deflate Stream',
        values: [0.4, 0.9, 2.1, 5.2, 10.1],
      },
    ],
    showGridlines: true,
    textColor: 'CBD5E1',
    w: inches(3.2),
    x: inches(1.2),
    y: inches(2.6),
  });

  // Card 2: 2D Pie Chart with Data Labels (Dark Theme)
  slide11.addShape('roundRect', {
    fill: '1E293B',
    h: card11H,
    line: { color: '334155', width: inches(0.015) },
    shadow: { blur: inches(0.12), color: '000000', opacity: 0.3 },
    w: card11W,
    x: inches(4.86),
    y: card11Y,
  });

  slide11.addText('🥧 2D Pie: Bundle Contribution', {
    bold: true,
    color: '34D399',
    fontSize: points(12),
    h: inches(0.35),
    w: inches(3.2),
    x: inches(5.06),
    y: inches(2.2),
  });

  slide11.addChart({
    categories: ['@pptx-reader', '@pptx-writer', '@pptx', '@pptx-core'],
    chartType: 'pie',
    colors: ['38BDF8', '818CF8', '34D399', 'FBBF24'],
    dataLabels: { showPercent: true, showVal: false },
    h: inches(3.6),
    legend: { color: 'F1F5F9', position: 'bottom' },
    series: [
      {
        name: 'Package Share',
        values: [38, 32, 18, 12],
      },
    ],
    textColor: 'CBD5E1',
    w: inches(3.2),
    x: inches(5.06),
    y: inches(2.6),
  });

  // Card 3: Multi-Series Radar / Spider Chart (Dark Theme)
  slide11.addShape('roundRect', {
    fill: '1E293B',
    h: card11H,
    line: { color: '334155', width: inches(0.015) },
    shadow: { blur: inches(0.12), color: '000000', opacity: 0.3 },
    w: card11W,
    x: inches(8.73),
    y: card11Y,
  });

  slide11.addText('🎯 Radar: Performance Scorecard', {
    bold: true,
    color: '818CF8',
    fontSize: points(12),
    h: inches(0.35),
    w: inches(3.2),
    x: inches(8.93),
    y: inches(2.2),
  });

  slide11.addChart({
    axisColor: '475569',
    categories: ['Throughput', 'Type Safety', 'Zero Deps', 'Spec Fidelity', 'Bundle Size', 'Cold Start'],
    chartType: 'radar',
    gridColor: '334155',
    h: inches(3.6),
    legend: { color: 'F1F5F9', position: 'bottom' },
    series: [
      {
        color: '34D399',
        name: '@hokkyss/pptx',
        values: [98, 100, 100, 99, 95, 99],
      },
      {
        color: 'F43F5E',
        name: 'Competitor A',
        values: [65, 70, 45, 80, 60, 55],
      },
    ],
    showGridlines: true,
    textColor: 'CBD5E1',
    w: inches(3.2),
    x: inches(8.93),
    y: inches(2.6),
  });

  slide11.setNotes([
    {
      runs: [
        { text: 'Analytical Topology Keynote: ', bold: true },
        { text: 'Area, Pie & Radar Spider Visualizations', underline: true },
      ],
    },
    { bullet: true, level: 0, text: 'Stacked area topologies for cumulative resource metric trends.' },
    { bullet: true, level: 0, text: 'Standard pie charts with automatic slice percentage data labels.' },
    { bullet: true, level: 0, text: 'Radar / spider charts for multi-dimensional multi-series competitive scorecards.' },
  ]);

  addDisclaimer(slide11, true);

  // ==================================================================
  // SLIDE 12: Native Hyperlinks & Interactive Presentation Navigation
  // ==================================================================
  console.log('🔗 Generating Slide 12: Native Hyperlinks & Interactive Navigation System...');
  const slide12 = pres.addSlide();
  slide12.setTransition('fade', { durationMs: 400 });
  slide12.setBackground('F8FAFC');

  slide12.addText('Native Hyperlinks & Interactive Navigation System', {
    bold: true,
    color: '0F172A',
    font: '+mj-lt',
    fontSize: points(24),
    h: inches(0.6),
    w: inches(11.73),
    x: inches(0.8),
    y: inches(0.8),
  });

  slide12.addText('OpenXML DrawingML <a:hlinkClick> external URLs, slide navigation jumps, hover ScreenTips, and presentation controls', {
    color: '64748B',
    fontSize: points(14),
    h: inches(0.4),
    w: inches(11.73),
    x: inches(0.8),
    y: inches(1.4),
  });

  // Left Card: Interactive Slide Navigation Hub
  slide12.addShape('roundRect', {
    fill: 'FFFFFF',
    h: inches(4.9),
    line: { color: 'E2E8F0', width: inches(0.015) },
    shadow: { blur: inches(0.12), color: '0F172A', opacity: 0.08 },
    w: inches(5.7),
    x: inches(0.8),
    y: inches(2.0),
  });

  slide12.addText('🧭 Interactive Presentation Jump Hub', {
    bold: true,
    color: '0F172A',
    fontSize: points(13),
    h: inches(0.35),
    w: inches(5.3),
    x: inches(1.0),
    y: inches(2.2),
  });

  slide12.addText('Click any card below during slide show mode to immediately jump to that section:', {
    color: '64748B',
    fontSize: points(10.5),
    h: inches(0.3),
    w: inches(5.3),
    x: inches(1.0),
    y: inches(2.55),
  });

  const jumpTiles = [
    {
      desc: '3-tier modular schemas & strict units',
      fill: 'EFF6FF',
      line: '3B82F6',
      slideIdx: 2,
      textColor: '1D4ED8',
      title: '📐 Slide 2: Modular Architecture',
      y: 2.95,
    },
    {
      desc: 'Clustered bar & line dual-series analytics',
      fill: 'ECFDF5',
      line: '10B981',
      slideIdx: 9,
      textColor: '047857',
      title: '📈 Slide 9: OpenXML Charts & Graphs',
      y: 3.9,
    },
    {
      desc: 'Linear, radial, alpha & multi-stop mesh cards',
      fill: 'F5F3FF',
      line: '8B5CF6',
      slideIdx: 13,
      textColor: '6D28D9',
      title: '🎨 Slide 13: DrawingML Gradient Engine',
      y: 4.85,
    },
    {
      desc: 'Empirical throughput & bundle size matrix',
      fill: 'FFFBEB',
      line: 'F59E0B',
      slideIdx: 14,
      textColor: 'B45309',
      title: '⚡ Slide 14: Engine Benchmark Matrix',
      y: 5.8,
    },
  ];

  for (const tile of jumpTiles) {
    slide12.addShape('roundRect', {
      fill: tile.fill,
      h: inches(0.82),
      hyperlink: { slideIndex: tile.slideIdx, tooltip: `Jump directly to Slide ${tile.slideIdx}` },
      line: { color: tile.line, width: inches(0.015) },
      text: [
        { bold: true, color: tile.textColor, fontSize: points(11), text: `${tile.title} ➔\n` },
        { color: '64748B', fontSize: points(9.5), text: tile.desc },
      ],
      w: inches(5.3),
      x: inches(1.0),
      y: inches(tile.y),
    });
  }

  // Right Top Card: External Endpoints & Rich Text Hyperlinks
  slide12.addShape('roundRect', {
    fill: 'FFFFFF',
    h: inches(2.6),
    line: { color: 'E2E8F0', width: inches(0.015) },
    shadow: { blur: inches(0.12), color: '0F172A', opacity: 0.08 },
    w: inches(5.7),
    x: inches(6.8),
    y: inches(2.0),
  });

  slide12.addText('🔗 External Web & Ecosystem Endpoints', {
    bold: true,
    color: '0F172A',
    fontSize: points(13),
    h: inches(0.35),
    w: inches(5.3),
    x: inches(7.0),
    y: inches(2.2),
  });

  slide12.addText([
    {
      level: 0,
      runs: [
        { text: '• ' },
        {
          bold: true,
          color: '0284C7',
          hyperlink: { tooltip: 'View GitHub Repository & Source Code', url: 'https://github.com/hokkyss/pptx-parser' },
          text: 'GitHub Repository (hokkyss/pptx-parser)',
          underline: true,
        },
        { text: ' — Star & contribute.' },
      ],
    },
    {
      level: 0,
      runs: [
        { text: '• ' },
        {
          bold: true,
          color: '6366F1',
          hyperlink: { tooltip: 'Inspect NPM Package Registry', url: 'https://www.npmjs.com/package/@hokkyss/pptx' },
          text: 'NPM Package Registry (@hokkyss/pptx)',
          underline: true,
        },
        { text: ' — Published release.' },
      ],
    },
    {
      level: 0,
      runs: [
        { text: '• ' },
        {
          bold: true,
          color: '10B981',
          hyperlink: { tooltip: 'View Issue Tracker & Discussions', url: 'https://github.com/hokkyss/pptx-parser/issues' },
          text: 'GitHub Issue Tracker & Community',
          underline: true,
        },
        { text: ' — Discussions & issues.' },
      ],
    },
  ], {
    h: inches(1.8),
    w: inches(5.3),
    x: inches(7.0),
    y: inches(2.6),
  });

  // Right Bottom Card: Slide Show Action Controls
  slide12.addShape('roundRect', {
    fill: 'FFFFFF',
    h: inches(2.1),
    line: { color: 'E2E8F0', width: inches(0.015) },
    shadow: { blur: inches(0.12), color: '0F172A', opacity: 0.08 },
    w: inches(5.7),
    x: inches(6.8),
    y: inches(4.8),
  });

  slide12.addText('🎮 Native Slide Show Action Controls', {
    bold: true,
    color: '0F172A',
    fontSize: points(13),
    h: inches(0.35),
    w: inches(5.3),
    x: inches(7.0),
    y: inches(5.0),
  });

  slide12.addText('OpenXML action jump commands executed directly inside PowerPoint Presentation Mode:', {
    color: '64748B',
    fontSize: points(10),
    h: inches(0.3),
    w: inches(5.3),
    x: inches(7.0),
    y: inches(5.35),
  });

  const actionButtons = [
    { action: 'firstSlide', fill: '0F172A', label: '⏮ First', tip: 'Jump to Slide 1', x: 7.0 },
    { action: 'previousSlide', fill: '475569', label: '◀ Prev', tip: 'Jump to Slide 11', x: 8.35 },
    { action: 'nextSlide', fill: '0284C7', label: 'Next ▶', tip: 'Jump to Slide 13', x: 9.7 },
    { action: 'endShow', fill: 'EF4444', label: '⏹ End', tip: 'End Slide Show', x: 11.05 },
  ];

  for (const btn of actionButtons) {
    slide12.addShape('roundRect', {
      fill: btn.fill,
      h: inches(0.65),
      hyperlink: { action: btn.action, tooltip: btn.tip },
      text: btn.label,
      textOptions: { align: 'center', bold: true, color: 'FFFFFF', fontSize: points(10.5) },
      w: inches(1.2),
      x: inches(btn.x),
      y: inches(5.95),
    });
  }

  slide12.setNotes([
    {
      runs: [
        { text: 'Hyperlinks & Slide Navigation Keynote: ', bold: true },
        { text: 'Interactive OpenXML Actions & Relationships', underline: true },
      ],
    },
    { bullet: true, level: 0, text: 'Full support for external web links, custom URLs, and hover tooltips.' },
    { bullet: true, level: 0, text: 'Internal slide jumping using strongly typed slideIndex referencing.' },
    { bullet: true, level: 0, text: 'Built-in slide show control actions: firstSlide, lastSlide, nextSlide, previousSlide, endShow.' },
  ]);

  addDisclaimer(slide12, false);

  // ==================================================================
  // SLIDE 13: Native OpenXML DrawingML Gradient Engine
  // ==================================================================
  console.log('🎨 Generating Slide 13: Native OpenXML DrawingML Gradient Engine...');
  const slide13 = pres.addSlide();
  slide13.setTransition('wipe', { direction: 'right', speed: 'fast' });
  slide13.setBackground({
    angle: degrees(135),
    stops: ['090D16', '0F172A', '1E293B'],
    type: 'linear',
  });

  slide13.addText('Native OpenXML DrawingML Gradient Engine', {
    bold: true,
    color: 'FFFFFF',
    font: '+mj-lt',
    fontSize: points(24),
    h: inches(0.6),
    w: inches(11.73),
    x: inches(0.8),
    y: inches(0.8),
  });

  slide13.addText('Linear, radial, multi-stop vector meshes, alpha transparency, and angle rotation', {
    color: '94A3B8',
    fontSize: points(14),
    h: inches(0.4),
    w: inches(11.73),
    x: inches(0.8),
    y: inches(1.4),
  });

  // Top Nav Buttons
  slide13.addShape('roundRect', {
    fill: '0284C7',
    h: inches(0.45),
    hyperlink: { slideIndex: 12, tooltip: 'Return to Hyperlinks Hub (Slide 12)' },
    line: { color: '38BDF8', width: inches(0.015) },
    shadow: { blur: inches(0.12), color: '0284C7', opacity: 0.3 },
    text: '⬅ Back to Nav Hub',
    textOptions: { align: 'center', bold: true, color: 'FFFFFF', fontSize: points(10) },
    w: inches(2.2),
    x: inches(8.0),
    y: inches(0.8),
  });

  slide13.addShape('roundRect', {
    fill: {
      angle: degrees(90),
      stops: ['10B981', '059669'],
    },
    h: inches(0.45),
    hyperlink: { slideIndex: 14, tooltip: 'Advance to Live Benchmarks (Slide 14)' },
    line: { color: '34D399', width: inches(0.015) },
    shadow: { blur: inches(0.12), color: '10B981', opacity: 0.3 },
    text: 'Benchmarks ➔',
    textOptions: { align: 'center', bold: true, color: 'FFFFFF', fontSize: points(10) },
    w: inches(2.0),
    x: inches(10.4),
    y: inches(0.8),
  });

  // Left Card: Linear Vector Gradients
  slide13.addShape('roundRect', {
    fill: '0F172A',
    h: inches(4.9),
    line: { color: '334155', width: inches(0.015) },
    shadow: { blur: inches(0.12), color: '000000', opacity: 0.4 },
    w: inches(5.7),
    x: inches(0.8),
    y: inches(2.0),
  });

  slide13.addText('🌈 Multi-Angle Linear Gradient Topologies', {
    bold: true,
    color: 'FFFFFF',
    fontSize: points(13),
    h: inches(0.35),
    w: inches(5.3),
    x: inches(1.0),
    y: inches(2.2),
  });

  const linearPills = [
    {
      angle: 45,
      desc: '45° Diagonal • Sky to Indigo Vector Fusion',
      stops: ['0284C7', '38BDF8', '818CF8'],
      title: 'Oceanic Aurora Mesh',
      y: 2.65,
    },
    {
      angle: 90,
      desc: '90° Vertical • Amber to Crimson Sunset',
      stops: ['F59E0B', 'EF4444', 'EC4899'],
      title: 'Sunset Ember Gradient',
      y: 3.85,
    },
    {
      angle: 135,
      desc: '135° Diagonal • Forest to Emerald Glow',
      stops: ['059669', '10B981', '34D399'],
      title: 'Emerald Surge Vector',
      y: 5.05,
    },
  ];

  for (const pill of linearPills) {
    slide13.addShape('roundRect', {
      fill: {
        angle: degrees(pill.angle),
        stops: pill.stops,
        type: 'linear',
      },
      h: inches(1.05),
      line: { color: 'FFFFFF', width: inches(0.01) },
      shadow: { blur: inches(0.1), color: pill.stops[0], opacity: 0.35 },
      text: [
        { bold: true, color: 'FFFFFF', fontSize: points(12), text: `${pill.title}\n` },
        { color: 'F1F5F9', fontSize: points(9.5), text: pill.desc },
      ],
      textOptions: { align: 'left' },
      w: inches(5.3),
      x: inches(1.0),
      y: inches(pill.y),
    });
  }

  // Right Card: Radial & Alpha Glass Meshes
  slide13.addShape('roundRect', {
    fill: '0F172A',
    h: inches(4.9),
    line: { color: '334155', width: inches(0.015) },
    shadow: { blur: inches(0.12), color: '000000', opacity: 0.4 },
    w: inches(5.7),
    x: inches(6.8),
    y: inches(2.0),
  });

  slide13.addText('🔮 Radial Glow & Alpha Opacity Transparency', {
    bold: true,
    color: 'FFFFFF',
    fontSize: points(13),
    h: inches(0.35),
    w: inches(5.3),
    x: inches(7.0),
    y: inches(2.2),
  });

  // Radial Circle Card
  slide13.addShape('ellipse', {
    fill: {
      stops: [
        { color: '38BDF8', position: 0 },
        { color: '0284C7', position: 0.5 },
        { color: '0F172A', position: 1 },
      ],
      type: 'radial',
    },
    h: inches(1.6),
    line: { color: '38BDF8', width: inches(0.015) },
    shadow: { blur: inches(0.2), color: '38BDF8', opacity: 0.4 },
    w: inches(1.6),
    x: inches(7.1),
    y: inches(2.65),
  });

  slide13.addText('Radial Orb Focus\n3-Stop Centered Radial Gradient (<a:path path="circle">) with automatic edge falloff', {
    color: 'E2E8F0',
    fontSize: points(10),
    h: inches(1.4),
    w: inches(3.5),
    x: inches(8.9),
    y: inches(2.75),
  });

  // Alpha Transparency Frosted Glass Card
  slide13.addShape('roundRect', {
    fill: {
      angle: degrees(135),
      stops: [
        { color: 'FFFFFF', opacity: 0.25, position: 0 },
        { color: '6366F1', opacity: 0.15, position: 0.5 },
        { color: '0F172A', opacity: 0.4, position: 1 },
      ],
      type: 'linear',
    },
    h: inches(1.05),
    line: { color: '818CF8', width: inches(0.015) },
    shadow: { blur: inches(0.12), color: '6366F1', opacity: 0.25 },
    text: [
      { bold: true, color: 'FFFFFF', fontSize: points(11), text: '🪟 Alpha Transparency & Frosted Glass\n' },
      { color: 'CBD5E1', fontSize: points(9.5), text: 'Multi-stop alpha stops (10% to 80% opacity) for glassmorphic card overlays' },
    ],
    w: inches(5.3),
    x: inches(7.0),
    y: inches(4.4),
  });

  // Fluent API Code Snippet
  slide13.addShape('roundRect', {
    fill: '020617',
    h: inches(1.15),
    line: { color: '1E293B', width: inches(0.015) },
    text: [
      { bold: true, color: '38BDF8', fontSize: points(9), text: '// 3-Line Fluent Gradient Fill\n' },
      { color: 'E2E8F0', fontSize: points(9), text: 'slide.addShape(\'roundRect\', {\n  fill: { type: \'linear\', angle: degrees(135), stops: [\'#0284C7\', \'#6366F1\'] }\n});' },
    ],
    w: inches(5.3),
    x: inches(7.0),
    y: inches(5.55),
  });

  slide13.setNotes([
    {
      runs: [
        { text: 'DrawingML Gradient Engine Keynote: ', bold: true },
        { text: 'Full Vector Fidelity & Transparency', underline: true },
      ],
    },
    { bullet: true, level: 0, text: 'Native <a:gradFill> support with linear, radial, and path gradients.' },
    { bullet: true, level: 0, text: 'Full alpha transparency channel support across all color stops.' },
    { bullet: true, level: 0, text: 'Seamless integration with shapes, cards, outlines, and slide backgrounds.' },
  ]);

  addDisclaimer(slide13, true);

  // ==================================================================
  // SLIDE 14: Live Benchmark Matrix (Performance & Bundle Size)
  // ==================================================================
  console.log('⚡ Generating Slide 14: Live Benchmark Matrix (Performance & Bundle Size)...');
  const slide14 = pres.addSlide();
  slide14.setTransition('zoom', { direction: 'in', speed: 'med' });
  slide14.setBackground('0F172A');

  slide14.addText('Engine Performance & Bundle Size Matrix', {
    bold: true,
    color: 'FFFFFF',
    font: '+mj-lt',
    fontSize: points(24),
    h: inches(0.6),
    w: inches(11.73),
    x: inches(0.8),
    y: inches(0.8),
  });

  slide14.addText('Empirical microsecond throughput benchmarks and ultra-lean tree-shakeable bundle budgets', {
    color: '94A3B8',
    fontSize: points(14),
    h: inches(0.4),
    w: inches(11.73),
    x: inches(0.8),
    y: inches(1.4),
  });

  // Table Card (Left)
  slide14.addShape('roundRect', {
    fill: '1E293B',
    h: inches(4.9),
    line: { color: '334155', width: inches(0.015) },
    shadow: { blur: inches(0.12), color: '000000', opacity: 0.3 },
    w: inches(5.7),
    x: inches(0.8),
    y: inches(2.0),
  });

  slide14.addText('⚡ Throughput & Latency Benchmarks (Vitest)', {
    bold: true,
    color: 'FFFFFF',
    fontSize: points(13),
    h: inches(0.35),
    w: inches(5.3),
    x: inches(1.0),
    y: inches(2.2),
  });

  slide14.addTable(
    (tbl) => {
      // Header Row
      tbl.addRow({
        cells: [
          { bold: true, color: 'FFFFFF', fill: '0284C7', fontSize: points(10), text: 'Workload Scenario' },
          { bold: true, color: 'FFFFFF', fill: '0284C7', fontSize: points(10), text: 'Ops / Sec', align: 'right' },
          { bold: true, color: 'FFFFFF', fill: '0284C7', fontSize: points(10), text: 'Avg Latency', align: 'right' },
          { bold: true, color: 'FFFFFF', fill: '0284C7', fontSize: points(10), text: 'Category', align: 'center' },
        ],
        h: inches(0.38),
      });

      const rows = [
        { scenario: 'Single Slide Emit', ops: '1,820 ops/s', lat: '0.55 ms', badge: 'Sub-ms' },
        { scenario: 'DrawingML Chart', ops: '1,243 ops/s', lat: '0.80 ms', badge: 'Charts' },
        { scenario: 'Data Table (5x10)', ops: '1,270 ops/s', lat: '0.79 ms', badge: 'Tables' },
        { scenario: '10-Slide Full Deck', ops: '358 ops/s', lat: '2.80 ms', badge: 'Decks' },
        { scenario: '50-Slide High Batch', ops: '165 ops/s', lat: '6.07 ms', badge: 'Batch Scale' },
        { scenario: 'Shape Serialization', ops: '5.07M ops/s', lat: '0.0002 ms', badge: 'Microsecond' },
      ];

      rows.forEach((r, idx) => {
        const bg = idx % 2 === 0 ? '1E293B' : '0F172A';
        tbl.addRow({
          cells: [
            { bold: true, color: 'FFFFFF', fill: bg, fontSize: points(9.5), text: r.scenario },
            { bold: true, color: '38BDF8', fill: bg, fontSize: points(9.5), text: r.ops, align: 'right' },
            { bold: true, color: '10B981', fill: bg, fontSize: points(9.5), text: r.lat, align: 'right' },
            { bold: true, color: '94A3B8', fill: bg, fontSize: points(9), text: r.badge, align: 'center' },
          ],
          h: inches(0.36),
        });
      });
    },
    {
      colWidths: [inches(2.1), inches(1.1), inches(1.0), inches(1.0)],
      w: inches(5.2),
      x: inches(1.0),
      y: inches(2.65),
    },
  );

  // Bundle Size Card (Right)
  slide14.addShape('roundRect', {
    fill: '1E293B',
    h: inches(4.9),
    line: { color: '334155', width: inches(0.015) },
    shadow: { blur: inches(0.12), color: '000000', opacity: 0.3 },
    w: inches(5.7),
    x: inches(6.8),
    y: inches(2.0),
  });

  slide14.addText('📦 Monorepo Gzip & Brotli Footprint (Bundlephobia-Aligned)', {
    bold: true,
    color: 'FFFFFF',
    fontSize: points(13),
    h: inches(0.35),
    w: inches(5.3),
    x: inches(7.0),
    y: inches(2.2),
  });

  slide14.addChart({
    axisColor: '94A3B8',
    categories: ['@pptx-core', '@pptx-reader', '@pptx-writer', 'Full SDK'],
    chartType: 'column',
    h: inches(3.9),
    legend: { position: 'bottom' },
    series: [
      {
        color: '38BDF8',
        name: 'Gzip Size (KB)',
        values: [0.44, 25.83, 31.69, 48.93],
      },
      {
        color: '10B981',
        name: 'Brotli Size (KB)',
        values: [0.38, 22.89, 27.74, 43.32],
      },
    ],
    showGridlines: true,
    textColor: 'CBD5E1',
    w: inches(5.3),
    x: inches(7.0),
    y: inches(2.65),
  });

  slide14.addShape('roundRect', {
    fill: '0284C7',
    h: inches(0.45),
    hyperlink: { slideIndex: 12, tooltip: 'Return to Hyperlinks Hub (Slide 12)' },
    line: { color: '38BDF8', width: inches(0.015) },
    shadow: { blur: inches(0.15), color: '0284C7', direction: degrees(90), distance: inches(0.03), opacity: 0.3 },
    text: '⬅ Back to Nav Hub',
    textOptions: { align: 'center', bold: true, color: 'FFFFFF', fontSize: points(10) },
    w: inches(2.4),
    x: inches(10.0),
    y: inches(0.8),
  });

  slide14.setNotes([
    {
      runs: [
        { text: 'Empirical Benchmark Keynote: ', bold: true },
        { text: 'Sub-Millisecond Speed & Zero Native Dependencies', underline: true },
      ],
    },
    { bullet: true, level: 0, text: 'Single slide emits in ~0.55ms (> 1,800 ops/sec).' },
    { bullet: true, level: 0, text: 'Shape serialization exceeds 5 Million ops/sec.' },
    { bullet: true, level: 0, text: 'Total bundle size under 49 KB gzipped with full PresentationML & DrawingML engines.' },
  ]);

  addDisclaimer(slide14, true);

  // 3. Export Presentation to Buffer and Save to Disk
  const outputFile = resolve(process.cwd(), 'hokkyss_showcase_deck.pptx');
  const buffer = await pres.toBuffer();
  writeFileSync(outputFile, buffer);
  const elapsed = (performance.now() - t0).toFixed(2);

  console.log(`\n==================================================================`);
  console.log(`✓ Showcase generated successfully with ${pres.slides.length} slides in ${elapsed} ms!`);
  console.log(`💾 Saved to: ${outputFile}`);
  console.log(`==================================================================\n`);
}

runShowcase().catch(console.error);




