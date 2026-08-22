import { LightningIcon } from '@phosphor-icons/react';
import { ClientOnly, createFileRoute } from '@tanstack/react-router';
import { useState } from 'react';
import MonacoEditor from '../../components/monaco-editor.component';
import SlideCanvas from '../../components/slide-canvas.component';

export const Route = createFileRoute('/playground/')({
  component: PlaygroundPage,
});

const PRESETS: Record<string, { code: string; name: string }> = {
  chart: {
    code: `const pres = Presentation.create({ title: 'Sales Analytics' });
const slide = pres.addSlide();
slide.setBackground('0F172A');

slide.addText('Monthly Revenue Analytics', {
  x: inches(1), y: inches(0.6), w: inches(11.33), h: inches(0.8),
  fontSize: points(24), bold: true, color: '38BDF8'
});

slide.addChart({
  chartType: 'bar',
  title: 'Q1 - Q4 Performance',
  x: inches(1), y: inches(1.5), w: inches(11.33), h: inches(4.8),
  categories: ['Q1', 'Q2', 'Q3', 'Q4'],
  series: [
    { name: '2025 Revenue ($K)', values: [45, 58, 62, 80], color: '38BDF8' },
    { name: '2026 Target ($K)', values: [60, 75, 85, 110], color: '818CF8' }
  ]
});`,
    name: 'Interactive Chart Grid',
  },
  connectors: {
    code: `const pres = Presentation.create({ title: 'Microservice Pipeline' });
const slide = pres.addSlide();
slide.setBackground('0F172A');

slide.addText('Distributed Cloud Architecture', {
  x: inches(1), y: inches(0.8), w: inches(11.33), h: inches(0.8),
  fontSize: points(24), bold: true, color: '38BDF8', align: 'center'
});

// Source Service Node
slide.addShape('roundRect', {
  id: 'api-gateway',
  text: 'API Gateway',
  x: inches(1.5), y: inches(2.8), w: inches(3.5), h: inches(2),
  fill: '1E293B',
  line: { color: '38BDF8', width: inches(0.02) },
  textOptions: { color: '38BDF8', bold: true, fontSize: points(18), align: 'center' }
});

// Target Service Node
slide.addShape('roundRect', {
  id: 'auth-service',
  text: 'Auth Engine',
  x: inches(8.33), y: inches(2.8), w: inches(3.5), h: inches(2),
  fill: '1E293B',
  line: { color: '6366F1', width: inches(0.02) },
  textOptions: { color: '818CF8', bold: true, fontSize: points(18), align: 'center' }
});

// Direct Connector between nodes
slide.addConnector({
  from: { x: inches(5), y: inches(3.8) },
  to: { x: inches(8.33), y: inches(3.8) },
  color: '38BDF8',
  width: inches(0.03),
});`,
    name: 'Microservice Connectors Flow',
  },
  quickstart: {
    code: `const pres = Presentation.create({ title: 'Cloud Keynote' });
pres.setThemeColors({ accent1: '0284C7', dk1: '0F172A', lt1: 'FFFFFF' });

const slide = pres.addSlide();
slide.setBackground('0F172A');

slide.addText('⚡ Next-Gen Presentation Engine', {
  x: inches(1),
  y: inches(1.5),
  w: inches(11.33),
  h: inches(1.2),
  fontSize: points(36),
  bold: true,
  color: '38BDF8',
  align: 'center',
});

slide.addText('100% Isomorphic TypeScript OpenXML Compiler', {
  x: inches(1),
  y: inches(3.0),
  w: inches(11.33),
  h: inches(0.8),
  fontSize: points(18),
  color: '94A3B8',
  align: 'center',
});

slide.addShape('roundRect', {
  text: 'Explore Full Documentation →',
  x: inches(4.16),
  y: inches(4.5),
  w: inches(5),
  h: inches(1),
  fill: '0284C7',
  textOptions: { color: 'FFFFFF', bold: true, fontSize: points(16), align: 'center' }
});`,
    name: 'Quick Start Keynote',
  },
  table: {
    code: `const pres = Presentation.create({ title: 'Q4 Revenue Table' });
const slide = pres.addSlide();
slide.setBackground('0F172A');

slide.addText('Q4 Regional Performance', {
  x: inches(1), y: inches(0.8), w: inches(11.33), h: inches(0.8),
  fontSize: points(24), bold: true, color: '38BDF8'
});

slide.addTable(
  [
    ['Region', 'Q3 ($M)', 'Q4 ($M)', 'Growth YoY'],
    ['North America', '$120.4', '$158.9', '+32%'],
    ['EMEA', '$84.2', '$105.1', '+25%'],
    ['APAC', '$55.0', '$78.3', '+42%'],
  ],
  {
    x: inches(1), y: inches(1.8), w: inches(11.33), h: inches(3.5),
    colWidths: [inches(3.5), inches(2.5), inches(2.5), inches(2.83)],
    header: true,
  }
);`,
    name: 'Financial Data Grid',
  },
};

/**
 * Live interactive in-browser Monaco playground page component.
 * @returns React node
 */
function PlaygroundPage() {
  const [selectedPreset, setSelectedPreset] = useState('quickstart');
  const [code, setCode] = useState(PRESETS['quickstart'].code);

  const handleSelectPreset = (key: string) => {
    setSelectedPreset(key);
    setCode(PRESETS[key].code);
  };

  return (
    <div className="w-full max-w-[1800px] mx-auto px-4 sm:px-6 lg:px-8 py-4 sm:py-6 flex flex-col min-h-[calc(100vh-4rem)] lg:h-[calc(100vh-4rem)]">
      {/* Header Bar */}
      <div className="flex flex-wrap items-center justify-between gap-3 pb-4 border-b border-border/40 mb-4 sm:mb-6">
        <div>
          <h1 className="text-xl sm:text-2xl font-bold tracking-tight text-foreground flex items-center gap-2">
            <LightningIcon className="h-5 w-5 sm:h-6 sm:w-6 text-primary shrink-0" />
            <span>Live PowerPoint Playground</span>
          </h1>
          <p className="text-xs text-muted-foreground mt-0.5">
            Write TypeScript code and compile .pptx files in real time in your browser.
          </p>
        </div>

        <div className="flex items-center gap-2">
          <span className="text-xs font-semibold text-muted-foreground">Preset:</span>
          <select
            className="px-2.5 py-1.5 rounded-lg bg-background border border-input text-xs font-medium text-foreground focus:outline-none focus:ring-1 focus:ring-ring"
            onChange={(e) => handleSelectPreset(e.target.value)}
            value={selectedPreset}
          >
            {Object.entries(PRESETS).map(([key, item]) => (
              <option key={key} value={key}>{item.name}</option>
            ))}
          </select>
        </div>
      </div>

      {/* Split Pane Editor & Live Preview */}
      <div className="flex-1 grid grid-cols-1 lg:grid-cols-2 gap-6 min-h-0">
        <ClientOnly fallback={<div className="h-100 lg:h-full bg-muted rounded-lg animate-pulse border border-border" />}>
          <div className="flex flex-col h-100 sm:h-115 lg:h-full min-h-0">
            <div className="text-xs font-mono text-muted-foreground mb-2 flex items-center justify-between px-1">
              <span>TypeScript Editor (Monaco)</span>
              <span className="hidden sm:inline">@hokkyss/pptx in scope</span>
            </div>
            <div className="flex-1 min-h-0 rounded-xl overflow-hidden border border-border">
              <MonacoEditor onChange={setCode} value={code} />
            </div>
          </div>
        </ClientOnly>

        <ClientOnly fallback={<div className="h-87.5 lg:h-full bg-muted rounded-xl animate-pulse border border-border" />}>
          <div className="flex flex-col min-h-87.5 lg:h-full">
            <SlideCanvas code={code} />
          </div>
        </ClientOnly>
      </div>
    </div>
  );
}
