import { LightningIcon } from '@phosphor-icons/react';
import { ClientOnly, createFileRoute } from '@tanstack/react-router';
import { useState } from 'react';
import MonacoEditor from '../../components/monaco-editor.component';
import SlideCanvas from '../../components/slide-canvas.component';

export const Route = createFileRoute('/playground/')({
  component: PlaygroundPage,
});

const PRESETS: Record<string, { code: string; name: string }> = {
  connectors: {
    code: `const pres = Presentation.create({ title: 'Microservice Pipeline' });
const slide = pres.addSlide();
slide.setBackground('0F172A');

// Add source and target cards
slide.addShape('roundRect', {
  id: 'api-gateway',
  text: 'API Gateway',
  x: inches(1.5), y: inches(3), w: inches(3), h: inches(1.5),
  fill: '1E293B',
  line: { color: '38BDF8', width: inches(0.015) },
  textOptions: { color: '38BDF8', bold: true, align: 'center' }
});

slide.addShape('roundRect', {
  id: 'auth-service',
  text: 'Auth Engine',
  x: inches(6.5), y: inches(3), w: inches(3), h: inches(1.5),
  fill: '1E293B',
  line: { color: '6366F1', width: inches(0.015) },
  textOptions: { color: '818CF8', bold: true, align: 'center' }
});

// Attach connector glued between shapes
slide.addConnector({
  from: { shapeId: 'api-gateway', position: 'right' },
  to: { shapeId: 'auth-service', position: 'left' },
  color: '38BDF8',
  width: inches(0.02),
  endArrow: 'triangle',
});`,
    name: 'Microservice Connectors Flow',
  },
  quickstart: {
    code: `const pres = Presentation.create({ title: 'Cloud Keynote' });
pres.setThemeColors({ accent1: '#0284C7', dk1: '#0F172A', lt1: '#FFFFFF' });

const slide = pres.addSlide();
slide.setBackground('0F172A');

slide.addText('⚡ Next-Gen Presentation Engine', {
  x: inches(1),
  y: inches(1.5),
  w: inches(11.33),
  h: inches(1),
  fontSize: points(36),
  bold: true,
  color: '38BDF8',
  align: 'center',
});

slide.addText('100% Isomorphic TypeScript OpenXML Compiler', {
  x: inches(1),
  y: inches(2.8),
  w: inches(11.33),
  h: inches(0.8),
  fontSize: points(18),
  color: '94A3B8',
  align: 'center',
});`,
    name: 'Quick Start Keynote',
  },
  table: {
    code: `const pres = Presentation.create({ title: 'Q4 Revenue Table' });
const slide = pres.addSlide();

slide.addText('Q4 Regional Performance', {
  x: inches(1), y: inches(0.8), w: inches(11.33), h: inches(0.8),
  fontSize: points(24), bold: true, color: '0284C7'
});

slide.addTable({
  x: inches(1), y: inches(1.8), w: inches(11.33),
  columns: [{ width: inches(3.5) }, { width: inches(2.5) }, { width: inches(2.5) }, { width: inches(2.83) }],
  headers: ['Region', 'Q3 ($M)', 'Q4 ($M)', 'Growth YoY'],
  headerStyle: { background: '0284C7', color: 'FFFFFF', bold: true },
  rows: [
    ['North America', '$120.4', '$158.9', '+32%'],
    ['EMEA', '$84.2', '$105.1', '+25%'],
    ['APAC', '$55.0', '$78.3', '+42%'],
  ],
  alternateRowBackground: 'F1F5F9',
});`,
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
    <div className="w-full max-w-[1800px] mx-auto px-4 sm:px-6 lg:px-8 py-6 flex flex-col h-[calc(100vh-4rem)]">
      {/* Header Bar */}
      <div className="flex flex-wrap items-center justify-between gap-4 pb-4 border-b border-border/40 mb-6">
        <div>
          <h1 className="text-2xl font-bold tracking-tight text-foreground flex items-center gap-2">
            <LightningIcon className="h-6 w-6 text-amber-500" />
            Live PowerPoint Playground
          </h1>
          <p className="text-xs text-muted-foreground mt-0.5">
            Write TypeScript code and compile .pptx files in real time in your browser.
          </p>
        </div>

        <div className="flex items-center gap-2">
          <span className="text-xs font-semibold text-muted-foreground">Preset:</span>
          <select
            className="px-3 py-1.5 rounded-lg bg-background border border-input text-xs font-medium text-foreground focus:outline-none focus:ring-1 focus:ring-ring"
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
        <ClientOnly fallback={<div className="h-full bg-muted rounded-lg animate-pulse border border-border" />}>
          <div className="flex flex-col h-full">
            <div className="text-xs font-mono text-muted-foreground mb-2 flex items-center justify-between">
              <span>TypeScript Editor (Monaco)</span>
              <span>@hokkyss/pptx in scope</span>
            </div>
            <div className="flex-1 min-h-0">
              <MonacoEditor onChange={setCode} value={code} />
            </div>
          </div>
        </ClientOnly>

        <ClientOnly fallback={<div className="h-full bg-muted rounded-xl animate-pulse border border-border" />}>
          <SlideCanvas code={code} />
        </ClientOnly>
      </div>
    </div>
  );
}
