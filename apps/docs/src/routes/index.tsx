import { ArrowRightIcon, LightningIcon, SparkleIcon } from '@phosphor-icons/react';
import { createFileRoute, Link } from '@tanstack/react-router';
import MarkdownRenderer from '../components/markdown-renderer.component';

export const Route = createFileRoute('/')({
  component: HomePage,
});

const CODE_TEASER_SNIPPET = `\`\`\`typescript
import { Presentation, inches, points } from '@hokkyss/pptx';

// 1. Initialize presentation
const pres = Presentation.create({ title: 'Cloud Architecture' });

// 2. Add slide with theme & multilevel bullets
const slide = pres.addSlide();
slide.setBackground('0F172A');

slide.addText('Distributed Edge Compute', {
  color: '38BDF8',
  fontSize: points(32),
  bold: true,
  x: inches(1), y: inches(1), w: inches(11), h: inches(0.8)
});

// 3. Attach connected cards
slide.addShape('roundRect', { id: 'gateway', text: 'Edge Gateway', x: inches(1), y: inches(2.5), w: inches(3), h: inches(1.5) });
slide.addShape('roundRect', { id: 'auth', text: 'Auth Service', x: inches(6), y: inches(2.5), w: inches(3), h: inches(1.5) });
slide.addConnector({ from: { shapeId: 'gateway', position: 'right' }, to: { shapeId: 'auth', position: 'left' }, endArrow: 'triangle' });

// 4. Save (or toArrayBuffer() in browser/workers)
await pres.save('deck.pptx');
\`\`\``;

/**
 * Home landing page component with hero and code teaser.
 * @returns React node
 */
function HomePage() {
  return (
    <div className="flex flex-col items-center">
      {/* Hero Section */}
      <section className="w-full max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 pt-20 pb-16 text-center">
        <div className="inline-flex items-center gap-2 px-3.5 py-1.5 rounded-full bg-primary/10 border border-primary/20 text-primary text-xs font-semibold mb-6 shadow-sm">
          <SparkleIcon className="h-4 w-4" />
          <span>Next-Generation Isomorphic PowerPoint Engine</span>
        </div>

        <h1 className="text-4xl sm:text-6xl font-extrabold tracking-tight text-foreground max-w-4xl mx-auto leading-tight">
          Parse, Construct & Mutate
          {' '}
          <span className="bg-linear-to-r from-sky-500 via-indigo-500 to-purple-500 bg-clip-text text-transparent">PowerPoint (.pptx)</span>
          {' '}
          in Pure TypeScript
        </h1>

        <p className="mt-6 text-lg sm:text-xl text-muted-foreground max-w-2xl mx-auto leading-relaxed">
          100% Isomorphic. Zero native C++ binaries. Runs sub-millisecond fast across Node.js, Web Browsers, Cloudflare Workers, and Bun.
        </p>

        <div className="mt-8 flex flex-wrap items-center justify-center gap-4">
          <Link
            className="inline-flex items-center gap-2 px-6 py-3.5 rounded-xl bg-primary hover:bg-primary/90 text-primary-foreground font-semibold text-sm shadow-lg shadow-primary/20 transition active:scale-95"
            params={{ _splat: 'getting-started/quickstart' }}
            to="/docs/$"
          >
            Get Started in 30 Seconds
            <ArrowRightIcon className="h-4 w-4" />
          </Link>
          <Link
            className="inline-flex items-center gap-2 px-6 py-3.5 rounded-xl bg-secondary hover:bg-secondary/80 text-secondary-foreground font-semibold text-sm border border-border shadow-sm transition active:scale-95"
            to="/playground"
          >
            <LightningIcon className="h-4 w-4 text-amber-500" />
            Try Live Playground
          </Link>
        </div>

        {/* Feature Pill Grid */}
        <div className="mt-16 grid grid-cols-2 sm:grid-cols-4 gap-4 max-w-5xl mx-auto text-left">
          <div className="p-5 rounded-2xl bg-card text-card-foreground border border-border shadow-sm hover:border-primary/40 transition">
            <span className="text-2xl font-bold text-sky-500 font-mono">190 KB</span>
            <p className="text-xs text-muted-foreground mt-1">Lightweight & tree-shakeable</p>
          </div>
          <div className="p-5 rounded-2xl bg-card text-card-foreground border border-border shadow-sm hover:border-primary/40 transition">
            <span className="text-2xl font-bold text-emerald-500 font-mono">0.10 ms</span>
            <p className="text-xs text-muted-foreground mt-1">Sub-millisecond parse latency</p>
          </div>
          <div className="p-5 rounded-2xl bg-card text-card-foreground border border-border shadow-sm hover:border-primary/40 transition">
            <span className="text-2xl font-bold text-purple-500 font-mono">100%</span>
            <p className="text-xs text-muted-foreground mt-1">Round-trip AST fidelity</p>
          </div>
          <div className="p-5 rounded-2xl bg-card text-card-foreground border border-border shadow-sm hover:border-primary/40 transition">
            <span className="text-2xl font-bold text-amber-500 font-mono">0 Native</span>
            <p className="text-xs text-muted-foreground mt-1">Pure TypeScript & fflate</p>
          </div>
        </div>
      </section>

      {/* Code Teaser Section */}
      <section className="w-full max-w-6xl mx-auto px-4 sm:px-6 lg:px-8 pb-24">
        <MarkdownRenderer content={CODE_TEASER_SNIPPET} />
      </section>
    </div>
  );
}
