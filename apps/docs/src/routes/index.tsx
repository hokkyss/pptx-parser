import Skeleton from '@monorepo/design-system/skeleton';
import { ArrowRightIcon, LightningIcon, SparkleIcon } from '@phosphor-icons/react';
import { useSuspenseQuery } from '@tanstack/react-query';
import { Await, createFileRoute, Link, useLoaderData } from '@tanstack/react-router';
import { Suspense } from 'react';
import renderTeaserQuery from '../lib/content/queries/render-teaser.query';

export const Route = createFileRoute('/')({
  component: HomePage,
  loader: (ctx) => {
    return {
      teaserPromise: ctx.context.queryClient.ensureQueryData(renderTeaserQuery()),
    };
  },
});

/**
 * Home landing page component with hero and deferred code teaser.
 * @returns React node
 */
function HomePage() {
  const teaserPromise = useLoaderData({
    from: '/',
    select: (d) => d.teaserPromise,
  });

  return (
    <div className="flex flex-col items-center">
      {/* Hero Section */}
      <section className="w-full max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 pt-10 sm:pt-16 lg:pt-20 pb-10 sm:pb-16 text-center">
        <div className="inline-flex items-center gap-2 px-3.5 py-1.5 rounded-full bg-primary/10 border border-primary/20 text-primary text-xs font-semibold mb-6 shadow-sm">
          <SparkleIcon className="h-4 w-4 shrink-0" />
          <span className="truncate">Next-Generation Isomorphic PowerPoint Engine</span>
        </div>

        <h1 className="text-3xl sm:text-5xl lg:text-6xl font-extrabold tracking-tight text-foreground max-w-4xl mx-auto leading-tight">
          Parse, Construct & Mutate
          {' '}
          <span className="text-primary">PowerPoint (.pptx)</span>
          {' '}
          in Pure TypeScript
        </h1>

        <p className="mt-4 sm:mt-6 text-base sm:text-lg lg:text-xl text-muted-foreground max-w-2xl mx-auto leading-relaxed">
          100% Isomorphic. Zero native C++ binaries. Runs sub-millisecond fast across Node.js, Web Browsers, Cloudflare Workers, and Bun.
        </p>

        <div className="mt-8 flex flex-col sm:flex-row items-center justify-center gap-3 sm:gap-4 max-w-md sm:max-w-none mx-auto w-full">
          <Link
            className="w-full sm:w-auto inline-flex items-center justify-center gap-2 px-6 py-3.5 rounded-xl bg-primary hover:bg-primary/90 text-primary-foreground font-semibold text-sm shadow-lg shadow-primary/20 transition active:scale-95 text-center"
            params={{ _splat: 'getting-started/quickstart' }}
            to="/docs/$"
          >
            <span>Get Started in 30 Seconds</span>
            <ArrowRightIcon className="h-4 w-4" />
          </Link>
          <Link
            className="w-full sm:w-auto inline-flex items-center justify-center gap-2 px-6 py-3.5 rounded-xl bg-secondary hover:bg-secondary/80 text-secondary-foreground font-semibold text-sm border border-border shadow-sm transition active:scale-95 text-center"
            to="/playground"
          >
            <LightningIcon className="h-4 w-4 text-primary" />
            <span>Try Live Playground</span>
          </Link>
        </div>

        {/* Feature Pill Grid */}
        <div className="mt-12 sm:mt-16 grid grid-cols-2 sm:grid-cols-2 lg:grid-cols-4 gap-3 sm:gap-4 max-w-5xl mx-auto text-left">
          <div className="p-4 sm:p-5 rounded-2xl bg-card text-card-foreground border border-border shadow-sm hover:border-primary/40 transition">
            <span className="text-xl sm:text-2xl font-bold text-primary font-mono">190 KB</span>
            <p className="text-xs text-muted-foreground mt-1">Lightweight & tree-shakeable</p>
          </div>
          <div className="p-4 sm:p-5 rounded-2xl bg-card text-card-foreground border border-border shadow-sm hover:border-primary/40 transition">
            <span className="text-xl sm:text-2xl font-bold text-primary font-mono">0.10 ms</span>
            <p className="text-xs text-muted-foreground mt-1">Sub-millisecond parse latency</p>
          </div>
          <div className="p-4 sm:p-5 rounded-2xl bg-card text-card-foreground border border-border shadow-sm hover:border-primary/40 transition">
            <span className="text-xl sm:text-2xl font-bold text-primary font-mono">100%</span>
            <p className="text-xs text-muted-foreground mt-1">Round-trip AST fidelity</p>
          </div>
          <div className="p-4 sm:p-5 rounded-2xl bg-card text-card-foreground border border-border shadow-sm hover:border-primary/40 transition">
            <span className="text-xl sm:text-2xl font-bold text-primary font-mono">0 Native</span>
            <p className="text-xs text-muted-foreground mt-1">Pure TypeScript & fflate</p>
          </div>
        </div>
      </section>

      {/* Code Teaser Section */}
      <section className="w-full max-w-6xl mx-auto px-4 sm:px-6 lg:px-8 pb-16 sm:pb-24">
        <Suspense fallback={<Skeleton className="h-96 w-full rounded-2xl" />}>
          <Await promise={teaserPromise}>
            {() => <Teaser />}
          </Await>
        </Suspense>
      </section>
    </div>
  );
}

/**
 * Teaser snippet component rendered from React Server Component query.
 * @returns Rendered React node
 */
function Teaser() {
  const { data: renderedTeaser } = useSuspenseQuery({
    ...renderTeaserQuery(),
  });

  return renderedTeaser;
}
