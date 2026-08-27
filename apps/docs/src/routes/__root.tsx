import type { ReactNode } from 'react';
import ThemeProvider from '@monorepo/design-system/application-theme-provider';
import Toaster from '@monorepo/design-system/toaster';
import { defineIcons, resolveIcons } from '@monorepo/seo/icons';
import { defineMetadata, resolveMetadata } from '@monorepo/seo/metadata';
import { defineOpenGraph, resolveOpenGraph } from '@monorepo/seo/opengraph';
import { defineTwitter, resolveTwitter } from '@monorepo/seo/twitter';
import { defineViewport, resolveViewport } from '@monorepo/seo/viewport';
import { tryit } from '@monorepo/utils';
import { TanStackDevtools } from '@tanstack/react-devtools';
import { type QueryClient, useSuspenseQuery } from '@tanstack/react-query';
import { ReactQueryDevtoolsPanel } from '@tanstack/react-query-devtools';
import {
  createRootRouteWithContext,
  HeadContent,
  Scripts,
} from '@tanstack/react-router';
import { TanStackRouterDevtoolsPanel } from '@tanstack/react-router-devtools';
import vsDarkCss from 'highlight.js/styles/vs-dark.min.css?url';
import katexCss from 'katex/dist/katex.min.css?url';
import DocsHeader from '../components/docs-header.component';
import MobileBottomNav from '../components/mobile-bottom-nav.component';
import getApplicationThemeQuery from '../lib/common/queries/get-application-theme.query';
import appCss from '../styles.css?url';

interface MyRouterContext {
  queryClient: QueryClient;
}

export const Route = createRootRouteWithContext<MyRouterContext>()({
  async loader(ctx) {
    const [, err] = await tryit(ctx.context.queryClient.ensureQueryData(getApplicationThemeQuery()));

    if (err) throw err;

    return {};
  },
  head: () => {
    const metaSeo = resolveMetadata(
      defineMetadata({
        applicationName: '@hokkyss/pptx',
        authors: [{ name: 'Hokki Suwanda', url: 'https://github.com/hokkyss' }],
        creator: 'Hokki Suwanda',
        description: 'High-performance, isomorphic, and type-safe PowerPoint (.pptx) toolkit for Node.js, Web Browsers, and Cloudflare Workers.',
        generator: 'TanStack Start',
        keywords: [
          'pptx',
          'powerpoint',
          'presentation',
          'typescript',
          'isomorphic',
          'openxml',
          'drawingml',
          'charts',
          'slide-builder',
          'cloudflare-workers',
          'nodejs',
          'browser',
          'bun',
          'deno',
        ],
        manifest: '/manifest.webmanifest',
        publisher: 'hokkyss',
        title: '@hokkyss/pptx — Modern Isomorphic PowerPoint Engine for TypeScript',
      }),
    );

    const ogSeo = resolveOpenGraph(
      defineOpenGraph({
        description: 'High-performance, isomorphic, and type-safe PowerPoint (.pptx) toolkit for Node.js, Web Browsers, and Cloudflare Workers.',
        images: [
          {
            alt: '@hokkyss/pptx — Modern Isomorphic PowerPoint Engine for TypeScript',
            height: 630,
            type: 'image/png',
            url: '/opengraph-image.png',
            width: 1200,
          },
        ],
        locale: 'en_US',
        siteName: '@hokkyss/pptx Documentation',
        title: '@hokkyss/pptx — Modern Isomorphic PowerPoint Engine for TypeScript',
        type: 'website',
      }),
    );

    const twitterSeo = resolveTwitter(
      defineTwitter({
        card: 'summary_large_image',
        creator: '@hokkyss',
        description: 'High-performance, isomorphic, and type-safe PowerPoint (.pptx) toolkit for Node.js, Web Browsers, and Cloudflare Workers.',
        images: [
          {
            alt: '@hokkyss/pptx — Modern Isomorphic PowerPoint Engine for TypeScript',
            url: '/twitter-image.png',
          },
        ],
        title: '@hokkyss/pptx — Modern Isomorphic PowerPoint Engine for TypeScript',
      }),
    );

    const iconsSeo = resolveIcons(
      defineIcons({
        apple: [
          {
            sizes: '180x180',
            type: 'image/png',
            url: '/apple-icon.png',
          },
        ],
        icon: [
          {
            sizes: 'any',
            type: 'image/x-icon',
            url: '/favicon.ico',
          },
          {
            sizes: 'any',
            type: 'image/svg+xml',
            url: '/favicon.svg',
          },
          {
            sizes: '512x512',
            type: 'image/png',
            url: '/icon.png',
          },
        ],
        shortcut: ['/favicon.ico'],
      }),
    );

    const viewportSeo = resolveViewport(
      defineViewport({
        initialScale: 1,
        themeColor: [
          { color: '#09090b', media: '(prefers-color-scheme: dark)' },
          { color: '#ffffff', media: '(prefers-color-scheme: light)' },
        ],
        width: 'device-width',
      }),
    );

    const jsonLd = {
      '@context': 'https://schema.org',
      '@type': 'SoftwareApplication',
      applicationCategory: 'DeveloperApplication',
      author: {
        '@type': 'Person',
        name: 'Hokki Suwanda',
        url: 'https://github.com/hokkyss',
      },
      description: 'High-performance, isomorphic, and type-safe PowerPoint (.pptx) toolkit for Node.js, Web Browsers, and Cloudflare Workers.',
      name: '@hokkyss/pptx',
      offers: {
        '@type': 'Offer',
        price: '0',
        priceCurrency: 'USD',
      },
      operatingSystem: 'Cross-platform (Node.js, Web Browsers, Cloudflare Workers, Bun, Deno)',
    };

    return {
      links: [
        { fetchPriority: 'high', href: appCss, rel: 'stylesheet' },
        { href: katexCss, rel: 'stylesheet' },
        { href: vsDarkCss, rel: 'stylesheet' },
        ...metaSeo.links,
        ...ogSeo.links,
        ...twitterSeo.links,
        ...iconsSeo.links,
        ...viewportSeo.links,
      ],
      meta: [
        { charSet: 'utf-8' },
        ...metaSeo.metas,
        ...ogSeo.metas,
        ...twitterSeo.metas,
        ...iconsSeo.metas,
        ...viewportSeo.metas,
      ],
      scripts: [
        {
          children: JSON.stringify(jsonLd),
          type: 'application/ld+json',
        },
      ],
    };
  },
  shellComponent: RootDocument,
  shouldReload: false,
});

/**
 * Root document shell component providing theme and layout providers.
 * @param root0 Component props
 * @param root0.children Child route elements
 * @returns React node
 */
function RootDocument({ children }: { children: ReactNode }) {
  const { data: theme } = useSuspenseQuery(getApplicationThemeQuery());

  return (
    <html className={theme} data-theme={theme} lang="en">
      <head>
        <HeadContent />
      </head>
      <body className="min-h-screen bg-background text-foreground antialiased flex flex-col font-sans">
        <ThemeProvider theme={theme}>
          <Toaster position="bottom-right" />
          <DocsHeader />
          <main className="flex-1 pb-20 md:pb-0">
            {children}
          </main>
          <MobileBottomNav />
          <footer className="border-t border-border/40 py-8 px-4 text-center text-xs text-muted-foreground pb-24 md:pb-8">
            <div className="max-w-7xl mx-auto flex flex-col sm:flex-row items-center justify-between gap-4">
              <span>MIT Licensed © 2026 hokkyss. Built for ultra-fast isomorphic presentation generation.</span>
              <div className="flex items-center gap-4">
                <a className="hover:text-foreground transition" href="https://github.com/hokkyss/pptx-parser" rel="noreferrer" target="_blank">
                  GitHub
                </a>
                <a className="hover:text-foreground transition" href="https://www.npmjs.com/package/@hokkyss/pptx" rel="noreferrer" target="_blank">
                  npm
                </a>
              </div>
            </div>
          </footer>
        </ThemeProvider>
        <TanStackDevtools
          config={{ position: 'bottom-right' }}
          plugins={[
            { name: 'Tanstack Router', render: <TanStackRouterDevtoolsPanel /> },
            { name: 'Tanstack Query', render: <ReactQueryDevtoolsPanel /> },
          ]}
        />
        <Scripts />
      </body>
    </html>
  );
}
