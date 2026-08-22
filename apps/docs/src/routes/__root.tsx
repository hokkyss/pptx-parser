import type { ReactNode } from 'react';
import ThemeProvider from '@monorepo/design-system/application-theme-provider';
import Toaster from '@monorepo/design-system/toaster';
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
import vsCss from 'highlight.js/styles/vs.min.css?url';
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
    const [theme, err] = await tryit(ctx.context.queryClient.ensureQueryData(getApplicationThemeQuery()));
    if (err) throw err;
    return {
      theme,
    };
  },
  head: ({ loaderData }) => ({
    links: [
      { fetchPriority: 'high', href: appCss, rel: 'stylesheet' },
      { href: katexCss, rel: 'stylesheet' },
      { href: loaderData?.theme === 'dark' ? vsDarkCss : vsCss, rel: 'stylesheet' },
    ],
    meta: [
      { charSet: 'utf-8' },
      { content: 'width=device-width, initial-scale=1.0', name: 'viewport' },
      { title: '@hokkyss/pptx — Modern Isomorphic PowerPoint Engine for TypeScript' },
      { content: 'High-performance, isomorphic, and type-safe PowerPoint (.pptx) toolkit for Node.js, Web Browsers, and Cloudflare Workers.', name: 'description' },
    ],
  }),
  shellComponent: RootDocument,
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
