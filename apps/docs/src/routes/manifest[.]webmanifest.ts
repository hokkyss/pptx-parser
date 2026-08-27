import { defineManifest, resolveManifest } from '@monorepo/seo/manifest';
import { createFileRoute } from '@tanstack/react-router';
import getOrigin from '../lib/common/utils/get-origin.util';

export const Route = createFileRoute('/manifest.webmanifest')({
  server: {
    handlers({ createHandlers }) {
      return createHandlers({
        GET: {
          handler: (ctx) => {
            const origin = getOrigin(ctx.request);

            const manifest = defineManifest({
              background_color: '#09090b',
              categories: ['developer tools', 'utilities', 'productivity', 'office'],
              description: 'High-performance, isomorphic, and type-safe PowerPoint (.pptx) toolkit for Node.js, Web Browsers, and Cloudflare Workers.',
              dir: 'ltr',
              display: 'standalone',
              icons: [
                {
                  sizes: '16x16 32x32 48x48',
                  src: '/favicon.ico',
                  type: 'image/x-icon',
                },
                {
                  sizes: 'any',
                  src: '/favicon.svg',
                  type: 'image/svg+xml',
                },
                {
                  sizes: '180x180',
                  src: '/apple-icon.png',
                  type: 'image/png',
                },
                {
                  purpose: 'any',
                  sizes: '512x512',
                  src: '/icon.png',
                  type: 'image/png',
                },
                {
                  purpose: 'maskable',
                  sizes: '512x512',
                  src: '/icon.png',
                  type: 'image/png',
                },
                {
                  sizes: '1200x630',
                  src: '/opengraph-image.png',
                  type: 'image/png',
                },
                {
                  sizes: '1200x630',
                  src: '/twitter-image.png',
                  type: 'image/png',
                },
              ],
              lang: 'en',
              name: '@hokkyss/pptx — TypeScript PowerPoint Engine',
              orientation: 'any',
              scope: origin,
              short_name: '@hokkyss/pptx',
              shortcuts: [
                {
                  description: 'Get started with @hokkyss/pptx in 30 seconds',
                  name: 'Quick Start',
                  short_name: 'Quickstart',
                  url: '/docs/getting-started/quickstart',
                },
                {
                  description: 'Explore API reference for @hokkyss/pptx',
                  name: 'API Reference',
                  short_name: 'API',
                  url: '/api-reference/pptx/presentation',
                },
                {
                  description: 'Production recipes and code patterns',
                  name: 'Cookbook',
                  short_name: 'Cookbook',
                  url: '/cookbook',
                },
                {
                  description: 'Visual gallery and 14-slide showcase',
                  name: 'Showcase Gallery',
                  short_name: 'Showcase',
                  url: '/showcase',
                },
              ],
              start_url: '/',
              theme_color: '#09090b',
            });

            return new Response(resolveManifest(manifest), {
              headers: {
                'Cache-Control': 'public, max-age=3600, s-maxage=86400',
                'Content-Type': 'application/manifest+json',
              },
            });
          },
        },
      });
    },
  },
});
