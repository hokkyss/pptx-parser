import { defineSitemap, resolveSitemap, type Sitemap } from '@monorepo/seo/sitemap';
import { createFileRoute } from '@tanstack/react-router';
import getOrigin from '../lib/common/utils/get-origin.util';
import { getAllDocs } from '../lib/content/content-manifest';

export const Route = createFileRoute('/sitemap.xml')({
  server: {
    handlers({ createHandlers }) {
      return createHandlers({
        GET: {
          handler: (ctx) => {
            const origin = getOrigin(ctx.request);
            const now = new Date().toISOString();
            const docs = getAllDocs();

            const sitemapEntries: Sitemap = [
              // 1. Root landing page
              {
                changeFrequency: 'daily',
                lastModified: now,
                priority: 1.0,
                url: origin,
              },
              // 2. Section landing pages
              {
                changeFrequency: 'weekly',
                lastModified: now,
                priority: 0.8,
                url: `${origin}/cookbook`,
              },
              // {
              //   changeFrequency: 'monthly',
              //   lastModified: now,
              //   priority: 0.7,
              //   url: `${origin}/showcase`,
              // },
            ];

            // 3. Documentation & API Reference & Cookbook Markdown Pages
            for (const doc of docs) {
              const path = doc.path;
              // Skip indexes that redirect or non-route aliases
              if (
                path === 'docs/index'
                || path === 'api-reference/index'
                || path === 'cookbook/index'
                || path === 'playground/index'
                || path === 'showcase/index'
                || path.startsWith('showcase/')
                || path.startsWith('playground/')
              ) {
                continue;
              }

              let priority = 0.8;
              const changeFrequency: Sitemap[number]['changeFrequency'] = 'weekly';

              if (path.startsWith('docs/getting-started/')) {
                priority = 0.9;
              } else if (path.startsWith('api-reference/')) {
                priority = 0.8;
              } else if (path.startsWith('cookbook/')) {
                priority = 0.7;
              }

              sitemapEntries.push({
                changeFrequency,
                lastModified: now,
                priority,
                url: `${origin}/${path}`,
              });
            }

            return new Response(resolveSitemap(defineSitemap(sitemapEntries)), {
              headers: {
                'Cache-Control': 'public, max-age=3600, s-maxage=86400',
                'Content-Type': 'application/xml; charset=utf-8',
              },
            });
          },
        },
      });
    },
  },
});
