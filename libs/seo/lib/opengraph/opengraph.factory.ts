import type { OpenGraph } from './opengraph.interface';

/**
 * The Open Graph metadata for the document. Follows the Open Graph protocol to enrich link previews.
 * @param opengraph The opengraph configuration.
 * @returns The opengraph configuration.
 * @example
 * ```tsx
 * openGraph: {
 *   type: "website",
 *   url: "https://example.com",
 *   title: "My Website",
 *   description: "My Website Description",
 *   siteName: "My Website",
 *   images: [{ url: "https://example.com/og.png" }]
 * }
 * ```
 * @see https://ogp.me/
 */
export function defineOpenGraph(opengraph: OpenGraph) {
  return opengraph;
}
