import { Twitter } from './twitter.interface';

/**
 * The Twitter metadata for the document.
 * Used for configuring Twitter Cards and can include details such as `card`, `site`, and `creator`.
 * @param twitter twitter metadata configuration
 * @returns The twitter metadata configuration object.
 * @example
 * ```tsx
 * twitter: {
 *   card: "summary_large_image",
 *   site: "@site",
 *   creator: "@creator",
 *   images: "https://example.com/og.png"
 * }
 * ```
 */
export function defineTwitter(twitter: Twitter) {
  return twitter;
}
