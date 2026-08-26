import YAML from 'yaml';

// Vite inlines all markdown files at build time as raw strings
const rawMarkdownFiles = import.meta.glob<string>('../../../content/**/*.md', {
  eager: true,
  import: 'default',
  query: '?raw',
});

export interface ParsedDoc {
  content: string;
  description?: string;
  frontmatter: Record<string, unknown>;
  order?: number;
  package?: string;
  path: string;
  slug: string;
  title: string;
  toc: TocItem[];
}

export interface TocItem {
  id: string;
  level: number;
  text: string;
}

/**
 * Returns all parsed documentation objects across all markdown files.
 * @returns Array of parsed documentation objects
 */
export function getAllDocs(): ParsedDoc[] {
  const docs: ParsedDoc[] = [];

  for (const [filePath, rawContent] of Object.entries(rawMarkdownFiles)) {
    docs.push(parseMarkdownFile(rawContent, filePath));
  }

  return docs;
}

/**
 * Looks up a documentation page by its relative content path.
 * @param docPath Relative content path (e.g. docs/getting-started/overview)
 * @returns Parsed document or null if not found
 */
export function getDocByPath(docPath: string): null | ParsedDoc {
  const normalized = docPath.replace(/^\//, '').replace(/\.md$/, '');
  const all = getAllDocs();
  return all.find((d) => d.path === normalized) || null;
}

/**
 * Parses raw markdown text and YAML frontmatter into a structured document model.
 * @param raw Raw markdown file content string
 * @param filePath Path of the file for slug fallback derivation
 * @returns Structured parsed doc
 */
export function parseMarkdownFile(raw: string, filePath: string): ParsedDoc {
  const match = raw.match(/^---\r?\n([\s\S]*?)\r?\n---\r?\n([\s\S]*)$/);
  let frontmatter: Record<string, unknown> = {};
  let content = raw;

  if (match) {
    frontmatter = (YAML.parse(match[1]) as Record<string, unknown>) || {};
    content = match[2].trim();
  }

  const cleanPath = filePath.replace(/^.*content\//, '').replace(/\.md$/, '');
  const title = (frontmatter.title as string) || cleanPath.split('/').pop() || 'Untitled';
  const description = frontmatter.description as string | undefined;
  const order = frontmatter.order as number | undefined;
  const pkg = frontmatter.package as string | undefined;

  // Clean leading redundant H1 heading if present in content
  const leadingH1Regex = /^#\s+(.+)\r?\n+/;
  const h1Match = content.match(leadingH1Regex);
  if (h1Match && h1Match[1].trim().toLowerCase() === title.trim().toLowerCase()) {
    content = content.replace(leadingH1Regex, '');
  }

  const toc = extractToc(content);

  return {
    content,
    description,
    frontmatter,
    order,
    package: pkg,
    path: cleanPath,
    slug: cleanPath,
    title,
    toc,
  };
}

/**
 * Extracts table of contents headings (h2-h4) from markdown content.
 * @param markdown Markdown body string
 * @returns Array of TOC heading items
 */
function extractToc(markdown: string): TocItem[] {
  const headingRegex = /^(#{2,4})\s+(.+)$/gm;
  const items: TocItem[] = [];
  let match: null | RegExpExecArray;

  while ((match = headingRegex.exec(markdown)) !== null) {
    const level = match[1].length;
    const text = match[2].trim().replace(/\\[^\\]*\\]/g, '').replace(/[`*_~]/g, '');
    const id = text
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, '-')
      .replace(/(^-|-$)/g, '');

    items.push({ id, level, text });
  }

  return items;
}
