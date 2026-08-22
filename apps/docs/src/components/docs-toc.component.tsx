import type { TocItem } from '../lib/content/content-manifest';

interface DocsTocProps {
  toc: TocItem[];
}

/**
 * Table of contents sidebar component showing heading navigation links.
 * @param root0 Component props
 * @param root0.toc Array of heading items
 * @returns React node or null
 */
export default function DocsToc({ toc }: DocsTocProps) {
  if (!toc || toc.length === 0) return null;

  return (
    <aside className="w-56 shrink-0 py-8 pl-6 border-l border-border/40 hidden xl:block overflow-y-auto max-h-[calc(100vh-4rem)] sticky top-16 text-xs">
      <h5 className="font-semibold text-foreground uppercase tracking-wider mb-3">
        On this page
      </h5>
      <ul className="space-y-2">
        {toc.map((item) => (
          <li
            key={item.id}
            style={{ paddingLeft: `${(item.level - 2) * 12}px` }}
          >
            <a
              className="text-muted-foreground hover:text-primary transition block leading-snug"
              href={`#${item.id}`}
            >
              {item.text}
            </a>
          </li>
        ))}
      </ul>
    </aside>
  );
}
