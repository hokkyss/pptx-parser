import { Link } from '@tanstack/react-router';
import type { ListDocsResponseDto } from '../lib/content/dto/list-docs.dto';

interface DocsSidebarProps {
  baseRoute?: '/api' | '/docs';
  sections: ListDocsResponseDto['sections'];
}

/**
 * Documentation navigation sidebar component.
 * @param root0 Component props
 * @param root0.baseRoute Base URL route prefix (e.g. /docs or /api)
 * @param root0.sections Grouped navigation sections
 * @returns React node
 */
export default function DocsSidebar({ baseRoute = '/docs', sections }: DocsSidebarProps) {
  return (
    <aside className="w-64 shrink-0 py-8 pr-6 border-r border-border/40 hidden lg:block overflow-y-auto max-h-[calc(100vh-4rem)] sticky top-16">
      <div className="space-y-6">
        {sections.map((sec) => (
          <div className="space-y-2" key={sec.title}>
            <h5 className="text-xs font-bold uppercase tracking-wider text-muted-foreground px-2">
              {sec.title}
            </h5>
            <ul className="space-y-1">
              {sec.items.map((item) => (
                <li key={item.path}>
                  <Link
                    activeProps={{
                      className: 'bg-primary/10 text-primary font-semibold border-r-2 border-primary',
                    }}
                    className="block px-2.5 py-1.5 rounded-md text-xs text-muted-foreground hover:text-foreground hover:bg-muted/40 transition"
                    params={{ _splat: item.path.replace(/^docs\//, '').replace(/^api\//, '') }}
                    to={baseRoute === '/api' ? '/api/$' : '/docs/$'}
                  >
                    {item.title}
                  </Link>
                </li>
              ))}
            </ul>
          </div>
        ))}
      </div>
    </aside>
  );
}
