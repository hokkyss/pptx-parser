import { ListIcon, XIcon } from '@phosphor-icons/react';
import { Link } from '@tanstack/react-router';
import { useEffect, useState } from 'react';
import type { ListDocsResponseDto } from '../lib/content/dto/list-docs.dto';

interface MobileSidebarSheetProps {
  baseRoute?: '/api-reference' | '/docs';
  sections: ListDocsResponseDto['sections'];
  title?: string;
}

/**
 * Mobile sliding sheet for documentation and API reference sidebar navigation.
 * @param root0 Component props
 * @param root0.baseRoute Base URL path prefix
 * @param root0.sections Documentation section items
 * @param root0.title Optional title for the sheet header
 * @returns React node
 */
export default function MobileSidebarSheet({
  baseRoute = '/docs',
  sections,
  title = 'Navigation',
}: MobileSidebarSheetProps) {
  const [isOpen, setIsOpen] = useState(false);

  // Close on ESC key press
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') setIsOpen(false);
    };
    if (isOpen) {
      window.addEventListener('keydown', handleKeyDown);
      document.body.style.overflow = 'hidden';
    }
    return () => {
      window.removeEventListener('keydown', handleKeyDown);
      document.body.style.overflow = '';
    };
  }, [isOpen]);

  return (
    <div className="lg:hidden">
      {/* Trigger Button */}
      <button
        aria-expanded={isOpen}
        aria-label="Open documentation navigation"
        className="inline-flex items-center gap-2 px-3.5 py-1.5 rounded-full bg-muted/70 hover:bg-muted text-foreground border border-border/80 text-xs font-semibold shadow-sm transition active:scale-95"
        onClick={() => setIsOpen(true)}
        type="button"
      >
        <ListIcon className="h-4 w-4 text-primary" />
        <span>Menu</span>
      </button>

      {/* Backdrop & Drawer Portal */}
      {isOpen && (
        <div className="fixed inset-0 z-50 flex">
          {/* Backdrop */}
          <div
            aria-hidden="true"
            className="fixed inset-0 bg-black/60 backdrop-blur-sm transition-opacity animate-in fade-in duration-200"
            onClick={() => setIsOpen(false)}
          />

          {/* Drawer Panel */}
          <div
            aria-label="Sidebar navigation"
            aria-modal="true"
            className="relative flex flex-col w-4/5 max-w-xs bg-background text-foreground h-full shadow-2xl border-r border-border z-10 animate-in slide-in-from-left duration-300"
            role="dialog"
          >
            {/* Drawer Header */}
            <div className="flex items-center justify-between px-4 py-3.5 border-b border-border">
              <span className="font-bold text-sm tracking-tight text-foreground">
                {title}
              </span>
              <button
                aria-label="Close navigation"
                className="p-1 rounded-md text-muted-foreground hover:text-foreground hover:bg-muted transition"
                onClick={() => setIsOpen(false)}
                type="button"
              >
                <XIcon className="h-5 w-5" />
              </button>
            </div>

            {/* Scrollable Section Tree */}
            <div className="flex-1 overflow-y-auto p-4 space-y-6">
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
                          className="block px-2.5 py-2 rounded-md text-xs text-muted-foreground hover:text-foreground hover:bg-muted/40 transition"
                          onClick={() => setIsOpen(false)}
                          params={{ _splat: item.path.replace(/^docs\//, '').replace(/^api-reference\//, '') }}
                          to={baseRoute === '/api-reference' ? '/api-reference/$' : '/docs/$'}
                        >
                          {item.title}
                        </Link>
                      </li>
                    ))}
                  </ul>
                </div>
              ))}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
