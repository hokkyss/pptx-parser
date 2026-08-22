import {
  BookOpenIcon,
  CodeIcon,
  CubeIcon,
  HouseIcon,
  PlayIcon,
  SparkleIcon,
} from '@phosphor-icons/react';
import { Link, useRouterState } from '@tanstack/react-router';
import { useEffect, useRef, useState } from 'react';

interface NavItem {
  icon: typeof HouseIcon;
  label: string;
  matchPattern?: RegExp;
  params?: Record<string, string>;
  to: string;
}

const NAV_ITEMS: NavItem[] = [
  { icon: HouseIcon, label: 'Home', to: '/' },
  {
    icon: BookOpenIcon,
    label: 'Docs',
    matchPattern: /^\/docs/,
    params: { _splat: 'getting-started/overview' },
    to: '/docs/$',
  },
  { icon: SparkleIcon, label: 'Cookbook', matchPattern: /^\/cookbook/, to: '/cookbook' },
  {
    icon: CodeIcon,
    label: 'API',
    matchPattern: /^\/api-reference/,
    params: { _splat: 'pptx/presentation' },
    to: '/api-reference/$',
  },
  { icon: PlayIcon, label: 'Playground', matchPattern: /^\/playground/, to: '/playground' },
  { icon: CubeIcon, label: 'Showcase', matchPattern: /^\/showcase/, to: '/showcase' },
];

/**
 * Ultra-translucent iOS-style liquid glass floating bottom navigation dock with directional sliding active pill.
 * @returns React node
 */
export default function MobileBottomNav() {
  const routerState = useRouterState();
  const currentPath = routerState.location.pathname;

  const activeIndex = NAV_ITEMS.findIndex((item) =>
    item.matchPattern ? item.matchPattern.test(currentPath) : currentPath === item.to,
  );

  const navContainerRef = useRef<HTMLDivElement>(null);
  const itemRefs = useRef<(HTMLAnchorElement | null)[]>([]);
  const [indicatorStyle, setIndicatorStyle] = useState<{ isReady: boolean; left: number; width: number }>({
    isReady: false,
    left: 0,
    width: 0,
  });

  useEffect(() => {
    const updateIndicator = () => {
      const targetIndex = activeIndex >= 0 ? activeIndex : 0;
      const currentEl = itemRefs.current[targetIndex];
      const containerEl = navContainerRef.current;
      if (currentEl && containerEl) {
        const containerRect = containerEl.getBoundingClientRect();
        const itemRect = currentEl.getBoundingClientRect();
        setIndicatorStyle({
          isReady: true,
          left: itemRect.left - containerRect.left,
          width: itemRect.width,
        });
      }
    };

    updateIndicator();
    window.addEventListener('resize', updateIndicator);
    return () => window.removeEventListener('resize', updateIndicator);
  }, [activeIndex]);

  return (
    <nav
      aria-label="Mobile navigation"
      className="fixed bottom-4 inset-x-3 sm:inset-x-8 max-w-md mx-auto z-40 md:hidden pointer-events-none"
    >
      <div
        className="relative pointer-events-auto flex items-center justify-around gap-1 px-2 py-1.5 rounded-full backdrop-blur-3xl backdrop-saturate-150 bg-white/15 dark:bg-black/25 border border-white/50 dark:border-white/10 shadow-[0_12px_40px_-8px_rgba(0,0,0,0.12),inset_0_1px_1px_0_rgba(255,255,255,0.7),inset_0_-1px_1px_0_rgba(0,0,0,0.03)] dark:shadow-[0_12px_40px_-8px_rgba(0,0,0,0.6),inset_0_1px_1px_0_rgba(255,255,255,0.15),inset_0_-1px_1px_0_rgba(0,0,0,0.3)] transition-all duration-300 ring-1 ring-black/5 dark:ring-white/5"
        ref={navContainerRef}
        style={{
          paddingBottom: 'calc(0.375rem + env(safe-area-inset-bottom, 0px))',
        }}
      >
        {/* Animated Directional Sliding Liquid Glass Pill Indicator */}
        <div
          aria-hidden="true"
          className={`absolute top-1.5 bottom-1.5 rounded-full bg-black/5 dark:bg-white/10 backdrop-blur-md border border-white/50 dark:border-white/15 shadow-[inset_0_1px_1px_0_rgba(255,255,255,0.3),0_2px_8px_0_rgba(0,0,0,0.03)] pointer-events-none z-0 ${
            indicatorStyle.isReady
              ? 'transition-all duration-350 ease-[cubic-bezier(0.32,0.72,0,1)] opacity-100'
              : 'opacity-0'
          }`}
          style={{
            left: `${indicatorStyle.left}px`,
            width: `${indicatorStyle.width}px`,
          }}
        />

        {NAV_ITEMS.map((item, idx) => {
          const Icon = item.icon;
          const isActive = idx === activeIndex;

          return (
            <Link
              className="relative flex-1 flex flex-col items-center justify-center py-1.5 px-1 rounded-full transition-all duration-200 z-10 hover:bg-white/10 dark:hover:bg-white/5 active:scale-95"
              key={item.label}
              params={item.params}
              ref={(el) => {
                itemRefs.current[idx] = el;
              }}
              to={item.to}
            >
              <Icon
                className={`h-4 w-4 shrink-0 transition-all duration-300 ${
                  isActive ? 'text-primary scale-110' : 'text-foreground/60'
                }`}
                weight={isActive ? 'fill' : 'regular'}
              />
              <span
                className={`text-[10px] leading-tight mt-0.5 tracking-tight transition-colors duration-200 ${
                  isActive ? 'text-primary font-bold' : 'text-foreground/60 font-medium'
                }`}
              >
                {item.label}
              </span>
            </Link>
          );
        })}
      </div>
    </nav>
  );
}
