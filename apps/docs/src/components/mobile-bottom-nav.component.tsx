import {
  BookOpenIcon,
  CodeIcon,
  CubeIcon,
  HouseIcon,
  PlayIcon,
  SparkleIcon,
} from '@phosphor-icons/react';
import { Link, useRouterState } from '@tanstack/react-router';

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
 * High-vibrancy iOS-style liquid glass floating bottom navigation dock with directional sliding active pill.
 * @returns React node
 */
export default function MobileBottomNav() {
  const routerState = useRouterState();
  const currentPath = routerState.location.pathname;

  const activeIndex = NAV_ITEMS.findIndex((item) =>
    item.matchPattern ? item.matchPattern.test(currentPath) : currentPath === item.to,
  );

  const safeIndex = activeIndex >= 0 ? activeIndex : 0;

  return (
    <nav
      aria-label="Mobile navigation"
      className="fixed bottom-4 inset-x-3 sm:inset-x-8 max-w-md mx-auto z-40 md:hidden pointer-events-none"
    >
      <div
        className="relative pointer-events-auto flex items-center justify-around px-2 py-1.5 rounded-full backdrop-blur-2xl backdrop-saturate-200 bg-white/70 dark:bg-neutral-900/75 border border-white/60 dark:border-white/15 shadow-[0_12px_40px_-8px_rgba(0,0,0,0.18),inset_0_1px_1px_0_rgba(255,255,255,0.9),inset_0_-1px_1px_0_rgba(0,0,0,0.06)] dark:shadow-[0_12px_40px_-8px_rgba(0,0,0,0.7),inset_0_1px_1px_0_rgba(255,255,255,0.2),inset_0_-1px_1px_0_rgba(0,0,0,0.4)] ring-1 ring-black/5 dark:ring-white/10 transition-all duration-300"
        style={{
          paddingBottom: 'calc(0.375rem + env(safe-area-inset-bottom, 0px))',
        }}
      >
        {/* Animated Directional Sliding Liquid Glass Pill Indicator */}
        <div
          aria-hidden="true"
          className="absolute top-1.5 bottom-1.5 rounded-full bg-primary/10 dark:bg-primary/20 backdrop-blur-md border border-primary/20 dark:border-primary/30 shadow-[inset_0_1px_1px_0_rgba(255,255,255,0.5),0_2px_8px_0_rgba(0,0,0,0.06)] pointer-events-none z-0 transition-transform duration-300 ease-[cubic-bezier(0.32,0.72,0,1)]"
          style={{
            left: '0.5rem',
            transform: `translateX(${safeIndex * 100}%)`,
            width: `calc((100% - 1rem) / ${NAV_ITEMS.length})`,
          }}
        />

        {NAV_ITEMS.map((item, idx) => {
          const Icon = item.icon;
          const isActive = idx === activeIndex;

          return (
            <Link
              className="relative flex-1 flex flex-col items-center justify-center py-1.5 px-1 rounded-full transition-all duration-200 z-10 hover:bg-white/20 dark:hover:bg-white/10 active:scale-95"
              key={item.label}
              params={item.params}
              to={item.to}
            >
              <Icon
                className={`h-4 w-4 shrink-0 transition-all duration-300 drop-shadow-[0_1px_1px_rgba(255,255,255,0.5)] dark:drop-shadow-[0_1px_1px_rgba(0,0,0,0.6)] ${
                  isActive ? 'text-primary scale-110' : 'text-foreground/80 hover:text-foreground'
                }`}
                weight={isActive ? 'fill' : 'bold'}
              />
              <span
                className={`text-[10px] leading-tight mt-0.5 tracking-tight transition-colors duration-200 drop-shadow-[0_1px_1px_rgba(255,255,255,0.4)] dark:drop-shadow-[0_1px_1px_rgba(0,0,0,0.6)] ${
                  isActive ? 'text-primary font-bold' : 'text-foreground/85 font-semibold hover:text-foreground'
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
