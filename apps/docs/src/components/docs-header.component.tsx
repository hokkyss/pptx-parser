import { BookOpenIcon, CodeIcon, CubeIcon, GithubLogoIcon, PlayIcon, SparkleIcon } from '@phosphor-icons/react';
import { Link } from '@tanstack/react-router';
import ThemeSwitcher from './theme-switcher.component';

/**
 * Global navigation header with navigation links, theme toggle, and GitHub repo link.
 * @returns React node
 */
export default function DocsHeader() {
  return (
    <header className="sticky top-0 z-40 w-full border-b border-border/60 bg-background/80 backdrop-blur-md">
      <div className="w-full flex h-16 items-center justify-between px-4 sm:px-6 lg:px-8 max-w-[1600px] mx-auto">
        {/* Left: Logo & Brand */}
        <div className="flex items-center gap-6">
          <Link className="flex items-center gap-2.5 group" to="/">
            <div className="h-8 w-8 rounded-lg bg-linear-to-br from-sky-500 to-indigo-600 flex items-center justify-center text-white font-bold text-base shadow-md group-hover:scale-105 transition">
              P
            </div>
            <div className="flex flex-col">
              <span className="font-bold text-base tracking-tight text-foreground flex items-center gap-1.5">
                @hokkyss/pptx
                <span className="text-[10px] font-mono px-1.5 py-0.2 rounded bg-sky-500/10 text-sky-400 border border-sky-500/20">v1.0</span>
              </span>
            </div>
          </Link>

          {/* Desktop Nav Links */}
          <nav className="hidden md:flex items-center gap-1 pl-4 text-sm font-medium text-muted-foreground">
            <Link
              activeProps={{ className: 'text-foreground font-semibold bg-muted/60' }}
              className="px-3 py-1.5 rounded-md hover:text-foreground hover:bg-muted/40 transition flex items-center gap-1.5"
              params={{ _splat: 'getting-started/overview' }}
              to="/docs/$"
            >
              <BookOpenIcon className="h-4 w-4 text-sky-400" />
              Docs
            </Link>
            <Link
              activeProps={{ className: 'text-foreground font-semibold bg-muted/60' }}
              className="px-3 py-1.5 rounded-md hover:text-foreground hover:bg-muted/40 transition flex items-center gap-1.5"
              to="/cookbook"
            >
              <SparkleIcon className="h-4 w-4 text-amber-400" />
              Cookbook
            </Link>
            <Link
              activeProps={{ className: 'text-foreground font-semibold bg-muted/60' }}
              className="px-3 py-1.5 rounded-md hover:text-foreground hover:bg-muted/40 transition flex items-center gap-1.5"
              params={{ _splat: 'pptx/presentation' }}
              to="/api/$"
            >
              <CodeIcon className="h-4 w-4 text-purple-400" />
              API Reference
            </Link>
            <Link
              activeProps={{ className: 'text-foreground font-semibold bg-muted/60' }}
              className="px-3 py-1.5 rounded-md hover:text-foreground hover:bg-muted/40 transition flex items-center gap-1.5"
              to="/playground"
            >
              <PlayIcon className="h-4 w-4 text-emerald-400" />
              Playground
            </Link>
            <Link
              activeProps={{ className: 'text-foreground font-semibold bg-muted/60' }}
              className="px-3 py-1.5 rounded-md hover:text-foreground hover:bg-muted/40 transition flex items-center gap-1.5"
              to="/showcase"
            >
              <CubeIcon className="h-4 w-4 text-rose-400" />
              Showcase
            </Link>
          </nav>
        </div>

        {/* Right: GitHub & Theme Toggle */}
        <div className="flex items-center gap-3">
          <ThemeSwitcher />
          <a
            className="flex items-center gap-1.5 px-3 py-1.5 rounded-md text-xs font-semibold bg-muted hover:bg-muted/80 text-foreground transition"
            href="https://github.com/hokkyss/pptx-parser"
            rel="noreferrer"
            target="_blank"
          >
            <GithubLogoIcon className="h-4 w-4" />
            <span className="hidden sm:inline">GitHub</span>
          </a>
        </div>
      </div>
    </header>
  );
}
