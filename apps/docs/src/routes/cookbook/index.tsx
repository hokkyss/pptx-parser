import { ArrowRightIcon, SparkleIcon } from '@phosphor-icons/react';
import { useSuspenseQuery } from '@tanstack/react-query';
import { createFileRoute, Link } from '@tanstack/react-router';
import listDocsQuery from '../../lib/content/queries/list-docs.query';

export const Route = createFileRoute('/cookbook/')({
  loader: ({ context }) => context.queryClient.ensureQueryData(listDocsQuery('cookbook/')),
  component: CookbookPage,
});

/**
 * Cookbook recipes grid page component.
 * @returns React node
 */
function CookbookPage() {
  const { data } = useSuspenseQuery(listDocsQuery('cookbook/'));
  const recipes = data.sections.flatMap((s) => s.items);

  return (
    <div className="w-full max-w-[1600px] mx-auto px-4 sm:px-6 lg:px-8 py-12">
      <div className="text-center max-w-2xl mx-auto mb-12">
        <div className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full bg-amber-500/10 text-amber-400 text-xs font-semibold mb-3 border border-amber-500/20">
          <SparkleIcon className="h-3.5 w-3.5" />
          Cookbook & Real-World Recipes
        </div>
        <h1 className="text-3xl sm:text-4xl font-extrabold tracking-tight text-foreground">
          Production PowerPoint Patterns
        </h1>
        <p className="mt-3 text-sm text-muted-foreground">
          Copy-pasteable, production-ready recipe implementations for common enterprise slide decks.
        </p>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
        {recipes.map((item) => (
          <Link
            className="p-6 rounded-2xl bg-card text-card-foreground border border-border hover:border-primary/50 hover:bg-card/80 transition flex flex-col justify-between group shadow-sm"
            key={item.path}
            params={{ slug: item.path.replace(/^cookbook\//, '') }}
            to="/cookbook/$slug"
          >
            <div>
              <span className="text-[10px] font-mono uppercase tracking-wider px-2 py-0.5 rounded bg-primary/10 text-primary border border-primary/20">
                Recipe
              </span>
              <h3 className="text-lg font-bold text-card-foreground mt-3 group-hover:text-primary transition">
                {item.title}
              </h3>
              <p className="text-xs text-muted-foreground mt-2 line-clamp-2">
                {item.description || 'Full walkthrough code and pattern architecture.'}
              </p>
            </div>

            <div className="mt-6 flex items-center gap-1.5 text-xs font-semibold text-primary group-hover:translate-x-1 transition">
              <span>View Recipe Walkthrough</span>
              <ArrowRightIcon className="h-3.5 w-3.5" />
            </div>
          </Link>
        ))}
      </div>
    </div>
  );
}
