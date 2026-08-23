import { useSuspenseQuery } from '@tanstack/react-query';
import { createFileRoute, useParams } from '@tanstack/react-router';
import DocsToc from '../../components/docs-toc.component';
import getDocQuery from '../../lib/content/queries/get-doc.query';

export const Route = createFileRoute('/docs/$')({
  loader: ({ context, params }) => {
    return context.queryClient.ensureQueryData(getDocQuery(`docs/${params._splat}`));
  },
  component: DocPage,
});

/**
 * Documentation page renderer component.
 * @returns React node
 */
function DocPage() {
  const splat = useParams({
    from: '/docs/$',
    select: (p) => p._splat,
  });
  const { data: doc } = useSuspenseQuery(getDocQuery(`docs/${splat}`));

  return (
    <div className="flex justify-between gap-8 lg:gap-12">
      <div className="flex-1 min-w-0 max-w-4xl xl:max-w-5xl">
        <div className="mb-6 pb-4 border-b border-border/40">
          <h1 className="text-3xl sm:text-4xl font-bold tracking-tight text-foreground">
            {doc.title}
          </h1>
          {doc.description && (
            <p className="mt-2 text-base text-muted-foreground">
              {doc.description}
            </p>
          )}
        </div>

        {doc.Renderable}
      </div>

      <DocsToc toc={doc.toc} />
    </div>
  );
}
