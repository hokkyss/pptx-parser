import { ArrowLeftIcon } from '@phosphor-icons/react';
import { useSuspenseQuery } from '@tanstack/react-query';
import { createFileRoute, Link } from '@tanstack/react-router';
import DocsToc from '../../components/docs-toc.component';
import getDocQuery, { type GetDocRscResponse } from '../../lib/content/queries/get-doc.query';

export const Route = createFileRoute('/cookbook/$slug')({
  loader: ({ context, params }) => {
    return context.queryClient.ensureQueryData(getDocQuery(`cookbook/${params.slug}`));
  },
  component: RecipePage,
});

/**
 * Cookbook recipe page component.
 * @returns React node
 */
function RecipePage() {
  const { slug } = Route.useParams();
  const { data: doc }: { data: GetDocRscResponse } = useSuspenseQuery(getDocQuery(`cookbook/${slug}`));

  return (
    <div className="w-full max-w-[1600px] mx-auto px-4 sm:px-6 lg:px-8 py-10 flex justify-between gap-8 lg:gap-12">
      <div className="flex-1 min-w-0 max-w-4xl xl:max-w-5xl">
        <Link className="inline-flex items-center gap-1.5 text-xs font-semibold text-muted-foreground hover:text-primary mb-6 transition" to="/cookbook">
          <ArrowLeftIcon className="h-3.5 w-3.5" />
          Back to all recipes
        </Link>

        <div className="mb-6 pb-4 border-b border-border/40">
          <h1 className="text-3xl font-bold text-foreground">{doc.title}</h1>
          {doc.description && <p className="mt-2 text-sm text-muted-foreground">{doc.description}</p>}
        </div>

        {doc.Renderable}
      </div>

      <DocsToc toc={doc.toc} />
    </div>
  );
}
