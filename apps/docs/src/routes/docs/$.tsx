import { useSuspenseQuery } from '@tanstack/react-query';
import { Await, createFileRoute } from '@tanstack/react-router';
import { Suspense } from 'react';
import DocPageSkeleton from '../../components/doc-page-skeleton.component';
import DocsToc from '../../components/docs-toc.component';
import getDocQuery from '../../lib/content/queries/get-doc.query';

export const Route = createFileRoute('/docs/$')({
  loader: ({ context, params }) => {
    return {
      docPromise: context.queryClient.ensureQueryData(getDocQuery(`docs/${params._splat}`)),
    };
  },
  component: DocPageWrapper,
});

/**
 * Documentation page content component rendered after Suspense resolution.
 * @returns React node
 */
function DocPageContent() {
  const { _splat } = Route.useParams();
  const { data: doc } = useSuspenseQuery(getDocQuery(`docs/${_splat}`));

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

/**
 * Documentation page route wrapper with unawaited loader and Suspense.
 * @returns React node
 */
function DocPageWrapper() {
  const { docPromise } = Route.useLoaderData();

  return (
    <Suspense fallback={<DocPageSkeleton />}>
      <Await promise={docPromise}>
        {() => <DocPageContent />}
      </Await>
    </Suspense>
  );
}
