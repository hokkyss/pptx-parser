import { useSuspenseQuery } from '@tanstack/react-query';
import { Await, createFileRoute, useParams } from '@tanstack/react-router';
import { Suspense } from 'react';
import DocPageSkeleton from '../../components/doc-page-skeleton.component';
import DocsToc from '../../components/docs-toc.component';
import getDocQuery from '../../lib/content/queries/get-doc.query';

export const Route = createFileRoute('/api-reference/$')({
  component: ApiPageWrapper,
  loader: ({ context, params }) => {
    return {
      docPromise: context.queryClient.ensureQueryData(getDocQuery(`api-reference/${params._splat}`)),
    };
  },
});

/**
 * API reference page content component rendered after Suspense resolution.
 * @returns React node
 */
function ApiPageContent() {
  const splat = useParams({
    from: '/api-reference/$',
    select: (p) => p._splat,
  });
  const { data: doc } = useSuspenseQuery(getDocQuery(`api-reference/${splat}`));
  const pkgName = doc.package ?? 'API';

  return (
    <div className="flex justify-between gap-8 lg:gap-12">
      <div className="flex-1 min-w-0 max-w-4xl xl:max-w-5xl">
        <div className="mb-6 pb-4 border-b border-border/40">
          <span className="text-xs font-mono text-primary bg-primary/10 px-2 py-0.5 rounded border border-primary/20">
            {pkgName}
          </span>
          <h1 className="text-3xl font-bold tracking-tight text-foreground mt-2">
            {doc.title}
          </h1>
          {doc.description && <p className="mt-2 text-sm text-muted-foreground">{doc.description}</p>}
        </div>

        {doc.Renderable}
      </div>

      <DocsToc toc={doc.toc} />
    </div>
  );
}

/**
 * API reference page route wrapper with unawaited loader and Suspense.
 * @returns React node
 */
function ApiPageWrapper() {
  const { docPromise } = Route.useLoaderData();

  return (
    <Suspense fallback={<DocPageSkeleton />}>
      <Await promise={docPromise}>
        {() => <ApiPageContent />}
      </Await>
    </Suspense>
  );
}
