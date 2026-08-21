import { useSuspenseQuery } from '@tanstack/react-query';
import { createFileRoute } from '@tanstack/react-router';
import type { GetDocResponseDto } from '../../lib/content/dto/get-doc.dto';
import DocsToc from '../../components/docs-toc.component';
import MarkdownRenderer from '../../components/markdown-renderer.component';
import getDocQuery from '../../lib/content/queries/get-doc.query';

export const Route = createFileRoute('/api-reference/$')({
  loader: ({ context, params }) => {
    return context.queryClient.ensureQueryData(getDocQuery(`api-reference/${params._splat}`));
  },
  component: ApiPage,
});

/**
 * API reference documentation page component.
 * @returns React node
 */
function ApiPage() {
  const { _splat } = Route.useParams();
  const { data: doc }: { data: GetDocResponseDto } = useSuspenseQuery(getDocQuery(`api-reference/${_splat}`));
  const pkgName = doc.package ?? 'API';

  return (
    <div className="flex justify-between gap-8 lg:gap-12">
      <div className="flex-1 min-w-0 max-w-4xl xl:max-w-5xl">
        <div className="mb-6 pb-4 border-b border-border/40">
          <span className="text-xs font-mono text-purple-600 dark:text-purple-400 bg-purple-500/10 px-2 py-0.5 rounded border border-purple-500/20">
            {pkgName}
          </span>
          <h1 className="text-3xl font-bold tracking-tight text-foreground mt-2">
            {doc.title}
          </h1>
          {doc.description && <p className="mt-2 text-sm text-muted-foreground">{doc.description}</p>}
        </div>

        <MarkdownRenderer content={doc.content} />
      </div>

      <DocsToc toc={doc.toc} />
    </div>
  );
}
