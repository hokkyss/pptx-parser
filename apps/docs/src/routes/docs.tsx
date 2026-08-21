import { useSuspenseQuery } from '@tanstack/react-query';
import { createFileRoute, Outlet } from '@tanstack/react-router';
import DocsSidebar from '../components/docs-sidebar.component';
import listDocsQuery from '../lib/content/queries/list-docs.query';

export const Route = createFileRoute('/docs')({
  loader: ({ context }) => context.queryClient.ensureQueryData(listDocsQuery('docs/')),
  component: DocsLayout,
});

/**
 * Documentation shell layout component with sidebar and outlet.
 * @returns React node
 */
function DocsLayout() {
  const { data } = useSuspenseQuery(listDocsQuery('docs/'));

  return (
    <div className="w-full max-w-[1600px] mx-auto px-4 sm:px-6 lg:px-8 flex gap-8">
      <DocsSidebar baseRoute="/docs" sections={data.sections} />
      <div className="flex-1 min-w-0 py-8">
        <Outlet />
      </div>
    </div>
  );
}
