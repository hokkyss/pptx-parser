import { useSuspenseQuery } from '@tanstack/react-query';
import { Await, createFileRoute, Outlet } from '@tanstack/react-router';
import { Suspense } from 'react';
import DocsSidebar from '../components/docs-sidebar.component';
import MobileSidebarSheet from '../components/mobile-sidebar-sheet.component';
import SidebarLayoutSkeleton from '../components/sidebar-layout-skeleton.component';
import listDocsQuery from '../lib/content/queries/list-docs.query';

export const Route = createFileRoute('/docs')({
  component: DocsLayoutWrapper,
  loader: ({ context }) => {
    return {
      docsPromise: context.queryClient.ensureQueryData(listDocsQuery('docs/')),
    };
  },
});

/**
 * Documentation shell layout content rendered after Suspense resolution.
 * @returns React node
 */
function DocsLayoutContent() {
  const { data } = useSuspenseQuery(listDocsQuery('docs/'));

  return (
    <div className="w-full max-w-[1600px] mx-auto px-4 sm:px-6 lg:px-8 flex flex-col lg:flex-row gap-4 lg:gap-8">
      {/* Mobile Chapter Trigger */}
      <div className="lg:hidden pt-4 pb-1 flex items-center justify-between border-b border-border/40">
        <div className="flex items-center gap-2">
          <MobileSidebarSheet baseRoute="/docs" sections={data.sections} title="Documentation" />
          <span className="text-xs font-medium text-muted-foreground">Select Documentation Page</span>
        </div>
      </div>

      <DocsSidebar baseRoute="/docs" sections={data.sections} />
      <div className="flex-1 min-w-0 py-4 lg:py-8">
        <Outlet />
      </div>
    </div>
  );
}

/**
 * Documentation shell layout wrapper with unawaited loader and Suspense.
 * @returns React node
 */
function DocsLayoutWrapper() {
  const { docsPromise } = Route.useLoaderData();

  return (
    <Suspense
      fallback={(
        <SidebarLayoutSkeleton>
          <Outlet />
        </SidebarLayoutSkeleton>
      )}
    >
      <Await promise={docsPromise}>
        {() => <DocsLayoutContent />}
      </Await>
    </Suspense>
  );
}
