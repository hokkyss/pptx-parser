import { useSuspenseQuery } from '@tanstack/react-query';
import { createFileRoute, Outlet } from '@tanstack/react-router';
import DocsSidebar from '../components/docs-sidebar.component';
import MobileSidebarSheet from '../components/mobile-sidebar-sheet.component';
import listDocsQuery from '../lib/content/queries/list-docs.query';

export const Route = createFileRoute('/api-reference')({
  loader: ({ context }) => context.queryClient.ensureQueryData(listDocsQuery('api-reference/')),
  component: ApiLayout,
});

/**
 * API reference shell layout component with sidebar and outlet.
 * @returns React node
 */
function ApiLayout() {
  const { data } = useSuspenseQuery(listDocsQuery('api-reference/'));

  return (
    <div className="w-full max-w-[1600px] mx-auto px-4 sm:px-6 lg:px-8 flex flex-col lg:flex-row gap-4 lg:gap-8">
      {/* Mobile Chapter Trigger */}
      <div className="lg:hidden pt-4 pb-1 flex items-center justify-between border-b border-border/40">
        <div className="flex items-center gap-2">
          <MobileSidebarSheet baseRoute="/api-reference" sections={data.sections} title="API Reference" />
          <span className="text-xs font-medium text-muted-foreground">Select API Reference Page</span>
        </div>
      </div>

      <DocsSidebar baseRoute="/api-reference" sections={data.sections} />
      <div className="flex-1 min-w-0 py-4 lg:py-8">
        <Outlet />
      </div>
    </div>
  );
}
