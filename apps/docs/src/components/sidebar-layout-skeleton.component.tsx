import type { ReactNode } from 'react';
import Skeleton from '@monorepo/design-system/skeleton';

interface SidebarLayoutSkeletonProps {
  children?: ReactNode;
}

/**
 * Skeleton fallback placeholder for documentation and API layout with sidebar.
 * @param props Component props
 * @param props.children Optional children to render in main outlet area
 * @returns React node
 */
export default function SidebarLayoutSkeleton({ children }: SidebarLayoutSkeletonProps) {
  return (
    <div className="w-full max-w-[1600px] mx-auto px-4 sm:px-6 lg:px-8 flex flex-col lg:flex-row gap-4 lg:gap-8">
      {/* Sidebar Placeholder */}
      <aside className="hidden lg:block w-64 shrink-0 py-8 space-y-6">
        <div className="space-y-3">
          <Skeleton className="h-4 w-28 rounded" />
          <div className="space-y-2 pl-2">
            <Skeleton className="h-7 w-48 rounded" />
            <Skeleton className="h-7 w-44 rounded" />
            <Skeleton className="h-7 w-40 rounded" />
          </div>
        </div>
        <div className="space-y-3">
          <Skeleton className="h-4 w-32 rounded" />
          <div className="space-y-2 pl-2">
            <Skeleton className="h-7 w-48 rounded" />
            <Skeleton className="h-7 w-40 rounded" />
          </div>
        </div>
      </aside>

      {/* Main Content Area */}
      <div className="flex-1 min-w-0 py-4 lg:py-8">
        {children}
      </div>
    </div>
  );
}
