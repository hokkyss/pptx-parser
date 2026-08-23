import Skeleton from '@monorepo/design-system/skeleton';

const SKELETON_SLOTS = ['item-1', 'item-2', 'item-3', 'item-4', 'item-5', 'item-6'];

/**
 * Skeleton fallback placeholder for the cookbook recipes grid.
 * @returns React node
 */
export default function CookbookSkeleton() {
  return (
    <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4 sm:gap-6 animate-pulse">
      {SKELETON_SLOTS.map((id) => (
        <div className="p-5 sm:p-6 rounded-2xl bg-card border border-border flex flex-col justify-between h-48 space-y-4" key={id}>
          <div className="space-y-3">
            <Skeleton className="h-5 w-16 rounded" />
            <Skeleton className="h-6 w-4/5 rounded-md" />
            <Skeleton className="h-4 w-full rounded" />
          </div>
          <Skeleton className="h-4 w-32 rounded" />
        </div>
      ))}
    </div>
  );
}
