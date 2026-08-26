import Skeleton from '@monorepo/design-system/skeleton';

/**
 * Skeleton fallback placeholder for documentation and API reference pages.
 * @returns React node
 */
export default function DocPageSkeleton() {
  return (
    <div className="w-full max-w-4xl space-y-6 animate-pulse">
      {/* Header Title & Subtitle */}
      <div className="pb-4 border-b border-border/40 space-y-3">
        <Skeleton className="h-9 sm:h-11 w-3/4 max-w-md rounded-lg" />
        <Skeleton className="h-5 w-full max-w-xl rounded-md" />
      </div>

      {/* Paragraph blocks */}
      <div className="space-y-2.5 pt-2">
        <Skeleton className="h-4 w-full rounded" />
        <Skeleton className="h-4 w-11/12 rounded" />
        <Skeleton className="h-4 w-4/5 rounded" />
      </div>

      {/* Code snippet block skeleton */}
      <div className="pt-2">
        <Skeleton className="h-56 w-full rounded-xl" />
      </div>

      {/* Subsequent paragraph block */}
      <div className="space-y-2.5 pt-2">
        <Skeleton className="h-4 w-full rounded" />
        <Skeleton className="h-4 w-3/4 rounded" />
      </div>
    </div>
  );
}
