import { queryOptions } from '@tanstack/react-query';
import renderTeaserFunction from '../functions/get-teaser.function';

/**
 * React Query options for fetching the teaser snippet as an RSC.
 * @returns React Query options object with structuralSharing: false
 */
export default function renderTeaserQuery() {
  return queryOptions({
    queryFn: ({ signal }) => renderTeaserFunction({ data: {}, signal }),
    queryKey: ['teaser'],
    structuralSharing: false,
  });
}
