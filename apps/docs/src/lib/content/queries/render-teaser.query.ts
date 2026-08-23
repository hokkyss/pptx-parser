import { queryOptions } from '@tanstack/react-query';
import renderTeaserFunction from '../functions/get-teaser.function';

/**
 */
export default function renderTeaserQuery() {
  return queryOptions({
    queryFn: ({ signal }) => renderTeaserFunction({ data: { }, signal }),
    queryKey: ['teaser'],
    structuralSharing: false,
  });
}
