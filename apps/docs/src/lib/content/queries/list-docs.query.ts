import { queryOptions } from '@tanstack/react-query';
import type { ListDocsResponseDto } from '../dto/list-docs.dto';
import listDocsFunction from '../functions/list-docs.function';

/**
 * Returns React Query options for listing documentation navigation sections.
 * @param section Optional section path prefix to filter
 * @returns Query options for listDocsFunction
 */
export default function listDocsQuery(section?: string) {
  return queryOptions<ListDocsResponseDto>({
    queryFn: ({ signal }) => listDocsFunction({ data: { section }, signal }),
    queryKey: ['docs-list', section ?? 'all'],
  });
}
