import { queryOptions } from '@tanstack/react-query';
import type { GetDocResponseDto } from '../dto/get-doc.dto';
import getDocFunction from '../functions/get-doc.function';

/**
 * Returns React Query options for fetching documentation markdown content by path.
 * @param path Documentation content relative path
 * @returns Query options for getDocFunction
 */
export default function getDocQuery(path: string) {
  return queryOptions<GetDocResponseDto>({
    queryFn: ({ signal }) => getDocFunction({ data: { path }, signal }),
    queryKey: ['doc', path],
  });
}
