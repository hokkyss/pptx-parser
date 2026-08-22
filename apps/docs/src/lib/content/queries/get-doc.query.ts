import type { ReactNode } from 'react';
import { queryOptions } from '@tanstack/react-query';
import type { GetDocResponseDto } from '../dto/get-doc.dto';
import getDocFunction from '../functions/get-doc.function';

export interface GetDocRscResponse extends Omit<GetDocResponseDto, 'content'> {
  Renderable: ReactNode;
}

/**
 * Returns React Query options for fetching documentation markdown content by path as an RSC.
 * @param path Documentation content relative path
 * @returns Query options for getDocFunction with structuralSharing: false
 */
export default function getDocQuery(path: string) {
  return queryOptions<GetDocRscResponse>({
    queryFn: ({ signal }) => getDocFunction({ data: { path }, signal }),
    queryKey: ['doc', path],
    structuralSharing: false,
  });
}
