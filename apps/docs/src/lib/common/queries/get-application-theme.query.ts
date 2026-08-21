import type { ApplicationTheme } from '@monorepo/design-system/application-theme-provider';
import { queryOptions } from '@tanstack/react-query';
import getApplicationThemeFunction from '../functions/get-application-theme.function';

interface GetApplicationThemeQueryProps<Selected = ApplicationTheme> {
  selector?: (data: ApplicationTheme) => Selected;
}

/**
 * Returns query options for retrieving the active application theme.
 * @param root0 Query props
 * @param root0.selector Optional selector function
 * @returns Query options
 */
export default function getApplicationThemeQuery<Selected = ApplicationTheme>({
  selector,
}: GetApplicationThemeQueryProps<Selected> = {}) {
  return queryOptions({
    queryFn: async ({ signal }) =>
      getApplicationThemeFunction({
        signal,
      }),
    queryKey: ['THEME'] as const,
    select: selector,
    staleTime: Infinity,
  });
}
