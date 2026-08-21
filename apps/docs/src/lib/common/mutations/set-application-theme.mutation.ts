import type { ApplicationTheme } from '@monorepo/design-system/application-theme-provider';
import { mutationOptions } from '@tanstack/react-query';
import setApplicationThemeFunction from '../functions/set-application-theme.function';

/**
 * Returns mutation options for updating the user application theme.
 * @returns Mutation options
 */
export default function setApplicationThemeMutation() {
  return mutationOptions({
    mutationFn: (value: ApplicationTheme) => setApplicationThemeFunction({
      data: {
        theme: value,
      },
    }),
    mutationKey: ['MUTATION', 'THEME'] as const,
  });
}
