import { MoonIcon, SunIcon } from '@phosphor-icons/react';
import { useMutation, useSuspenseQuery } from '@tanstack/react-query';
import setApplicationThemeMutation from '../lib/common/mutations/set-application-theme.mutation';
import getApplicationThemeQuery from '../lib/common/queries/get-application-theme.query';

/**
 * Theme switcher toggle component between light and dark modes.
 * @returns React node
 */
export default function ThemeSwitcher() {
  const { data: theme } = useSuspenseQuery(getApplicationThemeQuery());
  const { mutate: setTheme } = useMutation({
    ...setApplicationThemeMutation(),
    onSuccess: (_data, variables, _onMutateResult, context) => {
      void context.client.setQueryData(getApplicationThemeQuery().queryKey, () => variables);
    },
  });

  const toggleTheme = () => {
    setTheme(theme === 'dark' ? 'light' : 'dark');
  };

  return (
    <button
      className="p-2 rounded-md hover:bg-muted text-muted-foreground hover:text-foreground transition"
      onClick={toggleTheme}
      title="Toggle theme"
      type="button"
    >
      {theme === 'dark'
        ? (
            <SunIcon className="h-4 w-4 text-foreground" />
          )
        : (
            <MoonIcon className="h-4 w-4 text-foreground" />
          )}
    </button>
  );
}
