import { createIsomorphicFn } from '@tanstack/react-start';
import { getCookie as getServerCookie } from '@tanstack/react-start/server';
import { useCallback, useState, useSyncExternalStore } from 'react';

interface UseTabSyncOptions {
  defaultValue: string;
  syncKey?: string;
}

interface UseTabSyncReturn {
  activeTab: string;
  setActiveTab: (val: string) => void;
}

/**
 * Isomorphic cookie reader splitting server and client implementations:
 * - Server: Reads from request cookies via `@tanstack/react-start/server` getCookie
 * - Client: Reads from `document.cookie`
 */
const getCookie = createIsomorphicFn()
  .server((cookieName: string): null | string => {
    try {
      return getServerCookie(cookieName) ?? null;
    } catch {
      return null;
    }
  })
  .client((cookieName: string): null | string => {
    if (typeof document === 'undefined') return null;
    const match = document.cookie.match(new RegExp(`(^|;\\s*)(${cookieName})=([^;]*)`));
    return match ? decodeURIComponent(match[3]) : null;
  });

/**
 * Hook to synchronize tab selections across the application via cookies and window custom events.
 * If syncKey is provided, tab choice is synchronized across all tabs with the same syncKey and persisted in cookie.
 * @param options Hook options
 * @param options.syncKey Unique key to synchronize across matching tabs
 * @param options.defaultValue Fallback default tab value
 * @returns Object with activeTab and setActiveTab updater
 */
export default function useTabSync({ defaultValue, syncKey }: UseTabSyncOptions): UseTabSyncReturn {
  const cookieName = syncKey ? `tab_sync_${syncKey}` : undefined;

  const [localTab, setLocalTab] = useState<string>(defaultValue);

  const subscribe = useCallback(
    (onStoreChange: () => void) => {
      if (!syncKey) return () => {};

      const handleSync = () => {
        onStoreChange();
      };

      window.addEventListener('tab-sync', handleSync);
      return () => {
        window.removeEventListener('tab-sync', handleSync);
      };
    },
    [syncKey],
  );

  const getSnapshot = useCallback(() => {
    if (cookieName) {
      const stored = getCookie(cookieName);
      if (stored) return stored;
    }
    return defaultValue;
  }, [cookieName, defaultValue]);

  const getServerSnapshot = useCallback(() => {
    if (cookieName) {
      const serverStored = getCookie(cookieName);
      if (serverStored) return serverStored;
    }
    return defaultValue;
  }, [cookieName, defaultValue]);

  const syncedTab = useSyncExternalStore(subscribe, getSnapshot, getServerSnapshot);

  const activeTab = syncKey ? syncedTab : localTab;

  const setActiveTab = (val: string) => {
    if (syncKey && cookieName) {
      setCookie(cookieName, val);
      if (typeof window !== 'undefined') {
        window.dispatchEvent(
          new CustomEvent('tab-sync', {
            detail: { syncKey, value: val },
          }),
        );
      }
    } else {
      setLocalTab(val);
    }
  };

  return { activeTab, setActiveTab };
}

/**
 * Sets a persistent cookie for 1 year with SameSite=Lax.
 * @param name Cookie name
 * @param value Cookie value
 */
function setCookie(name: string, value: string): void {
  if (typeof document === 'undefined') return;
  document.cookie = `${name}=${encodeURIComponent(value)}; path=/; max-age=31536000; SameSite=Lax`;
}
