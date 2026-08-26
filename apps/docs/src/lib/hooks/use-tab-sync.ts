'use client';

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

      window.addEventListener('tab-sync', onStoreChange);

      return () => window.removeEventListener('tab-sync', onStoreChange);
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
      document.cookie = `${cookieName}=${encodeURIComponent(val)}; path=/; max-age=31536000; SameSite=Lax`;

      window.dispatchEvent(
        new CustomEvent('tab-sync', {
          detail: { syncKey, value: val },
        }),
      );
    } else {
      setLocalTab(val);
    }
  };

  return { activeTab, setActiveTab };
}
