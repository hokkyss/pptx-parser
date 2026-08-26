import * as React from 'react';

const MOBILE_BREAKPOINT = 768;

/**
 * Hook to detect if current viewport width is below mobile breakpoint.
 * @returns Boolean indicating if viewport is mobile width
 */
export function useIsMobile(): boolean {
  return React.useSyncExternalStore(
    (onStoreChange) => {
      if (typeof window === 'undefined') return () => {};
      const mql = window.matchMedia(`(max-width: ${MOBILE_BREAKPOINT - 1}px)`);
      mql.addEventListener('change', onStoreChange);
      return () => mql.removeEventListener('change', onStoreChange);
    },
    () => (typeof window !== 'undefined' ? window.innerWidth < MOBILE_BREAKPOINT : false),
    () => false,
  );
}
