import { format, parseISO } from 'date-fns';

/**
 * Format an ISO date string (e.g. "2023-01-01") to "MMM yyyy" (e.g. "Jan 2023").
 * @param isoDate ISO date string
 * @returns Formatted date string
 */
export function formatDate(isoDate: string): string {
  return format(parseISO(isoDate), 'MMM yyyy');
}
