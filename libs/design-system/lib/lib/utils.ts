import { type ClassValue, clsx } from 'clsx';
import { twMerge } from 'tailwind-merge';

/**
 * Merges class names safely using clsx and tailwind-merge.
 * @param inputs Class name values to merge
 * @returns Merged class names string
 */
export function cn(...inputs: ClassValue[]): string {
  return twMerge(clsx(inputs));
}
