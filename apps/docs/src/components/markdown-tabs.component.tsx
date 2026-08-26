'use client';

import type { ComponentProps } from 'react';
import cn from '@monorepo/design-system/cn';
import Tabs from '@monorepo/design-system/tabs';
import useTabSync from '../lib/hooks/use-tab-sync';

export interface MarkdownTabsProps extends Omit<ComponentProps<typeof Tabs>, 'onValueChange' | 'value'> {
  'data-default-value'?: string;
  'data-sync-key'?: string;
}

/**
 * Interactive Client Component rendering synchronized, pill-styled tabs for markdown containers.
 * @param props."data-sync-key"
 * @param props."data-default-value"
 * @param props.className
 * @param props.children
 * @param props Component props
 * @returns React node
 */
export default function MarkdownTabs({ children, className, 'data-default-value': defaultValue = '', 'data-sync-key': syncKey, ...props }: MarkdownTabsProps) {
  const { activeTab, setActiveTab } = useTabSync({
    defaultValue,
    syncKey,
  });

  return (
    <Tabs
      className={cn('w-full', className)}
      onValueChange={setActiveTab}
      value={activeTab}
      {...props}
    >
      {children}
    </Tabs>
  );
}
