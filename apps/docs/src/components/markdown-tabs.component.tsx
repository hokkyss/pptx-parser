'use client';

import type { ComponentProps } from 'react';
import cn from '@monorepo/design-system/cn';
import Tabs from '@monorepo/design-system/tabs';
import TabsContent from '@monorepo/design-system/tabs-content';
import TabsList from '@monorepo/design-system/tabs-list';
import TabsTrigger from '@monorepo/design-system/tabs-trigger';
import useTabSync from '../lib/hooks/use-tab-sync';

export interface MarkdownTabsContentProps extends ComponentProps<typeof TabsContent> {
}

export interface MarkdownTabsProps extends Omit<ComponentProps<typeof Tabs>, 'defaultValue' | 'onValueChange' | 'value'> {
  defaultValue: string;
  syncKey?: string;
}

export interface MarkdownTabsTriggerProps extends ComponentProps<typeof TabsTrigger> {
}

/**
 * Root interactive tab container that synchronizes active tab states across instances and cookies.
 * @param root0 Component props
 * @param root0.children Child nodes containing TabsList and TabsContent
 * @param root0.className Optional additional styling classes
 * @param root0."data-default-value" Fallback default tab value from markdown directive
 * @param root0."data-sync-key" Cookie and event synchronization key
 * @param root0.defaultValue Default value prop
 * @param root0.syncKey Synchronization key for tab selection across page and cookies
 * @returns React node
 */
export function MarkdownTabs({
  children,
  className,
  defaultValue = '',
  syncKey,
  ...props
}: MarkdownTabsProps) {
  const { activeTab, setActiveTab } = useTabSync({
    defaultValue,
    syncKey,
  });

  return (
    <Tabs
      className={cn('w-full my-6 not-prose', className)}
      onValueChange={setActiveTab}
      value={activeTab}
      {...props}
    >
      {children}
    </Tabs>
  );
}

/**
 * Tab panel content corresponding to a tab value.
 * @param root0 Component props
 * @param root0.className Optional additional styling classes
 * @param root0.children Panel content (e.g. code blocks)
 * @param root0."data-value" Value attribute from markdown AST
 * @param root0.value Explicit tab value prop
 * @returns React node
 */
export function MarkdownTabsContent({
  children,
  className,
  ...props
}: MarkdownTabsContentProps) {
  return (
    <TabsContent
      className={cn('flex-1 text-sm outline-none mt-2', className)}
      {...props}
    >
      {children}
    </TabsContent>
  );
}

/**
 * Container list for tab item triggers.
 * @param root0 Component props
 * @param root0.className Optional additional styling classes
 * @param root0.children Tab trigger children
 * @returns React node
 */
export function MarkdownTabsList({
  children,
  className,
  ...props
}: ComponentProps<typeof TabsList>) {
  return (
    <TabsList
      className={cn(
        'inline-flex items-center gap-1 p-1 bg-muted/60 dark:bg-muted/40 rounded-lg border border-border/50',
        className,
      )}
      {...props}
    >
      {children}
    </TabsList>
  );
}

/**
 * Individual tab button trigger.
 * @param root0 Component props
 * @param root0.className Optional additional styling classes
 * @param root0.children Label content
 * @param root0."data-value" Value attribute from markdown AST
 * @param root0.value Explicit tab value prop
 * @returns React node
 */
export function MarkdownTabsTrigger({
  children,
  className,
  ...props
}: MarkdownTabsTriggerProps) {
  return (
    <TabsTrigger
      className={cn(
        'px-3 py-1 rounded-md text-xs font-semibold font-mono text-muted-foreground hover:text-foreground data-active:bg-background data-active:text-foreground data-active:shadow-sm dark:data-active:bg-card transition',
        className,
      )}
      {...props}
    >
      {children}
    </TabsTrigger>
  );
}
