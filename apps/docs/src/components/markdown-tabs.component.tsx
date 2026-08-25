'use client';

import type { ComponentPropsWithoutRef, ReactNode } from 'react';
import Tabs from '@monorepo/design-system/tabs';
import TabsContent from '@monorepo/design-system/tabs-content';
import TabsList from '@monorepo/design-system/tabs-list';
import TabsTrigger from '@monorepo/design-system/tabs-trigger';
import { isValidElement } from 'react';
import type { TabItemMeta } from '../lib/plugins/remark-tabs.plugin';
import useTabSync from '../lib/hooks/use-tab-sync';

export interface MarkdownTabsProps extends ComponentPropsWithoutRef<'div'> {
  'data-default-value'?: string;
  'data-sync-key'?: string;
  'data-tabs'?: string;
}

/**
 * Interactive Client Component rendering synchronized, pill-styled tabs for markdown containers.
 * @param props Component props
 * @returns React node
 */
export default function MarkdownTabs(props: MarkdownTabsProps) {
  const syncKey = props['data-sync-key'];
  const rawTabs = props['data-tabs'];

  let tabs: TabItemMeta[] = [];
  try {
    tabs = rawTabs ? (JSON.parse(rawTabs) as TabItemMeta[]) : [];
  } catch {
    tabs = [];
  }

  const fallbackDefault = tabs[0]?.value ?? 'tab-0';
  const defaultValue = props['data-default-value'] || fallbackDefault;

  const { activeTab, setActiveTab } = useTabSync({
    defaultValue,
    syncKey,
  });

  const childrenArray = Array.isArray(props.children) ? props.children : [props.children];

  return (
    <div className="my-6">
      <Tabs
        className="w-full"
        onValueChange={setActiveTab}
        value={activeTab}
      >
        <div className="flex items-center justify-between mb-2">
          <TabsList className="inline-flex items-center gap-1 p-1 bg-muted/60 dark:bg-muted/40 rounded-lg border border-border/50">
            {tabs.map((tab) => (
              <TabsTrigger
                className="px-3 py-1 rounded-md text-xs font-semibold font-mono text-muted-foreground hover:text-foreground data-active:bg-background data-active:text-foreground data-active:shadow-sm dark:data-active:bg-card transition"
                key={tab.value}
                value={tab.value}
              >
                {tab.label}
              </TabsTrigger>
            ))}
          </TabsList>
        </div>

        <div className="relative">
          {childrenArray.map((child: ReactNode) => {
            if (!isValidElement(child)) return null;
            const tabVal = (child.props as Record<string, unknown> | undefined)?.['data-tab-value'] as string | undefined;

            if (tabVal) {
              return (
                <TabsContent
                  className="mt-0 focus-visible:outline-none data-[state=inactive]:hidden"
                  key={tabVal}
                  value={tabVal}
                >
                  {(child.props as { children?: ReactNode }).children}
                </TabsContent>
              );
            }

            return null;
          })}
        </div>
      </Tabs>
    </div>
  );
}
