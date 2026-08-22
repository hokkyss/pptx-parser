'use client';

import type { ComponentPropsWithoutRef } from 'react';
import Alert from '@monorepo/design-system/alert';
import AlertDescription from '@monorepo/design-system/alert-description';
import AlertTitle from '@monorepo/design-system/alert-title';
import cn from '@monorepo/design-system/cn';
import { InfoIcon, LightbulbIcon, ProhibitIcon, WarningCircleIcon, WarningIcon } from '@phosphor-icons/react';

export type CalloutType = 'caution' | 'important' | 'note' | 'tip' | 'warning';

export interface MarkdownCalloutProps extends ComponentPropsWithoutRef<'div'> {
  'data-callout-type'?: CalloutType;
}

const ICON_MAP: Record<CalloutType, typeof InfoIcon> = {
  caution: ProhibitIcon,
  important: WarningCircleIcon,
  note: InfoIcon,
  tip: LightbulbIcon,
  warning: WarningIcon,
};

const STYLE_MAP: Record<CalloutType, string> = {
  caution: 'bg-danger/10 text-danger border-danger/20 [&>svg]:text-danger',
  important: 'bg-important/10 text-important border-important/20 [&>svg]:text-important',
  note: 'bg-info/10 text-info border-info/20 [&>svg]:text-info',
  tip: 'bg-success/10 text-success border-success/20 [&>svg]:text-success',
  warning: 'bg-warning/10 text-warning border-warning/20 [&>svg]:text-warning',
};

/**
 * Callout alert component for markdown callouts matching Portfolio design system.
 * @param props Component props
 * @returns React node
 */
export default function MarkdownCallout(props: MarkdownCalloutProps) {
  if (props.className === 'markdown-callout') {
    const type = props['data-callout-type'] ?? 'note';
    const Icon = ICON_MAP[type] || InfoIcon;

    return (
      <Alert className={cn(STYLE_MAP[type], 'my-6')}>
        <Icon className="size-4" />
        {props.children}
      </Alert>
    );
  }

  if (props.className === 'markdown-callout-title') {
    return <AlertTitle className="font-semibold">{props.children}</AlertTitle>;
  }

  if (props.className === 'markdown-callout-description') {
    return <AlertDescription className="text-foreground/90">{props.children}</AlertDescription>;
  }

  return <div {...props} />;
}
