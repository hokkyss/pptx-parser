import type { ReactNode } from 'react';
import { CheckCircleIcon, InfoIcon, ShieldWarningIcon, WarningIcon } from '@phosphor-icons/react';

interface MarkdownCalloutProps {
  children: ReactNode;
  title?: string;
  type?: 'caution' | 'important' | 'note' | 'tip' | 'warning';
}

/**
 * Callout alert box component matching GitHub note/tip/warning/caution alerts.
 * @param root0 Component props
 * @param root0.children Content inside the callout
 * @param root0.title Optional custom title
 * @param root0.type Alert severity level
 * @returns React node
 */
export default function MarkdownCallout({
  children,
  title,
  type = 'note',
}: MarkdownCalloutProps) {
  const config = {
    caution: {
      bg: 'bg-rose-500/10 border-rose-500/30 text-rose-300',
      defaultTitle: 'Caution',
      icon: <ShieldWarningIcon className="h-5 w-5 text-rose-400 shrink-0" />,
    },
    important: {
      bg: 'bg-purple-500/10 border-purple-500/30 text-purple-300',
      defaultTitle: 'Important',
      icon: <InfoIcon className="h-5 w-5 text-purple-400 shrink-0" />,
    },
    note: {
      bg: 'bg-blue-500/10 border-blue-500/30 text-blue-300',
      defaultTitle: 'Note',
      icon: <InfoIcon className="h-5 w-5 text-blue-400 shrink-0" />,
    },
    tip: {
      bg: 'bg-emerald-500/10 border-emerald-500/30 text-emerald-300',
      defaultTitle: 'Tip',
      icon: <CheckCircleIcon className="h-5 w-5 text-emerald-400 shrink-0" />,
    },
    warning: {
      bg: 'bg-amber-500/10 border-amber-500/30 text-amber-300',
      defaultTitle: 'Warning',
      icon: <WarningIcon className="h-5 w-5 text-amber-400 shrink-0" />,
    },
  }[type];

  return (
    <div className={`my-6 rounded-lg border p-4 ${config.bg}`}>
      <div className="flex items-center gap-2 font-semibold">
        {config.icon}
        <span>{title || config.defaultTitle}</span>
      </div>
      <div className="mt-2 text-sm leading-relaxed text-foreground/90">
        {children}
      </div>
    </div>
  );
}
