'use client';

import { CheckIcon, CopyIcon } from '@phosphor-icons/react';
import { useState } from 'react';

interface CodeCopyButtonProps {
  code: string;
}

/**
 * Interactive client-side copy button for syntax-highlighted code blocks in RSC.
 * @param root0 Component props
 * @param root0.code Plaintext code to copy to clipboard
 * @returns React node
 */
export default function CodeCopyButton({ code }: CodeCopyButtonProps) {
  const [copied, setCopied] = useState(false);

  const handleCopy = async () => {
    await navigator.clipboard.writeText(code);
    setCopied(true);
    setTimeout(() => {
      setCopied(false);
    }, 2000);
  };

  return (
    <button
      className="flex items-center gap-1.5 px-2 py-1 rounded bg-secondary hover:bg-secondary/80 text-secondary-foreground transition border border-border/50 active:scale-95 text-xs font-sans"
      onClick={() => {
        void handleCopy();
      }}
      title="Copy code"
      type="button"
    >
      {copied
        ? (
            <>
              <CheckIcon className="h-3.5 w-3.5 text-success" />
              <span className="text-success">Copied</span>
            </>
          )
        : (
            <>
              <CopyIcon className="h-3.5 w-3.5" />
              <span>Copy</span>
            </>
          )}
    </button>
  );
}
