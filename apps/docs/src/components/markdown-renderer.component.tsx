import type { ComponentPropsWithoutRef, ReactNode } from 'react';
import { CheckIcon, CopyIcon } from '@phosphor-icons/react';
import { useState } from 'react';
import ReactMarkdown from 'react-markdown';
import rehypeHighlight from 'rehype-highlight';
import rehypeKatex from 'rehype-katex';
import remarkGemoji from 'remark-gemoji';
import remarkGfm from 'remark-gfm';
import remarkMath from 'remark-math';

interface CodeBlockProps extends ComponentPropsWithoutRef<'code'> {
  children?: ReactNode;
  className?: string;
}

interface MarkdownRendererProps {
  content: string;
}

/**
 * Markdown renderer supporting GFM tables, math (KaTeX), emojis, and syntax highlighting.
 * @param root0 Component props
 * @param root0.content Raw markdown string to render
 * @returns React node
 */
export default function MarkdownRenderer({ content }: MarkdownRendererProps) {
  return (
    <article className="prose dark:prose-invert max-w-none text-foreground leading-relaxed">
      <ReactMarkdown
        components={{
          a: ({ children, href, ...props }) => (
            <a
              className="text-primary hover:text-primary/80 underline underline-offset-4 decoration-primary/40 hover:decoration-primary transition"
              href={href}
              {...props}
            >
              {children}
            </a>
          ),
          blockquote: ({ children }) => {
            return (
              <blockquote className="border-l-4 border-primary/80 bg-muted/50 my-4 px-4 py-2 rounded-r italic text-muted-foreground">
                {children}
              </blockquote>
            );
          },
          code: ({ children, className, ...props }) => {
            const isInline = !className || !className.includes('language-');
            if (isInline) {
              return (
                <code className="px-1.5 py-0.5 rounded font-mono text-sm" {...props}>
                  {children}
                </code>
              );
            }
            return <CodeBlock className={className} {...props}>{children}</CodeBlock>;
          },
          h1: ({ children }) => (
            <h1 className="text-3xl font-bold tracking-tight text-foreground mt-8 mb-4 border-b border-border/40 pb-2">
              {children}
            </h1>
          ),
          h2: ({ children }) => {
            const text = extractNodeText(children);
            const id = text.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/(^-|-$)/g, '');
            return (
              <h2 className="text-2xl font-semibold tracking-tight text-foreground mt-8 mb-3 scroll-mt-20" id={id}>
                {children}
              </h2>
            );
          },
          h3: ({ children }) => {
            const text = extractNodeText(children);
            const id = text.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/(^-|-$)/g, '');
            return (
              <h3 className="text-xl font-medium tracking-tight text-foreground mt-6 mb-2 scroll-mt-20" id={id}>
                {children}
              </h3>
            );
          },
          table: ({ children }) => (
            <div className="my-6 overflow-x-auto rounded-lg border border-border">
              <table className="w-full text-left text-sm border-collapse">{children}</table>
            </div>
          ),
          td: ({ children }) => (
            <td className="px-4 py-2.5 border-b border-border/40 text-muted-foreground">
              {children}
            </td>
          ),
          th: ({ children }) => (
            <th className="bg-muted/60 px-4 py-3 font-semibold text-foreground border-b border-border">
              {children}
            </th>
          ),
        }}
        rehypePlugins={[
          [rehypeHighlight, { ignoreMissing: true }],
          rehypeKatex,
        ]}
        remarkPlugins={[remarkGfm, remarkGemoji, remarkMath]}
      >
        {content}
      </ReactMarkdown>
    </article>
  );
}

/**
 * Renders a syntax-highlighted code block with a one-click copy button.
 * @param root0 Component props
 * @param root0.children Code text content
 * @param root0.className CSS class containing language name
 * @returns React node
 */
function CodeBlock({ children, className, ...props }: CodeBlockProps) {
  const [copied, setCopied] = useState(false);
  const codeText = extractNodeText(children).replace(/\n$/, '');
  const match = /language-(\w+)/.exec(className || '');
  const language = match ? match[1] : '';

  const handleCopy = async () => {
    await navigator.clipboard.writeText(codeText);
    setCopied(true);
    setTimeout(() => {
      setCopied(false);
    }, 2000);
  };

  return (
    <div className="relative my-6 group rounded-xl overflow-hidden border border-border bg-card shadow-sm">
      <div className="flex items-center justify-between px-4 py-2.5 bg-muted/80 border-b border-border text-xs font-mono text-muted-foreground">
        <div className="flex items-center gap-2">
          <span className="h-2.5 w-2.5 rounded-full bg-rose-500/80" />
          <span className="h-2.5 w-2.5 rounded-full bg-amber-500/80" />
          <span className="h-2.5 w-2.5 rounded-full bg-emerald-500/80" />
          <span className="ml-1.5 font-semibold text-foreground">{language || 'text'}</span>
        </div>
        <button
          className="flex items-center gap-1.5 px-2 py-1 rounded bg-secondary hover:bg-secondary/80 text-secondary-foreground transition border border-border/50"
          onClick={() => {
            void handleCopy();
          }}
          title="Copy code"
          type="button"
        >
          {copied
            ? (
                <>
                  <CheckIcon className="h-3.5 w-3.5 text-emerald-500" />
                  <span className="text-emerald-500">Copied</span>
                </>
              )
            : (
                <>
                  <CopyIcon className="h-3.5 w-3.5" />
                  <span>Copy</span>
                </>
              )}
        </button>
      </div>
      <pre className="p-4 overflow-x-auto text-sm leading-relaxed font-mono bg-slate-950 text-slate-200">
        <code className={className} {...props}>
          {children}
        </code>
      </pre>
    </div>
  );
}

/**
 * Recursively extracts plain text from ReactNode for heading IDs and code snippets.
 * @param node ReactNode to extract text from
 * @returns Extracted string
 */
function extractNodeText(node: ReactNode): string {
  if (typeof node === 'string') return node;
  if (typeof node === 'number') return String(node);
  if (Array.isArray(node)) return node.map(extractNodeText).join('');
  if (node && typeof node === 'object' && 'props' in node) {
    const props = (node as { props?: { children?: ReactNode } }).props;
    return extractNodeText(props?.children);
  }
  return '';
}
