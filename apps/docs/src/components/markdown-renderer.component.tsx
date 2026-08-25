import type { ComponentPropsWithoutRef, ReactNode } from 'react';
import ReactMarkdown from 'react-markdown';
import rehypeHighlight from 'rehype-highlight';
import rehypeKatex from 'rehype-katex';
import remarkDirective from 'remark-directive';
import remarkGemoji from 'remark-gemoji';
import remarkGfm from 'remark-gfm';
import remarkMath from 'remark-math';
import remarkCallout from '../lib/plugins/remark-callout.plugin';
import remarkTabs from '../lib/plugins/remark-tabs.plugin';
import CodeCopyButton from './code-copy-button.component';
import MarkdownCallout from './markdown-callout.component';
import MarkdownTabs from './markdown-tabs.component';

interface CodeBlockProps extends ComponentPropsWithoutRef<'code'> {
  children?: ReactNode;
  className?: string;
}

interface MarkdownRendererProps {
  content: string;
}

/**
 * React Server Component (RSC) markdown renderer supporting GFM tables, math (KaTeX), emojis, syntax highlighting, callouts, and container tabs.
 * @param root0 Component props
 * @param root0.content Raw markdown string to render
 * @returns React node
 */
export default function MarkdownRenderer({ content }: MarkdownRendererProps) {
  return (
    <article className="prose prose-brand dark:prose-invert max-w-none text-foreground leading-relaxed">
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
          div: (props) => {
            if (
              props.className === 'markdown-callout'
              || props.className === 'markdown-callout-title'
              || props.className === 'markdown-callout-description'
            ) {
              return <MarkdownCallout {...props} />;
            }
            if (props.className === 'markdown-tabs') {
              return <MarkdownTabs {...props} />;
            }
            return <div {...props} />;
          },
          h1: ({ children }) => (
            <h1 className="text-2xl sm:text-3xl font-bold tracking-tight text-foreground mt-8 mb-4 border-b border-border/40 pb-2">
              {children}
            </h1>
          ),
          h2: ({ children }) => {
            const text = extractNodeText(children);
            const id = text.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/(^-|-$)/g, '');
            return (
              <h2 className="text-xl sm:text-2xl font-semibold tracking-tight text-foreground mt-8 mb-3 scroll-mt-20" id={id}>
                {children}
              </h2>
            );
          },
          h3: ({ children }) => {
            const text = extractNodeText(children);
            const id = text.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/(^-|-$)/g, '');
            return (
              <h3 className="text-lg sm:text-xl font-medium tracking-tight text-foreground mt-6 mb-2 scroll-mt-20" id={id}>
                {children}
              </h3>
            );
          },
          table: ({ children }) => (
            <div className="my-6 overflow-x-auto rounded-lg border border-border max-w-full">
              <table className="w-full text-left text-sm border-collapse min-w-125 sm:min-w-full">{children}</table>
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
        remarkPlugins={[remarkGfm, remarkGemoji, remarkMath, remarkDirective, remarkTabs, remarkCallout]}
      >
        {content}
      </ReactMarkdown>
    </article>
  );
}

/**
 * Server Component rendering syntax-highlighted code blocks with an embedded client copy button.
 * @param root0 Component props
 * @param root0.children Code text content
 * @param root0.className CSS class containing language name
 * @returns React node
 */
function CodeBlock({ children, className, ...props }: CodeBlockProps) {
  const codeText = extractNodeText(children).replace(/\n$/, '');
  const match = /language-(\w+)/.exec(className || '');
  const rawLang = match ? match[1] : '';
  const isTerminal = rawLang === 'bash' || rawLang === 'sh' || rawLang === 'shell' || rawLang === 'zsh';
  const label = isTerminal ? '>_ Terminal' : rawLang || 'text';

  return (
    <div className="relative my-4 group rounded-xl overflow-hidden border border-border bg-card shadow-sm max-w-full">
      <div className="flex items-center justify-between px-3 sm:px-4 py-2.5 bg-muted/80 border-b border-border text-xs font-mono text-muted-foreground">
        <div className="flex items-center gap-2">
          <span className="h-2.5 w-2.5 rounded-full bg-destructive/80" />
          <span className="h-2.5 w-2.5 rounded-full bg-warning/80" />
          <span className="h-2.5 w-2.5 rounded-full bg-success/80" />
          <span className="ml-1.5 font-semibold text-foreground">{label}</span>
        </div>
        <CodeCopyButton code={codeText} />
      </div>
      <pre className="p-3 sm:p-4 overflow-x-auto text-xs sm:text-sm leading-relaxed font-mono bg-card text-card-foreground">
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
