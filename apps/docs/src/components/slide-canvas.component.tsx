import { inches, points, Presentation } from '@hokkyss/pptx';
import { useState } from 'react';

type ExecutableFn = (
  p: typeof Presentation,
  i: typeof inches,
  pt: typeof points,
) => Promise<ExecutablePresentation>;

interface ExecutablePresentation {
  toArrayBuffer: () => Promise<ArrayBuffer>;
  toAst?: () => { metadata?: { title?: string } };
}

interface SlideCanvasProps {
  code: string;
}

/**
 * Slide canvas preview that compiles code and triggers PPTX downloads.
 * @param root0 Component props
 * @param root0.code TypeScript source code to execute
 * @returns React node
 */
export default function SlideCanvas({ code }: SlideCanvasProps) {
  const [error, setError] = useState<null | string>(null);
  const [downloading, setDownloading] = useState(false);

  const executeCodeAndDownload = async () => {
    try {
      setDownloading(true);
      setError(null);

      // eslint-disable-next-line @typescript-eslint/no-unsafe-member-access
      const AsyncFunction = Object.getPrototypeOf(async () => {}).constructor as new (
        ...args: string[]
      ) => ExecutableFn;
      const fn = new AsyncFunction('Presentation', 'inches', 'points', `${code}\nreturn pres;`);

      const pres = await fn(Presentation, inches, points);

      if (!pres || typeof pres.toArrayBuffer !== 'function') {
        throw new Error('Script must define and initialize a valid "pres" (Presentation.create()) instance.');
      }

      const buffer = await pres.toArrayBuffer();
      const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      const ast = typeof pres.toAst === 'function' ? pres.toAst() : undefined;
      const title = ast?.metadata?.title || 'presentation';
      a.download = `${title.replace(/[^a-z0-9_-]/gi, '_')}.pptx`;
      a.click();
      URL.revokeObjectURL(url);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to generate PowerPoint deck.');
    } finally {
      setDownloading(false);
    }
  };

  return (
    <div className="flex flex-col h-full bg-card rounded-xl border border-border p-5 shadow-sm">
      <div className="flex items-center justify-between pb-4 border-b border-border">
        <div className="flex items-center gap-2">
          <span className="h-3 w-3 rounded-full bg-emerald-500 animate-pulse" />
          <span className="text-xs font-semibold uppercase tracking-wider text-muted-foreground">
            Interactive Slide Canvas Preview
          </span>
        </div>
        <button
          className="flex items-center gap-2 px-4 py-2 text-xs font-semibold rounded-lg bg-primary hover:bg-primary/90 text-primary-foreground shadow-sm transition active:scale-95 disabled:opacity-50"
          disabled={downloading}
          onClick={() => {
            void executeCodeAndDownload();
          }}
          type="button"
        >
          {downloading ? 'Compiling PPTX...' : '⚡ Download .pptx Deck'}
        </button>
      </div>

      {error
        ? (
            <div className="my-auto p-4 rounded-lg bg-rose-500/10 border border-rose-500/30 text-rose-300 text-sm font-mono">
              <p className="font-bold mb-1">Execution Error:</p>
              <p>{error}</p>
            </div>
          )
        : (
            <div className="my-auto flex flex-col items-center justify-center p-6">
              {/* 16:9 Aspect Ratio Slide Canvas Box */}
              <div className="w-full max-w-xl aspect-[16/9] rounded-lg bg-slate-950 border-2 border-dashed border-sky-500/40 shadow-2xl p-6 flex flex-col justify-between relative overflow-hidden group">
                <div className="absolute inset-0 bg-gradient-to-br from-sky-500/5 to-indigo-500/5 pointer-events-none" />

                <div className="flex items-center justify-between">
                  <span className="text-xs font-mono text-sky-400 bg-sky-500/10 px-2 py-0.5 rounded border border-sky-500/20">
                    16:9 Widescreen (13.33" x 7.5")
                  </span>
                  <span className="text-xs text-muted-foreground font-mono">Slide #1</span>
                </div>

                <div className="text-center my-auto space-y-2">
                  <h4 className="text-xl font-bold text-white tracking-tight">Live In-Browser Presentation Engine</h4>
                  <p className="text-xs text-slate-400 max-w-md mx-auto">
                    Edits in the Monaco code editor compile in real time into binary OpenXML archives via @hokkyss/pptx.
                  </p>
                </div>

                <div className="flex items-center justify-between text-xs text-muted-foreground pt-3 border-t border-border/20">
                  <span>0ms Local V8 Engine</span>
                  <span>100% ECMA-376 Standard</span>
                </div>
              </div>

              <p className="mt-4 text-xs text-muted-foreground text-center">
                Click
                {' '}
                <strong>Download .pptx Deck</strong>
                {' '}
                to compile this script and open in PowerPoint, Keynote, or Google Slides.
              </p>
            </div>
          )}
    </div>
  );
}
