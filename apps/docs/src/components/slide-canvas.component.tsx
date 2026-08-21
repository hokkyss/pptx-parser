/* eslint-disable @eslint-react/no-array-index-key */
import type {
  PptxChart,
  PptxElement,
  PptxTable,
} from '@hokkyss/pptx-core';
import * as PptxModule from '@hokkyss/pptx';
import {
  CaretLeftIcon,
  CaretRightIcon,
  ChartBarIcon,
  DownloadSimpleIcon,
  FilePptIcon,
  ImageIcon,
  WarningCircleIcon,
} from '@phosphor-icons/react';
import { useMemo, useState, useTransition } from 'react';

interface EvalResult {
  error: null | string;
  pres: null | PptxModule.Presentation;
}

interface SlideCanvasProps {
  code: string;
}

/**
 * Slide canvas preview that evaluates code in real-time and renders visual slides with download support.
 * @param root0 Component props
 * @param root0.code TypeScript source code to execute
 * @returns React node
 */
export default function SlideCanvas({ code }: SlideCanvasProps) {
  const [currentSlideIndex, setCurrentSlideIndex] = useState(0);
  const [isDownloading, startDownloadTransition] = useTransition();
  const [downloadError, setDownloadError] = useState<null | string>(null);

  // Evaluate presentation synchronously on every code update
  const evalResult = useMemo(() => evaluatePresentationCode(code), [code]);

  const slides = evalResult.pres?.ast.slides || [];
  const slideCount = slides.length;
  const activeSlide = slides[Math.min(currentSlideIndex, Math.max(0, slideCount - 1))];

  const slideWidthEmu = Number(evalResult.pres?.metadata.slideWidth || 12192000);
  const slideHeightEmu = Number(evalResult.pres?.metadata.slideHeight || 6858000);

  // Determine slide background color
  const backgroundColor = useMemo(() => {
    if (!activeSlide?.background?.fill) return '#0F172A';
    const fill = activeSlide.background.fill;
    if (fill.type === 'solid' && fill.solidColor) {
      return resolveHexColor(fill.solidColor.value, '#0F172A');
    }
    return '#0F172A';
  }, [activeSlide]);

  const handleDownload = () => {
    if (!evalResult.pres) return;
    setDownloadError(null);

    startDownloadTransition(async () => {
      try {
        const buffer = await evalResult.pres!.toArrayBuffer();
        const blob = new Blob([buffer], {
          type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        const title = evalResult.pres!.metadata.title || 'presentation';
        a.download = `${title.replace(/[^a-z0-9_-]/gi, '_')}.pptx`;
        a.click();
        URL.revokeObjectURL(url);
      } catch (err) {
        setDownloadError(err instanceof Error ? err.message : 'Failed to export PPTX file.');
      }
    });
  };

  const connectors = (activeSlide?.elements || []).filter(
    (el): el is { type: 'connector' } & PptxElement => el.type === 'connector',
  );
  const nonConnectors = (activeSlide?.elements || []).filter(
    (el) => el.type !== 'connector',
  );

  return (
    <div className="flex flex-col h-full bg-card rounded-xl border border-border p-5 shadow-sm min-h-0">
      {/* Top Action & Navigation Bar */}
      <div className="flex items-center justify-between pb-4 border-b border-border gap-2">
        <div className="flex items-center gap-3">
          <div className="flex items-center gap-1.5">
            <span
              className={
                evalResult.error
                  ? 'h-2.5 w-2.5 rounded-full bg-rose-500'
                  : 'h-2.5 w-2.5 rounded-full bg-emerald-500 animate-pulse'
              }
            />
            <span className="text-xs font-semibold uppercase tracking-wider text-muted-foreground">
              {evalResult.error ? 'Evaluation Error' : 'Live Slide Preview'}
            </span>
          </div>

          {slideCount > 1 && (
            <div className="flex items-center gap-1 bg-muted/60 px-2 py-0.5 rounded-md border border-border text-xs font-medium">
              <button
                className="hover:text-foreground disabled:opacity-30 transition"
                disabled={currentSlideIndex <= 0}
                onClick={() => setCurrentSlideIndex((prev) => Math.max(0, prev - 1))}
                type="button"
              >
                <CaretLeftIcon className="h-3.5 w-3.5" />
              </button>
              <span className="px-1 text-muted-foreground">
                {currentSlideIndex + 1}
                {' '}
                /
                {' '}
                {slideCount}
              </span>
              <button
                className="hover:text-foreground disabled:opacity-30 transition"
                disabled={currentSlideIndex >= slideCount - 1}
                onClick={() => setCurrentSlideIndex((prev) => Math.min(slideCount - 1, prev + 1))}
                type="button"
              >
                <CaretRightIcon className="h-3.5 w-3.5" />
              </button>
            </div>
          )}
        </div>

        <button
          className="flex items-center gap-1.5 px-3.5 py-1.5 text-xs font-semibold rounded-lg bg-primary hover:bg-primary/90 text-primary-foreground shadow-sm transition active:scale-95 disabled:opacity-50"
          disabled={isDownloading || Boolean(evalResult.error)}
          onClick={handleDownload}
          type="button"
        >
          <DownloadSimpleIcon className="h-4 w-4" />
          {isDownloading ? 'Compiling PPTX...' : 'Download .pptx'}
        </button>
      </div>

      {/* Main Preview / Error Viewport */}
      <div className="flex-1 min-h-0 flex flex-col items-center justify-center p-2 sm:p-4 overflow-auto">
        {evalResult.error
          ? (
              <div className="w-full max-w-lg p-4 rounded-xl bg-rose-500/10 border border-rose-500/30 text-rose-300 text-xs font-mono space-y-2">
                <div className="flex items-center gap-2 font-bold text-rose-400">
                  <WarningCircleIcon className="h-4 w-4 shrink-0" />
                  <span>Compilation Error</span>
                </div>
                <pre className="whitespace-pre-wrap leading-relaxed">{evalResult.error}</pre>
              </div>
            )
          : !activeSlide
              ? (
                  <div className="text-center text-muted-foreground space-y-2">
                    <FilePptIcon className="h-10 w-10 mx-auto text-muted-foreground/40" />
                    <p className="text-sm">No slides found in presentation.</p>
                    <p className="text-xs">
                      Call pres.addSlide() in your code to create a slide.
                    </p>
                  </div>
                )
              : (
                  <div className="w-full max-w-3xl flex flex-col items-center gap-3">
                    {/* 16:9 Responsive Slide Canvas Container */}
                    <div
                      className="w-full aspect-video rounded-xl shadow-2xl relative overflow-hidden border border-border/80 select-none"
                      style={{
                        backgroundColor,
                        containerType: 'inline-size',
                      }}
                    >
                      {/* Connectors SVG Layer */}
                      {connectors.length > 0 && (
                        <svg
                          className="absolute inset-0 w-full h-full pointer-events-none z-0"
                          viewBox={`0 0 ${slideWidthEmu} ${slideHeightEmu}`}
                        >
                          <defs>
                            <marker
                              id="arrow"
                              markerHeight="6"
                              markerWidth="6"
                              orient="auto-start-reverse"
                              refX="5"
                              refY="3"
                              viewBox="0 0 6 6"
                            >
                              <path d="M 0 0 L 6 3 L 0 6 z" fill="currentColor" />
                            </marker>
                          </defs>
                          {connectors.map((connector) => {
                            const x1 = Number(connector.position.x);
                            const y1 = Number(connector.position.y);
                            const x2 = x1 + Number(connector.position.cx);
                            const y2 = y1 + Number(connector.position.cy);
                            const color = resolveHexColor(connector.line?.fill?.solidColor?.value, '#38BDF8');
                            const strokeWidth = Math.max(15000, Number(connector.line?.width || 20000));

                            return (
                              <g key={connector.id} style={{ color }}>
                                <line
                                  markerEnd="url(#arrow)"
                                  stroke={color}
                                  strokeDasharray={connector.line?.dashStyle === 'dash' ? '60000,30000' : undefined}
                                  strokeWidth={strokeWidth}
                                  x1={x1}
                                  x2={x2}
                                  y1={y1}
                                  y2={y2}
                                />
                              </g>
                            );
                          })}
                        </svg>
                      )}

                      {/* Slide Elements Layer */}
                      {nonConnectors.map((element) => {
                        const left = `${(Number(element.position.x) / slideWidthEmu) * 100}%`;
                        const top = `${(Number(element.position.y) / slideHeightEmu) * 100}%`;
                        const width = `${(Number(element.position.cx) / slideWidthEmu) * 100}%`;
                        const height = `${(Number(element.position.cy) / slideHeightEmu) * 100}%`;
                        const rotation = element.rotation ? Number(element.rotation) / 60000 : 0;

                        return (
                          <div
                            className="absolute flex overflow-hidden"
                            key={element.id}
                            style={{
                              height,
                              left,
                              top,
                              transform: rotation ? `rotate(${rotation}deg)` : undefined,
                              width,
                              zIndex: element.zIndex || 1,
                            }}
                          >
                            <RenderSlideElement element={element} />
                          </div>
                        );
                      })}
                    </div>

                    {/* Slide Metadata & Dimension Footer */}
                    <div className="w-full flex items-center justify-between text-[11px] font-mono text-muted-foreground px-1">
                      <span>
                        {`Dimensions: ${(slideWidthEmu / 914400).toFixed(2)}" x ${(slideHeightEmu / 914400).toFixed(2)}" (16:9)`}
                      </span>
                      <span>
                        {`Elements: ${activeSlide.elements?.length || 0}`}
                      </span>
                    </div>
                  </div>
                )}

        {downloadError && (
          <p className="text-xs text-rose-400 mt-2 font-mono">{downloadError}</p>
        )}
      </div>
    </div>
  );
}

/**
 * Safely evaluates user TypeScript / JavaScript code in the browser context with Pptx in scope.
 * @param code Source code string
 * @returns Evaluated Presentation instance or error message
 */
function evaluatePresentationCode(code: string): EvalResult {
  try {
    // Strip import statements if pasted from module files
    const cleanCode = code.replace(/import\s+.*?from\s+['"][^'"]+['"];?/g, '');

    const paramNames = Object.keys(PptxModule);
    const paramValues = Object.values(PptxModule);

    // Create function with all @hokkyss/pptx exports in local scope
    // eslint-disable-next-line @typescript-eslint/no-implied-eval
    const fn = new Function(
      ...paramNames,
      `
      try {
        ${cleanCode}
        if (typeof pres !== 'undefined') return pres;
        return null;
      } catch (err) {
        throw err;
      }
      `,
    );

    // eslint-disable-next-line @typescript-eslint/no-unsafe-call
    const result = fn(...paramValues) as null | PptxModule.Presentation;

    if (!result || typeof result !== 'object' || !result.ast) {
      return {
        error: 'Script must create and assign an instance to "pres" (e.g. const pres = Presentation.create(...);)',
        pres: null,
      };
    }

    return { error: null, pres: result };
  } catch (err) {
    return {
      error: err instanceof Error ? err.message : String(err),
      pres: null,
    };
  }
}

/**
 * Renders a visual chart preview on the slide canvas.
 * @param root0 Component props
 * @param root0.chart Chart data
 * @returns React node
 */
function RenderChart({ chart }: { chart: PptxChart }) {
  const series = chart.series || [];
  const categories = chart.categories || [];
  const maxVal = Math.max(...series.flatMap((s) => s.values || [0]), 1);

  return (
    <div className="w-full h-full rounded-xl bg-slate-900/90 border border-slate-800 p-[1.5cqw] flex flex-col justify-between text-white">
      {chart.title && (
        <div className="flex items-center gap-2 font-bold text-[1.5cqw] text-sky-400">
          <ChartBarIcon className="h-[1.8cqw] w-[1.8cqw]" />
          <span>{chart.title}</span>
        </div>
      )}

      {/* Bar visual layout */}
      <div className="flex-1 flex items-end justify-between gap-[1.5cqw] my-[1cqw] pt-[1cqw] border-b border-slate-700/60 pb-1">
        {categories.map((cat, cIdx) => (
          <div className="flex-1 flex flex-col items-center gap-1 h-full justify-end" key={`chart-cat-${cat}`}>
            <div className="w-full flex items-end justify-center gap-1 h-4/5">
              {series.map((s) => {
                const val = s.values?.[cIdx] ?? 0;
                const heightPct = `${Math.max(5, (val / maxVal) * 100)}%`;
                const seriesColor = s.fill?.solidColor?.value;
                const barColor = resolveHexColor(typeof seriesColor === 'string' ? seriesColor : undefined, '#38BDF8');

                return (
                  <div
                    className="flex-1 rounded-t transition-all hover:opacity-80"
                    key={`bar-${s.name}-${cat}`}
                    style={{
                      backgroundColor: barColor,
                      height: heightPct,
                    }}
                    title={`${s.name}: ${val}`}
                  />
                );
              })}
            </div>
            <span className="text-[0.9cqw] text-slate-400 truncate max-w-full font-mono">{cat}</span>
          </div>
        ))}
      </div>

      {/* Series Legend */}
      <div className="flex items-center justify-center gap-4 text-[0.9cqw]">
        {series.map((s) => {
          const seriesColor = s.fill?.solidColor?.value;
          const legendColor = resolveHexColor(typeof seriesColor === 'string' ? seriesColor : undefined, '#38BDF8');

          return (
            <div className="flex items-center gap-1.5" key={`legend-${s.name}`}>
              <span
                className="h-[0.8cqw] w-[0.8cqw] rounded-full"
                style={{ backgroundColor: legendColor }}
              />
              <span className="text-slate-300">{s.name}</span>
            </div>
          );
        })}
      </div>
    </div>
  );
}

/**
 * Renders individual visual slide elements (shapes, text, tables, charts, pictures).
 * @param root0 Component props
 * @param root0.element PptxElement AST node
 * @returns React node
 */
function RenderSlideElement({ element }: { element: PptxElement }) {
  // 1. Table element
  if (element.table) {
    return <RenderTable table={element.table} />;
  }

  // 2. Chart element
  if (element.chart) {
    return <RenderChart chart={element.chart} />;
  }

  // 3. Picture element
  if (element.type === 'picture' || element.picture) {
    return (
      <div className="w-full h-full rounded-lg bg-muted/40 border border-border/50 flex flex-col items-center justify-center text-muted-foreground p-2">
        <ImageIcon className="h-6 w-6 opacity-60" />
        <span className="text-[1.2cqw] font-mono mt-1 opacity-75">{element.name || 'Image Asset'}</span>
      </div>
    );
  }

  // 4. Shape & Text Box elements
  const isRoundRect = element.shapeType === 'roundRect' || element.geometry?.presetGeometry === 'roundRect';
  const isEllipse = element.shapeType === 'ellipse' || element.geometry?.presetGeometry === 'ellipse';

  const fillBg = element.fill?.type === 'solid' && element.fill.solidColor
    ? resolveHexColor(element.fill.solidColor.value)
    : 'transparent';

  const borderColor = element.line?.fill?.type === 'solid' && element.line.fill.solidColor
    ? resolveHexColor(element.line.fill.solidColor.value)
    : undefined;

  const paragraphs = element.textBody?.paragraphs || [];

  return (
    <div
      className="w-full h-full flex flex-col justify-center px-[1.2cqw] py-[0.8cqw] transition-colors"
      style={{
        backgroundColor: fillBg,
        borderColor,
        borderRadius: isEllipse ? '9999px' : isRoundRect ? '0.6cqw' : '0.2cqw',
        borderStyle: element.line ? (element.line.dashStyle === 'dash' ? 'dashed' : 'solid') : 'none',
        borderWidth: borderColor ? '0.15cqw' : 0,
      }}
    >
      {paragraphs.map((paragraph, pIdx) => {
        const align = paragraph.properties?.alignment || 'left';
        const textAlignClass = align === 'center'
          ? 'text-center'
          : align === 'right'
            ? 'text-right'
            : 'text-left';

        return (
          <p className={`leading-snug ${textAlignClass}`} key={`p-${element.id}-${pIdx}`}>
            {paragraph.runs.map((run, rIdx) => {
              const runColor = run.properties?.color;
              const color = resolveHexColor(typeof runColor === 'string' ? runColor : undefined, '#FFFFFF');
              const fontSizePt = (Number(run.properties?.fontSize) || 1400) / 100;
              const fontSizeCqw = `${Math.max(0.8, fontSizePt * 0.104166)}cqw`;

              return (
                <span
                  key={`run-${element.id}-${pIdx}-${rIdx}`}
                  style={{
                    color,
                    fontFamily: run.properties?.fontFamily || 'inherit',
                    fontSize: fontSizeCqw,
                    fontStyle: run.properties?.italic ? 'italic' : 'normal',
                    fontWeight: run.properties?.bold ? '700' : '400',
                    textDecoration: run.properties?.underline ? 'underline' : 'none',
                  }}
                >
                  {run.text}
                </span>
              );
            })}
          </p>
        );
      })}
    </div>
  );
}

/**
 * Renders an interactive table grid on the slide canvas.
 * @param root0 Component props
 * @param root0.table Table data
 * @returns React node
 */
function RenderTable({ table }: { table: PptxTable }) {
  return (
    <div className="w-full h-full rounded-lg overflow-hidden border border-border/50 bg-card/60">
      <table className="w-full h-full border-collapse">
        <tbody>
          {table.rows.map((row, rIdx) => (
            <tr key={`tbl-row-${rIdx}`}>
              {row.cells.map((cell, cIdx) => {
                const cellColor = cell.properties?.fill?.type === 'solid'
                  ? cell.properties.fill.solidColor?.value
                  : undefined;
                const fill = cellColor
                  ? resolveHexColor(cellColor)
                  : rIdx === 0
                    ? '#1E293B'
                    : 'transparent';

                const runs = cell.textBody?.paragraphs?.[0]?.runs || [];
                const align = cell.textBody?.paragraphs?.[0]?.properties?.alignment || (rIdx === 0 ? 'left' : 'left');

                return (
                  <td
                    className={`px-[1cqw] py-[0.5cqw] border border-border/40 text-[1.1cqw] ${
                      align === 'center' ? 'text-center' : align === 'right' ? 'text-right' : 'text-left'
                    }`}
                    key={`tbl-cell-${rIdx}-${cIdx}`}
                    style={{ backgroundColor: fill }}
                  >
                    {runs.map((r, i) => {
                      const textColor = r.properties?.color;
                      return (
                        <span
                          key={`tbl-cell-run-${rIdx}-${cIdx}-${i}`}
                          style={{
                            color: resolveHexColor(typeof textColor === 'string' ? textColor : undefined, rIdx === 0 ? '#38BDF8' : '#F8FAFC'),
                            fontWeight: r.properties?.bold || rIdx === 0 ? '700' : '400',
                          }}
                        >
                          {r.text}
                        </span>
                      );
                    })}
                  </td>
                );
              })}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

/**
 * Resolves hex color string to CSS color format (#RRGGBB).
 * @param color Color hex or named color
 * @param fallback Fallback color
 * @returns Formatted hex string
 */
function resolveHexColor(color?: string, fallback = 'transparent'): string {
  if (!color) return fallback;
  const clean = color.trim().replace(/^#/, '');
  if (/^[0-9a-fA-F]{3,8}$/.test(clean)) {
    return `#${clean}`;
  }
  return color;
}
