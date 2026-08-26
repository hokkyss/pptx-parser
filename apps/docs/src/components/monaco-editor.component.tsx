import * as monaco from 'monaco-editor';
import editorWorker from 'monaco-editor/esm/vs/editor/editor.worker?worker';
import tsWorker from 'monaco-editor/esm/vs/language/typescript/ts.worker?worker';
import { useEffect, useRef } from 'react';

interface MonacoEditorProps {
  language?: string;
  onChange?: (val: string) => void;
  readOnly?: boolean;
  value: string;
}

// Configure Monaco Environment for Vite web worker bundling
if (typeof window !== 'undefined') {
  // MonacoEnvironment is global on window
  window.MonacoEnvironment = {
    getWorker(_: unknown, label: string) {
      if (label === 'typescript' || label === 'javascript') {
        return new tsWorker();
      }
      return new editorWorker();
    },
  };
}

// Setup TypeScript compiler options and extra type definitions globally
monaco.languages.typescript.typescriptDefaults.setCompilerOptions({
  allowJs: true,
  allowNonTextFiles: true,
  allowSyntheticDefaultImports: true,
  checkJs: false,
  esModuleInterop: true,
  jsx: monaco.languages.typescript.JsxEmit.React,
  module: monaco.languages.typescript.ModuleKind.ESNext,
  moduleResolution: monaco.languages.typescript.ModuleResolutionKind.NodeJs,
  noEmit: true,
  target: monaco.languages.typescript.ScriptTarget.ES2020,
  typeRoots: ['node_modules/@types'],
});

monaco.languages.typescript.typescriptDefaults.addExtraLib(
  `
declare module '@hokkyss/pptx' {
  export type Inches = number & { readonly __brand: 'Inches' };
  export type Points = number & { readonly __brand: 'Points' };
  export type Degrees = number & { readonly __brand: 'Degrees' };
  export type Emu = number & { readonly __brand: 'Emu' };
  export type HundredthsPoint = number & { readonly __brand: 'HundredthsPoint' };
  export type ThousandthsPercent = number & { readonly __brand: 'ThousandthsPercent' };
  export type EmuDegree = number & { readonly __brand: 'EmuDegree' };

  export function inches(n: number): Inches;
  export function points(n: number): Points;
  export function degrees(n: number): Degrees;
  export function emu(n: number): Emu;
  export function cm(n: number): Inches;
  export function percent(n: number): ThousandthsPercent;
  export function decimal(n: number): ThousandthsPercent;

  export function inchesToEmu(val: Inches | number): Emu;
  export function emuToInches(val: Emu | number): Inches;
  export function pointsToEmu(val: Points | number): Emu;
  export function emuToPoints(val: Emu | number): Points;
  export function cmToEmu(val: number): Emu;
  export function emuToCm(val: Emu | number): number;
  export function pxToEmu(val: number): Emu;
  export function emuToPx(val: Emu | number): number;
  export function degreesToEmuDegree(val: Degrees | number): EmuDegree;
  export function rotationToDegrees(val: EmuDegree | number): Degrees;
  export function pointsToHundredthsPoint(val: Points | number): HundredthsPoint;
  export function hundredthsPointToPoints(val: HundredthsPoint | number): Points;
  export function percentToThousandthsPercent(val: number): ThousandthsPercent;
  export function thousandthsPercentToPercent(val: ThousandthsPercent | number): number;
  export function decimalToThousandthsPercent(val: number): ThousandthsPercent;
  export function thousandthsPercentToDecimal(val: ThousandthsPercent | number): number;

  export interface AddTextOptions {
    x?: Inches;
    y?: Inches;
    w?: Inches;
    h?: Inches;
    fontSize?: Points;
    font?: string;
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    color?: string;
    align?: 'left' | 'center' | 'right' | 'justify';
    verticalAlignment?: 'top' | 'middle' | 'bottom';
    fill?: string;
    rotation?: Degrees;
    zIndex?: number;
    id?: string;
  }

  export interface AddShapeOptions {
    text?: string;
    x?: Inches;
    y?: Inches;
    w?: Inches;
    h?: Inches;
    fill?: string | { type: 'solid' | 'linear'; color?: string; angle?: Degrees; stops?: Array<{ color: string; position: ThousandthsPercent | number }> };
    line?: { color?: string; width?: Inches; dashStyle?: 'solid' | 'dash' | 'dot' };
    textOptions?: AddTextOptions;
    rotation?: Degrees;
    zIndex?: number;
    id?: string;
    name?: string;
  }

  export interface AddConnectorOptions {
    from: { x: Inches; y: Inches };
    to: { x: Inches; y: Inches };
    color?: string;
    width?: Inches;
    dashStyle?: 'solid' | 'dash' | 'dot';
    shapeType?: string;
    zIndex?: number;
    id?: string;
    name?: string;
  }

  export interface AddTableOptions {
    x?: Inches;
    y?: Inches;
    w?: Inches;
    h?: Inches;
    colWidths?: Inches[];
    columns?: Array<{ w?: Inches } | Inches>;
    header?: boolean | { bold?: boolean; color?: string; fill?: string; fontSize?: Points; font?: string; align?: 'left' | 'center' | 'right' };
    id?: string;
    name?: string;
    zIndex?: number;
  }

  export interface ChartSeriesConfig {
    name: string;
    values: number[];
    color?: string;
  }

  export interface AddChartOptions {
    chartType: 'bar' | 'col' | 'line' | 'pie' | 'doughnut' | 'area';
    categories: string[];
    series: ChartSeriesConfig[];
    title?: string;
    x?: Inches;
    y?: Inches;
    w?: Inches;
    h?: Inches;
    id?: string;
    name?: string;
    zIndex?: number;
  }

  export class Slide {
    addText(content: string | any[], options?: AddTextOptions): this;
    addShape(type: string, options?: AddShapeOptions): this;
    addConnector(options: AddConnectorOptions): this;
    addTable(matrix: any[][], options?: AddTableOptions): this;
    addChart(options: AddChartOptions): this;
    addImage(data: ArrayBuffer | Uint8Array, options?: any): this;
    addGroup(options: any, callback: (group: any) => void): this;
    setBackground(fill: string | any): this;
    setNotes(notes: string | any[]): this;
  }

  export class Presentation {
    static create(options?: { title?: string; author?: string; width?: Inches; height?: Inches; firstSlideNumber?: number }): Presentation;
    static load(buffer: ArrayBuffer | Uint8Array): Promise<Presentation>;
    addSlide(options?: { layout?: string; master?: string }): Slide;
    getSlides(): Slide[];
    getSlide(index: number): Slide | undefined;
    setThemeColors(colors: { accent1?: string; accent2?: string; accent3?: string; accent4?: string; accent5?: string; accent6?: string; dk1?: string; dk2?: string; lt1?: string; lt2?: string }): this;
    setThemeFonts(fonts: { majorFont?: string; minorFont?: string }): this;
    save(filePath: string): Promise<void>;
    toArrayBuffer(): Promise<ArrayBuffer>;
    toBuffer(): Promise<Uint8Array>;
    get ast(): any;
    get metadata(): any;
  }
}

declare const Presentation: typeof import('@hokkyss/pptx').Presentation;
declare const inches: typeof import('@hokkyss/pptx').inches;
declare const points: typeof import('@hokkyss/pptx').points;
declare const degrees: typeof import('@hokkyss/pptx').degrees;
declare const emu: typeof import('@hokkyss/pptx').emu;
declare const cm: typeof import('@hokkyss/pptx').cm;
declare const percent: typeof import('@hokkyss/pptx').percent;
declare const decimal: typeof import('@hokkyss/pptx').decimal;
`,
  'file:///node_modules/@types/hokkyss__pptx/index.d.ts',
);

/**
 * Monaco code editor component for interactive TypeScript editing.
 * @param root0 Component props
 * @param root0.language Editor syntax language
 * @param root0.onChange Content change callback
 * @param root0.readOnly Read-only state flag
 * @param root0.value Initial / controlled code string
 * @returns React node
 */
export default function MonacoEditor({
  language = 'typescript',
  onChange,
  readOnly = false,
  value,
}: MonacoEditorProps) {
  const containerRef = useRef<HTMLDivElement>(null);
  const editorRef = useRef<monaco.editor.IStandaloneCodeEditor | null>(null);

  useEffect(() => {
    if (!containerRef.current) return;

    // Use an explicit file URI so Monaco TypeScript language service links the model
    const modelUri = monaco.Uri.parse('file:///playground.ts');
    let model = monaco.editor.getModel(modelUri);

    if (!model) {
      model = monaco.editor.createModel(value, language, modelUri);
    } else {
      model.setValue(value);
    }

    const editor = monaco.editor.create(containerRef.current, {
      automaticLayout: true,
      fontFamily: 'ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace',
      fontSize: 13,
      lineNumbers: 'on',
      minimap: { enabled: false },
      model,
      padding: { bottom: 12, top: 12 },
      readOnly,
      roundedSelection: true,
      scrollBeyondLastLine: false,
      tabSize: 2,
      theme: 'vs-dark',
      wordWrap: 'on',
    });

    editorRef.current = editor;

    const disposable = editor.onDidChangeModelContent(() => {
      onChange?.(editor.getValue());
    });

    return () => {
      disposable.dispose();
      editor.dispose();
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  useEffect(() => {
    if (editorRef.current) {
      const currentModel = editorRef.current.getModel();
      if (currentModel && currentModel.getValue() !== value) {
        currentModel.setValue(value);
      }
    }
  }, [value]);

  return <div className="w-full h-full min-h-112.5 rounded-lg overflow-hidden border border-border/60" ref={containerRef} />;
}
