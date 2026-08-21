import * as monaco from 'monaco-editor';
import { useEffect, useRef } from 'react';

interface MonacoEditorProps {
  language?: string;
  onChange?: (val: string) => void;
  readOnly?: boolean;
  value: string;
}

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

    // Configure TypeScript compiler options in Monaco
    monaco.languages.typescript.typescriptDefaults.setCompilerOptions({
      allowJs: true,
      allowNonTextFiles: true,
      esModuleInterop: true,
      jsx: monaco.languages.typescript.JsxEmit.React,
      module: monaco.languages.typescript.ModuleKind.ESNext,
      moduleResolution: monaco.languages.typescript.ModuleResolutionKind.NodeJs,
      noEmit: true,
      target: monaco.languages.typescript.ScriptTarget.ES2020,
    });

    // Provide @hokkyss/pptx type definitions for autocomplete
    monaco.languages.typescript.typescriptDefaults.addExtraLib(
      `
      declare module '@hokkyss/pptx' {
        export type Inches = number & { readonly __brand: 'Inches' };
        export type Points = number & { readonly __brand: 'Points' };
        export type Degrees = number & { readonly __brand: 'Degrees' };
        export type Emu = number & { readonly __brand: 'Emu' };

        export function inches(n: number): Inches;
        export function points(n: number): Points;
        export function degrees(n: number): Degrees;
        export function emu(n: number): Emu;

        export interface ShapeAttachment {
          shapeId: string;
          position: 'top' | 'bottom' | 'left' | 'right';
        }

        export class Slide {
          addText(content: any, options?: any): this;
          addShape(type: string, options?: any): this;
          addConnector(options: any): this;
          addTable(options: any): this;
          addChart(type: string, options: any): this;
          addImage(options: any): this;
          setBackground(fill: any): this;
          setTransition(type: string, options?: any): this;
          setNotes(notes: any): this;
        }

        export class Presentation {
          static create(options?: any): Presentation;
          static load(buffer: any): Promise<Presentation>;
          addSlide(options?: any): Slide;
          setThemeColors(colors: any): this;
          setThemeFonts(fonts: any): this;
          setSize(size: any): this;
          save(filePath: string): Promise<void>;
          toArrayBuffer(): Promise<ArrayBuffer>;
          toAst(): any;
        }
      }
      `,
      '@hokkyss/pptx/index.d.ts',
    );

    const editor = monaco.editor.create(containerRef.current, {
      automaticLayout: true,
      fontFamily: 'ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace',
      fontSize: 13,
      language,
      lineNumbers: 'on',
      minimap: { enabled: false },
      padding: { bottom: 12, top: 12 },
      readOnly,
      roundedSelection: true,
      scrollBeyondLastLine: false,
      tabSize: 2,
      theme: 'vs-dark',
      value,
      wordWrap: 'on',
    });

    editorRef.current = editor;

    editor.onDidChangeModelContent(() => {
      onChange?.(editor.getValue());
    });

    return () => {
      editor.dispose();
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  useEffect(() => {
    if (editorRef.current && editorRef.current.getValue() !== value) {
      editorRef.current.setValue(value);
    }
  }, [value]);

  return <div className="w-full h-full min-h-112.5 rounded-lg overflow-hidden border border-border/60" ref={containerRef} />;
}
