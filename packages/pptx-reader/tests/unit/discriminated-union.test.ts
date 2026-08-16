import { describe, it, expect } from 'vitest';
import type { Emu, EmuDegree, PptxElement } from '../../lib/types/ast.js';

describe('Discriminated Union PptxElement', () => {
  it('should narrow types via elementType discriminator', () => {
    const mockTableElement: PptxElement = {
      elementType: 'table',
      id: 'table1',
      isVisible: true,
      name: 'Table 1',
      position: { cx: 100 as Emu, cy: 100 as Emu, x: 0 as Emu, y: 0 as Emu },
      rotation: 0 as EmuDegree,
      table: {
        columnWidths: [100 as Emu],
        rows: [{ cells: [{ textBody: { bodyProperties: {}, paragraphs: [] } }], height: 50 as Emu }],
      },
      type: 'graphicFrame',
      zIndex: 0,
    };

    if (mockTableElement.elementType === 'table') {
      // TypeScript narrowing test
      expect(mockTableElement.table.rows).toHaveLength(1);
    } else {
      throw new Error('Type narrowing failed');
    }
  });

  it('should distinguish text box shapes from generic visual shapes', () => {
    const mockTextBox: PptxElement = {
      elementType: 'shape',
      id: 'tb1',
      isTextBox: true,
      isVisible: true,
      name: 'Text Box 1',
      position: { cx: 100 as Emu, cy: 100 as Emu, x: 0 as Emu, y: 0 as Emu },
      rotation: 0 as EmuDegree,
      textBody: {
        bodyProperties: {},
        paragraphs: [{ properties: {}, runs: [{ properties: {}, text: 'Hello' }] }],
      },
      type: 'shape',
      zIndex: 1,
    };

    expect(mockTextBox.elementType).toBe('shape');
    if (mockTextBox.elementType === 'shape') {
      expect(mockTextBox.isTextBox).toBe(true);
      expect(mockTextBox.textBody?.paragraphs[0].runs[0].text).toBe('Hello');
    }
  });
});
