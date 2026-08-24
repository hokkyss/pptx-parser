import { describe, expect, it } from 'vitest';
import type { PptxTheme } from '@hokkyss/pptx-core';
import { serializeTheme } from '../../lib/serializers/theme-serializer';

describe('Theme Serializer', () => {
  it('serializes theme with existing extraClrSchemeLst replacing existing entries', () => {
    const theme: PptxTheme = {
      id: 'theme1.xml',
      name: 'Custom Theme',
      rawXml: '<?xml version="1.0" encoding="UTF-8"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:extraClrSchemeLst><a:extraClrScheme/></a:extraClrSchemeLst></a:themeElements></a:theme>',
      colorScheme: {
        dk1: '000000',
        lt1: 'FFFFFF',
        dk2: '111111',
        lt2: 'EEEEEE',
        accent1: '0284C7',
        accent2: '10B981',
        accent3: 'F59E0B',
        accent4: 'EF4444',
        accent5: '8B5CF6',
        accent6: 'EC4899',
        hlink: '3B82F6',
        folHlink: '6366F1',
      },
    };

    const xml = serializeTheme(theme);
    expect(xml).toContain('val="0284C7"');
    expect(xml).toContain('<a:extraClrScheme>');
  });

  it('serializes custom color scheme name object', () => {
    const theme: PptxTheme = {
      id: 'theme1.xml',
      colorScheme: {
        accent1: '0284C7',
        accent2: '10B981',
        accent3: 'F59E0B',
        accent4: 'EF4444',
        accent5: '8B5CF6',
        accent6: 'EC4899',
        dk1: '000000',
        dk2: '111111',
        folHlink: '6366F1',
        hlink: '3B82F6',
        lt1: 'FFFFFF',
        lt2: 'EEEEEE',
        name: 'CustomPalette',
      },
    };
    const xml = serializeTheme(theme);
    expect(xml).toContain('name="CustomPalette"');
  });
});
