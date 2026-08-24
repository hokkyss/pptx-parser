import { describe, it, expect } from 'vitest';
import { createThemeResolver } from '../../lib/resolvers/theme-resolver';

describe('createThemeResolver', () => {
  const sampleThemeXml = `
    <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="NRIGroupPowerPointTemplate2023">
      <a:themeElements>
        <a:clrScheme name="NRIGroup">
          <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
          <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
          <a:accent1><a:srgbClr val="000F78"/></a:accent1>
          <a:accent2><a:srgbClr val="3C64AA"/></a:accent2>
          <a:accent3><a:srgbClr val="64AADC"/></a:accent3>
        </a:clrScheme>
        <a:fontScheme name="Yu Gothic UI">
          <a:majorFont><a:latin typeface="Yu Gothic UI"/></a:majorFont>
          <a:minorFont><a:latin typeface="Yu Gothic UI"/></a:minorFont>
        </a:fontScheme>
      </a:themeElements>
    </a:theme>
  `;

  it('should parse theme name', () => {
    const resolver = createThemeResolver();
    const theme = resolver.parseTheme(sampleThemeXml);
    expect(theme.name).toBe('NRIGroupPowerPointTemplate2023');
  });

  it('should parse color scheme accent colors', () => {
    const resolver = createThemeResolver();
    const theme = resolver.parseTheme(sampleThemeXml);

    expect(theme.colorScheme.dk1).toBe('000000');
    expect(theme.colorScheme.lt1).toBe('FFFFFF');
    expect(theme.colorScheme.accent1).toBe('000F78');
    expect(theme.colorScheme.accent2).toBe('3C64AA');
    expect(theme.colorScheme.accent3).toBe('64AADC');
  });

  it('should parse major and minor font scheme', () => {
    const resolver = createThemeResolver();
    const theme = resolver.parseTheme(sampleThemeXml);

    expect(theme.fontScheme.majorFont).toBe('Yu Gothic UI');
    expect(theme.fontScheme.minorFont).toBe('Yu Gothic UI');
  });
});

describe('createThemeResolver fallbacks and custom colors', () => {
  it('falls back to default theme when XML lacks theme root', () => {
    const resolver = createThemeResolver();
    const fallback = resolver.parseTheme('<invalidXml/>');
    expect(fallback.name).toBe('Default Theme');
    expect(fallback.colorScheme.accent1).toBe('4F81BD');
  });

  it('parses custom colors list (<a:custClrLst>)', () => {
    const xml = `
      <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Custom Theme">
        <a:custClrLst>
          <a:custClr name="BrandRed">
            <a:srgbClr val="FF0000"/>
          </a:custClr>
        </a:custClrLst>
      </a:theme>
    `;
    const resolver = createThemeResolver();
    const theme = resolver.parseTheme(xml);
    expect(theme.customColors['BrandRed']).toBe('FF0000');
  });
});

describe('createThemeResolver font fallback without latin typeface', () => {
  it('falls back to default font when majorFont lacks latin typeface', () => {
    const xml = `
      <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Fontless Theme">
        <a:themeElements>
          <a:fontScheme name="Empty Fonts">
            <a:majorFont/>
            <a:minorFont/>
          </a:fontScheme>
        </a:themeElements>
      </a:theme>
    `;
    const resolver = createThemeResolver();
    const theme = resolver.parseTheme(xml);
    expect(theme.fontScheme.majorFont).toBe('Calibri');
    expect(theme.fontScheme.minorFont).toBe('Calibri');
  });
});

describe('createThemeResolver custom color list array', () => {
  it('parses multiple custom colors from array', () => {
    const xml = `
      <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Multi Custom Theme">
        <a:custClrLst>
          <a:custClr name="Color1"><a:srgbClr val="111111"/></a:custClr>
          <a:custClr name="Color2"><a:srgbClr val="222222"/></a:custClr>
        </a:custClrLst>
      </a:theme>
    `;
    const resolver = createThemeResolver();
    const theme = resolver.parseTheme(xml);
    expect(theme.customColors['Color1']).toBe('111111');
    expect(theme.customColors['Color2']).toBe('222222');
  });
});

describe('createThemeResolver edge cases', () => {
  it('covers extractHexColor fallback when color is unknown', () => {
    const xml = `
      <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Unknown Color Theme">
        <a:custClrLst>
          <a:custClr name="OddColor"><a:hslClr/></a:custClr>
        </a:custClrLst>
      </a:theme>
    `;
    const resolver = createThemeResolver();
    const theme = resolver.parseTheme(xml);
    expect(theme.customColors['OddColor']).toBe('000000');
  });

  it('covers extractFontTypeface fallback when latin has no typeface attr', () => {
    const xml = `
      <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Empty Latin Theme">
        <a:themeElements>
          <a:fontScheme name="Empty Latin Scheme">
            <a:majorFont><a:latin/></a:majorFont>
            <a:minorFont><a:latin/></a:minorFont>
          </a:fontScheme>
        </a:themeElements>
      </a:theme>
    `;
    const resolver = createThemeResolver();
    const theme = resolver.parseTheme(xml);
    expect(theme.fontScheme.majorFont).toBe('Calibri');
  });

  it('covers custom parser non-array custClr in parseCustomColors', () => {
    const mockParser = {
      parse: <T>(_xml: string): T => ({
        'a:theme': {
          'a:custClrLst': {
            'a:custClr': { '@_name': 'SingleCustom', 'a:srgbClr': { '@_val': 'ABCDEF' } },
          },
        },
      }) as T,
    };
    const resolver = createThemeResolver(mockParser);
    const theme = resolver.parseTheme('<fake/>');
    expect(theme.customColors['SingleCustom']).toBe('ABCDEF');
  });
});
