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
