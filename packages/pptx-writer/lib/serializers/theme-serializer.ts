import type { PptxColorScheme, PptxTheme } from '@hokkyss/pptx-core';
import { DEFAULT_THEME_XML } from '../templates/defaults';

/**
 * Serializes a PptxTheme AST into OpenXML `<a:theme>` format.
 * Accurately updates colorScheme, fontScheme, and theme name while preserving
 * full formatScheme fidelity.
 */
export function serializeTheme(theme: PptxTheme): string {
  let xml = theme.rawXml || DEFAULT_THEME_XML;

  // 1. Update theme name: <a:theme ... name="...">
  if (theme.name) {
    xml = xml.replace(/(<a:theme\b[^>]*?\bname=")[^"]*(")/, `$1${theme.name}$2`);
  }

  // 2. Update color scheme name: <a:clrScheme name="...">
  const clrSchemeName = (theme.colorScheme as { name?: string })?.name || theme.name || 'Office';
  xml = xml.replace(/(<a:clrScheme\b[^>]*?\bname=")[^"]*(")/, `$1${clrSchemeName}$2`);

  // 3. Update color scheme values
  if (theme.colorScheme) {
    const clrKeys: (keyof PptxColorScheme)[] = [
      'dk1',
      'lt1',
      'dk2',
      'lt2',
      'accent1',
      'accent2',
      'accent3',
      'accent4',
      'accent5',
      'accent6',
      'hlink',
      'folHlink',
    ];

    for (const key of clrKeys) {
      const val = theme.colorScheme[key];
      if (val) {
        const regex = new RegExp(`(<a:${key}>)[\\s\\S]*?(<\\/a:${key}>)`, 'g');
        if (regex.test(xml)) {
          xml = xml.replace(regex, `$1<a:srgbClr val="${val}"/>$2`);
        }
      }
    }
  }

  // 4. Update font scheme
  if (theme.fontScheme) {
    if (theme.fontScheme.name) {
      xml = xml.replace(/(<a:fontScheme\b[^>]*?\bname=")[^"]*(")/, `$1${theme.fontScheme.name}$2`);
    }
    if (theme.fontScheme.majorFont) {
      xml = xml.replace(
        /(<a:majorFont>[\s\S]*?<a:latin\b[^>]*?\btypeface=")[^"]*(")/,
        `$1${theme.fontScheme.majorFont}$2`,
      );
      xml = xml.replace(
        /(<a:majorFont>[\s\S]*?<a:cs\b[^>]*?\btypeface=")[^"]*(")/,
        `$1${theme.fontScheme.majorFont}$2`,
      );
    }
    if (theme.fontScheme.minorFont) {
      xml = xml.replace(
        /(<a:minorFont>[\s\S]*?<a:latin\b[^>]*?\btypeface=")[^"]*(")/,
        `$1${theme.fontScheme.minorFont}$2`,
      );
      xml = xml.replace(
        /(<a:minorFont>[\s\S]*?<a:cs\b[^>]*?\btypeface=")[^"]*(")/,
        `$1${theme.fontScheme.minorFont}$2`,
      );
    }
  }

  // 5. Update extraClrSchemeLst (registers color palette under Custom in Slide Master view)
  if (theme.colorScheme && clrSchemeName !== 'Office') {
    const extraClrXml = `<a:extraClrSchemeLst><a:extraClrScheme><a:clrScheme name="${clrSchemeName}">`
      + `<a:dk1><a:srgbClr val="${theme.colorScheme.dk1.replace(/^#/, '')}"/></a:dk1>`
      + `<a:lt1><a:srgbClr val="${theme.colorScheme.lt1.replace(/^#/, '')}"/></a:lt1>`
      + `<a:dk2><a:srgbClr val="${theme.colorScheme.dk2.replace(/^#/, '')}"/></a:dk2>`
      + `<a:lt2><a:srgbClr val="${theme.colorScheme.lt2.replace(/^#/, '')}"/></a:lt2>`
      + `<a:accent1><a:srgbClr val="${theme.colorScheme.accent1.replace(/^#/, '')}"/></a:accent1>`
      + `<a:accent2><a:srgbClr val="${theme.colorScheme.accent2.replace(/^#/, '')}"/></a:accent2>`
      + `<a:accent3><a:srgbClr val="${theme.colorScheme.accent3.replace(/^#/, '')}"/></a:accent3>`
      + `<a:accent4><a:srgbClr val="${theme.colorScheme.accent4.replace(/^#/, '')}"/></a:accent4>`
      + `<a:accent5><a:srgbClr val="${theme.colorScheme.accent5.replace(/^#/, '')}"/></a:accent5>`
      + `<a:accent6><a:srgbClr val="${theme.colorScheme.accent6.replace(/^#/, '')}"/></a:accent6>`
      + `<a:hlink><a:srgbClr val="${theme.colorScheme.hlink.replace(/^#/, '')}"/></a:hlink>`
      + `<a:folHlink><a:srgbClr val="${theme.colorScheme.folHlink.replace(/^#/, '')}"/></a:folHlink>`
      + `</a:clrScheme></a:extraClrScheme></a:extraClrSchemeLst>`;

    if (xml.includes('<a:extraClrSchemeLst/>')) {
      xml = xml.replace('<a:extraClrSchemeLst/>', extraClrXml);
    } else if (xml.includes('<a:extraClrSchemeLst>')) {
      xml = xml.replace(/<a:extraClrSchemeLst>[\s\S]*?<\/a:extraClrSchemeLst>/, extraClrXml);
    }
  }

  return xml;
}
