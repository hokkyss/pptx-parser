import { describe, expect, it } from 'vitest';
import { type ContentTypesOptions, serializeContentTypes } from '../../lib/serializers/content-types-serializer';

describe('Content Types Serializer', () => {
  it('serializes [Content_Types].xml with slide, layout, master, and theme overrides', () => {
    const options: ContentTypesOptions = {
      hasCharts: false,
      layoutCount: 1,
      masterCount: 1,
      mediaExtensions: ['png', 'jpeg'],
      slideCount: 2,
      themeCount: 1,
    };

    const xml = serializeContentTypes(options);
    expect(xml).toContain('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">');
    expect(xml).toContain('Extension="rels"');
    expect(xml).toContain('ContentType="application/vnd.openxmlformats-package.relationships+xml"');
    expect(xml).toContain('Extension="xml"');
    expect(xml).toContain('ContentType="application/xml"');
    expect(xml).toContain('Extension="png"');
    expect(xml).toContain('ContentType="image/png"');
    expect(xml).toContain('Extension="jpeg"');
    expect(xml).toContain('ContentType="image/jpeg"');
    expect(xml).toContain('PartName="/ppt/presentation.xml"');
    expect(xml).toContain('PartName="/ppt/slides/slide1.xml"');
    expect(xml).toContain('PartName="/ppt/slides/slide2.xml"');
    expect(xml).toContain('PartName="/ppt/slideLayouts/slideLayout1.xml"');
    expect(xml).toContain('PartName="/ppt/slideMasters/slideMaster1.xml"');
    expect(xml).toContain('PartName="/ppt/theme/theme1.xml"');
  });
});

describe('Content Types Serializer extended options', () => {
  it('serializes custom extensions, explicit names, and customPartOverrides', () => {
    const xml = serializeContentTypes({
      mediaExtensions: ['svg', 'webp', 'mp4'],
      layoutNames: ['slideLayout1'],
      masterNames: ['slideMaster1'],
      themeNames: ['theme1'],
      customPartOverrides: [{ contentType: 'application/xml', partName: 'customXml/item1.xml' }],
    });
    expect(xml).toContain('Extension="svg"');
    expect(xml).toContain('PartName="/customXml/item1.xml"');
    expect(xml).toContain('PartName="/ppt/slideLayouts/slideLayout1.xml"');
  });
});
