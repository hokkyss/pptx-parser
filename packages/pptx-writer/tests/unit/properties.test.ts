import type { PptxMetadata } from '@hokkyss/pptx-core';
import { emu } from '@hokkyss/pptx-core';
import { describe, expect, it } from 'vitest';
import { serializeAppProperties, serializeCoreProperties } from '../../lib/serializers/properties-serializer';
import { serializePresentation } from '../../lib/serializers/presentation-serializer';

describe('Properties Serializer', () => {
  const metadata: PptxMetadata = {
    created: new Date('2026-01-01T10:00:00Z'),
    creator: 'Alice Smith',
    lastModifiedBy: 'Bob Jones',
    modified: new Date('2026-01-02T12:00:00Z'),
    revision: 3,
    slideCount: 5,
    slideHeight: emu(6858000),
    slideWidth: emu(12192000),
    title: 'Quarterly Report',
  };

  it('serializes docProps/core.xml correctly', () => {
    const xml = serializeCoreProperties(metadata);
    expect(xml).toContain('<dc:title>Quarterly Report</dc:title>');
    expect(xml).toContain('<dc:creator>Alice Smith</dc:creator>');
    expect(xml).toContain('<cp:lastModifiedBy>Bob Jones</cp:lastModifiedBy>');
    expect(xml).toContain('<cp:revision>3</cp:revision>');
    expect(xml).toContain('<dcterms:created');
    expect(xml).toContain('2026-01-01T10:00:00');
    expect(xml).toContain('<dcterms:modified');
    expect(xml).toContain('2026-01-02T12:00:00');
  });

  it('serializes docProps/app.xml correctly', () => {
    const xml = serializeAppProperties(metadata);
    expect(xml).toContain('<Slides>5</Slides>');
    expect(xml).toContain('<Application>Microsoft Office PowerPoint</Application>');
    expect(xml).toContain('<PresentationFormat>Widescreen</PresentationFormat>');
  });

  it('handles empty metadata defaults for app and core properties', () => {
    const appXml = serializeAppProperties({});
    expect(appXml).toContain('<Slides>1</Slides>');

    const coreXml = serializeCoreProperties({});
    expect(coreXml).toContain('cp:coreProperties');
  });
});

describe('Presentation Serializer (firstSlideNum)', () => {
  const baseDoc = {
    customXml: [],
    media: [],
    metadata: {
      slideCount: 0,
      slideHeight: emu(6858000),
      slideWidth: emu(12192000),
    },
    slideLayouts: [],
    slideMasters: [],
    slides: [],
    themes: [],
  };

  const baseOptions = { slideRelIds: [] };

  it('omits firstSlideNum attribute when firstSlideNumber is not set', () => {
    const xml = serializePresentation(baseDoc, baseOptions);
    expect(xml).not.toContain('firstSlideNum');
  });

  it('emits firstSlideNum="0" when firstSlideNumber is 0', () => {
    const doc = {
      ...baseDoc,
      metadata: { ...baseDoc.metadata, firstSlideNumber: 0 },
    };
    const xml = serializePresentation(doc, baseOptions);
    expect(xml).toContain('firstSlideNum="0"');
  });

  it('emits firstSlideNum="5" when firstSlideNumber is 5', () => {
    const doc = {
      ...baseDoc,
      metadata: { ...baseDoc.metadata, firstSlideNumber: 5 },
    };
    const xml = serializePresentation(doc, baseOptions);
    expect(xml).toContain('firstSlideNum="5"');
  });
});
