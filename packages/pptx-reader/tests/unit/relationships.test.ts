import { describe, it, expect } from 'vitest';
import { createRelationshipResolver } from '../../lib/resolvers/relationship-resolver';

describe('createRelationshipResolver', () => {
  const sampleRelsXml = `
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
      <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
      <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
    </Relationships>
  `;

  it('should parse relationship target and id', () => {
    const resolver = createRelationshipResolver(sampleRelsXml, 'ppt/slides/slide1.xml');
    const rel1 = resolver.getRelationship('rId1');

    expect(rel1).toBeDefined();
    expect(rel1?.id).toBe('rId1');
    expect(rel1?.target).toBe('../slideLayouts/slideLayout1.xml');
  });

  it('should resolve relative path correctly relative to source path', () => {
    const resolver = createRelationshipResolver(sampleRelsXml, 'ppt/slides/slide1.xml');
    const rel1 = resolver.getRelationship('rId1');

    expect(rel1?.resolvedTarget).toBe('ppt/slideLayouts/slideLayout1.xml');
  });

  it('should filter relationships by type suffix', () => {
    const resolver = createRelationshipResolver(sampleRelsXml, 'ppt/slides/slide1.xml');

    const layoutRels = resolver.getRelationshipsByType('slideLayout');
    expect(layoutRels.length).toBe(1);
    expect(layoutRels[0].resolvedTarget).toBe('ppt/slideLayouts/slideLayout1.xml');

    const imageRels = resolver.getRelationshipsByType('image');
    expect(imageRels.length).toBe(1);
    expect(imageRels[0].resolvedTarget).toBe('ppt/media/image1.png');
  });
});

describe('createRelationshipResolver extended APIs', () => {
  it('covers getTarget, getAll, slide, and slideMaster filters', () => {
    const relsXml = `
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
        <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
      </Relationships>
    `;
    const resolver = createRelationshipResolver(relsXml, 'ppt/presentation.xml');
    expect(resolver.getAll()).toHaveLength(2);
    expect(resolver.getTarget('rId1')).toBe('ppt/slides/slide1.xml');
    expect(resolver.getTarget('rIdNonExistent')).toBeUndefined();
    expect(resolver.getRelationshipsByType('slide')).toHaveLength(1);
    expect(resolver.getRelationshipsByType('slideMaster')).toHaveLength(1);
    expect(resolver.getRelationship('ppt/presentation.xml', 'rId1')).toBeDefined();
  });
});

describe('createRelationshipResolver addRelationships', () => {
  it('adds custom relationships dynamically', () => {
    const resolver = createRelationshipResolver('', 'ppt/slides/slide1.xml');
    resolver.addRelationships('ppt/slides/slide1.xml', [
      {
        id: 'rId10',
        target: '../media/image10.png',
        type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
      },
    ]);

    expect(resolver.getTarget('rId10')).toBe('ppt/media/image10.png');
    expect(resolver.getRelationship('ppt/slides/slide1.xml', 'rId10')?.resolvedTarget).toBe('ppt/media/image10.png');
  });
});

describe('createRelationshipResolver single relationship and absolute path', () => {
  it('parses single relationship and resolves absolute targets', () => {
    const singleRelXml = `
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="/ppt/media/image1.png"/>
      </Relationships>
    `;
    const resolver = createRelationshipResolver(singleRelXml, 'ppt/slides/slide1.xml');
    expect(resolver.getTarget('rId1')).toBe('ppt/media/image1.png');
  });
});

describe('createRelationshipResolver with custom parser returning non-array Relationship', () => {
  it('wraps non-array Relationship into array', () => {
    const mockParser = {
      parse: <T>(_xml: string): T => ({
        Relationships: {
          Relationship: {
            '@_Id': 'rId99',
            '@_Target': 'target99.xml',
            '@_Type': 'http://example.com/type',
          },
        },
      }) as T,
    };
    const resolver = createRelationshipResolver('<fake/>', 'ppt/presentation.xml', mockParser);
    expect(resolver.getTarget('rId99')).toBe('ppt/target99.xml');
  });
});
