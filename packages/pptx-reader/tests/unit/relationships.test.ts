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
