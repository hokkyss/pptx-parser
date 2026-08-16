import { describe, expect, it } from 'vitest';
import { serializeRelationships, type RelationshipEntry } from '../../lib/serializers/relationship-serializer';

describe('Relationships Serializer', () => {
  it('serializes package root relationships correctly', () => {
    const rels: RelationshipEntry[] = [
      {
        id: 'rId1',
        type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
        target: 'ppt/presentation.xml',
      },
      {
        id: 'rId2',
        type: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
        target: 'docProps/core.xml',
      },
    ];

    const xml = serializeRelationships(rels);
    expect(xml).toContain('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
    expect(xml).toContain('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">');
    expect(xml).toContain('Id="rId1"');
    expect(xml).toContain('Target="ppt/presentation.xml"');
    expect(xml).toContain('Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"');
    expect(xml).toContain('Id="rId2"');
    expect(xml).toContain('Target="docProps/core.xml"');
  });

  it('serializes external hyperlink relationships correctly', () => {
    const rels: RelationshipEntry[] = [
      {
        id: 'rId1',
        type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
        target: 'https://example.com',
        targetMode: 'External',
      },
    ];

    const xml = serializeRelationships(rels);
    expect(xml).toContain('TargetMode="External"');
    expect(xml).toContain('Target="https://example.com"');
  });
});
