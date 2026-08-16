import type { Relationship } from '@hokkyss/pptx-core';
import { serializeXml } from '../xml/xml-builder';

export type RelationshipEntry = { id: string; target: string; type: string } & Partial<Relationship>;

/**
 * Serializes a list of Relationships to a `.rels` XML string.
 * @param relationships Array of Relationship objects.
 * @returns Complete `.rels` XML string.
 */
export function serializeRelationships(relationships: RelationshipEntry[]): string {
  const relElements = relationships.map((rel) => {
    const relNode: Record<string, unknown> = {
      '@_Id': rel.id,
      '@_Target': rel.target,
      '@_Type': rel.type,
    };
    if (rel.targetMode) {
      relNode['@_TargetMode'] = rel.targetMode;
    }
    return relNode;
  });

  const root = {
    Relationships: {
      '@_xmlns': 'http://schemas.openxmlformats.org/package/2006/relationships',
      Relationship: relElements,
    },
  };

  return serializeXml(root);
}
