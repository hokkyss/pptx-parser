import type { Relationship } from '@hokkyss/pptx-core';
import { el, serializeXml } from '../xml/xml-element';

export type RelationshipEntry = { id: string; target: string; type: string } & Partial<Relationship>;

/**
 * Serializes a list of Relationships to a `.rels` XML string.
 * @param relationships Array of Relationship objects.
 * @returns Complete `.rels` XML string.
 */
export function serializeRelationships(relationships: RelationshipEntry[]): string {
  const relElements = relationships.map((rel) => {
    const attrs: Record<string, string | undefined> = {
      Id: rel.id,
      Target: rel.target,
      Type: rel.type,
    };
    if (rel.targetMode) {
      attrs.TargetMode = rel.targetMode;
    }
    return el('Relationship', attrs);
  });

  const root = el('Relationships', {
    xmlns: 'http://schemas.openxmlformats.org/package/2006/relationships',
  }, relElements);

  return serializeXml(root);
}
