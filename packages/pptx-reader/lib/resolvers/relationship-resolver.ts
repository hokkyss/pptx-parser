import type { Relationship, RelationshipResolver } from '@hokkyss/pptx-core';
import type { XmlParser } from '@hokkyss/pptx-core';
import { defaultXmlParser } from '../xml/xml-parser';

export type { Relationship, RelationshipResolver };

/**
 * Helper resolving relative target paths against a source base file path.
 * @param sourcePath Source file path (e.g. `'ppt/slides/slide1.xml'`).
 * @param target Relative target path (e.g. `'../media/image1.png'`).
 * @returns Resolved package-relative path string (e.g. `'ppt/media/image1.png'`).
 */
function resolvePath(sourcePath: string, target: string): string {
  if (!target) return '';
  if (target.startsWith('/')) {
    return target.slice(1);
  }

  const sourceDirParts = sourcePath ? sourcePath.split('/').slice(0, -1) : [];
  const targetParts = target.split('/');

  const resultParts = [...sourceDirParts];
  for (const part of targetParts) {
    if (part === '.' || part === '') continue;
    if (part === '..') {
      resultParts.pop();
    } else {
      resultParts.push(part);
    }
  }

  return resultParts.join('/');
}

/**
 * Creates a relationship resolver mapping OpenXML relationship IDs (`rId1`, `rId2`) to resource target paths.
 * @param relsXml Optional content of the `.rels` XML file.
 * @param sourcePath File path that owns this `.rels` file (used to resolve relative targets).
 * @param parser Optional custom `XmlParser` instance.
 * @returns Frozen `RelationshipResolver` instance.
 * @example
 * ```ts
 * const relResolver = createRelationshipResolver(relsXml, 'ppt/slides/slide1.xml');
 * const layoutRel = relResolver.getRelationshipsByType('slideLayout')[0];
 * console.log(layoutRel.resolvedTarget); // 'ppt/slideLayouts/slideLayout1.xml'
 * ```
 */
export function createRelationshipResolver(
  relsXml?: string,
  sourcePath: string = '',
  parser: XmlParser = defaultXmlParser,
): RelationshipResolver {
  const relationships = new Map<string, Relationship>();
  const sourceMap = new Map<string, Map<string, Relationship>>();

  const parseRels = (xml: string, srcPath: string = ''): Relationship[] => {
    const parsed = parser.parse<Record<string, unknown>>(xml);
    const relsContainer = parsed['Relationships'] as Record<string, unknown> | undefined;
    if (!relsContainer) return [];

    let relList = relsContainer['Relationship'];
    if (!relList) return [];
    if (!Array.isArray(relList)) {
      relList = [relList];
    }

    const parsedRels: Relationship[] = [];

    for (const relRaw of relList as Record<string, unknown>[]) {
      const id = relRaw['@_Id'] as string;
      const type = relRaw['@_Type'] as string;
      const target = relRaw['@_Target'] as string;
      const targetMode = relRaw['@_TargetMode'] as string | undefined;

      if (!id || !target) continue;

      const resolvedTarget = targetMode === 'External' ? target : resolvePath(srcPath, target);

      const rel: Relationship = {
        id,
        resolvedTarget,
        target,
        targetMode,
        type,
      };

      relationships.set(id, rel);
      parsedRels.push(rel);

      if (srcPath) {
        if (!sourceMap.has(srcPath)) {
          sourceMap.set(srcPath, new Map());
        }
        sourceMap.get(srcPath)!.set(id, rel);
      }
    }

    return parsedRels;
  };

  if (relsXml) {
    parseRels(relsXml, sourcePath);
  }

  return Object.freeze({
    addRelationships(srcPath: string, rels: { id: string; target: string; type: string }[]): void {
      if (!sourceMap.has(srcPath)) {
        sourceMap.set(srcPath, new Map());
      }
      const map = sourceMap.get(srcPath)!;

      for (const r of rels) {
        const resolvedTarget = resolvePath(srcPath, r.target);
        const rel: Relationship = {
          id: r.id,
          resolvedTarget,
          target: r.target,
          type: r.type,
        };
        map.set(r.id, rel);
        relationships.set(r.id, rel);
      }
    },
    getAll(): Relationship[] {
      return Array.from(relationships.values());
    },
    getRelationship(sourcePathOrId: string, id?: string): Relationship | undefined {
      if (id !== undefined) {
        return sourceMap.get(sourcePathOrId)?.get(id);
      }
      return relationships.get(sourcePathOrId);
    },
    getRelationshipsByType(typeKeyword: string): Relationship[] {
      return Array.from(relationships.values()).filter((rel) => {
        if (typeKeyword === 'slide') {
          return rel.type.endsWith('/slide');
        }
        if (typeKeyword === 'slideMaster') {
          return rel.type.endsWith('/slideMaster');
        }
        if (typeKeyword === 'slideLayout') {
          return rel.type.endsWith('/slideLayout');
        }
        return rel.type.includes(typeKeyword);
      });
    },
    getTarget(id: string): string | undefined {
      return relationships.get(id)?.resolvedTarget;
    },
    parseRels,
  });
}
