import type { PptxSlideLayout, PptxSlideMaster } from '@hokkyss/pptx-core';
import type { XmlParser } from '@hokkyss/pptx-core';
import type { ZipReader } from '@hokkyss/pptx-core';
import { getXmlChild, parseShapes } from '../parsers/shape-parser';
import { defaultXmlParser } from '../xml/xml-parser';
import { createRelationshipResolver } from './relationship-resolver';

export interface MasterLayoutResolver {
  /**
   * Parses all slide layout XML files in the presentation package into `PptxSlideLayout` structures.
   * @param layoutPaths Array of zip paths to slide layout files (e.g. `['ppt/slideLayouts/slideLayout1.xml']`).
   * @returns Array of parsed `PptxSlideLayout` objects with template elements, layout type, layout name, and parent master ID link.
   */
  parseLayouts(layoutPaths: string[]): PptxSlideLayout[];

  /**
   * Parses all slide master XML files in the presentation package into `PptxSlideMaster` structures.
   * @param masterPaths Array of zip paths to slide master files (e.g. `['ppt/slideMasters/slideMaster1.xml']`).
   * @returns Array of parsed `PptxSlideMaster` objects with background elements and associated layout IDs.
   */
  parseMasters(masterPaths: string[]): PptxSlideMaster[];
}

/**
 * Creates a MasterLayoutResolver for resolving Slide Master and Layout structures.
 * @param zipReader Active `ZipReader` instance.
 * @param parser Optional custom `XmlParser` instance.
 * @returns Frozen `MasterLayoutResolver` instance.
 */
export function createMasterLayoutResolver(
  zipReader: ZipReader,
  parser: XmlParser = defaultXmlParser,
): MasterLayoutResolver {
  return Object.freeze({
    parseLayouts(layoutPaths: string[]): PptxSlideLayout[] {
      const layouts: PptxSlideLayout[] = [];

      for (const layoutPath of layoutPaths) {
        const xml = zipReader.getFileText(layoutPath);
        if (!xml) continue;

        const id = layoutPath.split('/').pop()?.replace('.xml', '') || layoutPath;
        const parsed = parser.parse<Record<string, unknown>>(xml);
        const sldLayout = getXmlChild(parsed, 'sldLayout') || (parsed);
        const cSld = getXmlChild(sldLayout, 'cSld') || {};

        const matchingName = sldLayout['@_matchingName'] ? String(sldLayout['@_matchingName']) : undefined;
        const name = (cSld['@_name'] || matchingName || sldLayout['@_type'] || id) as string;
        const type = (sldLayout['@_type'] as string) || 'custom';
        const preserve = sldLayout['@_preserve'] === '1' || sldLayout['@_preserve'] === true;
        const userDrawn = sldLayout['@_userDrawn'] === '1' || sldLayout['@_userDrawn'] === true;

        const shapes = parseShapes(xml, parser);

        // Parse layout .rels to find master ID
        const relsPath = layoutPath.replace('slideLayouts/', 'slideLayouts/_rels/').concat('.rels');
        const relsXml = zipReader.getFileText(relsPath);
        const relResolver = createRelationshipResolver(relsXml, layoutPath, parser);

        const masterRel = relResolver.getRelationshipsByType('slideMaster')[0];
        const masterId = masterRel ? masterRel.resolvedTarget.split('/').pop()?.replace('.xml', '') || masterRel.resolvedTarget : '';

        layouts.push({
          elements: shapes,
          id,
          masterId,
          matchingName,
          name,
          preserve: preserve || undefined,
          shapes,
          type,
          userDrawn: userDrawn || undefined,
        });
      }

      return layouts;
    },

    parseMasters(masterPaths: string[]): PptxSlideMaster[] {
      const masters: PptxSlideMaster[] = [];

      for (const masterPath of masterPaths) {
        const xml = zipReader.getFileText(masterPath);
        if (!xml) continue;

        const id = masterPath.split('/').pop()?.replace('.xml', '') || masterPath;
        const parsed = parser.parse<Record<string, unknown>>(xml);
        const sldMaster = getXmlChild(parsed, 'sldMaster') || parsed;
        const cSld = getXmlChild(sldMaster, 'cSld') || {};
        const name = (cSld['@_name'] || sldMaster['@_name'] || id) as string;
        const preserve = sldMaster['@_preserve'] === '1' || sldMaster['@_preserve'] === true;

        const shapes = parseShapes(xml, parser);

        // Parse master .rels to find layout IDs
        const relsPath = masterPath.replace('slideMasters/', 'slideMasters/_rels/').concat('.rels');
        const relsXml = zipReader.getFileText(relsPath);
        const relResolver = createRelationshipResolver(relsXml, masterPath, parser);

        const layoutRels = relResolver.getRelationshipsByType('slideLayout');
        const layoutIds = layoutRels.map((r) => r.resolvedTarget.split('/').pop()?.replace('.xml', '') || r.resolvedTarget);

        masters.push({
          elements: shapes,
          id,
          layoutIds,
          name,
          preserve: preserve || undefined,
          shapes,
        });
      }

      return masters;
    },
  });
}
