import { describe, expect, it } from 'vitest';
import type { ZipReader } from '@hokkyss/pptx-core';
import { createMasterLayoutResolver } from '../../lib/resolvers/master-layout-resolver';

// ── Helpers ───────────────────────────────────────────────────────────────────

/**
 *
 */
function mockZipReader(files: Record<string, null | string>): ZipReader {
  return {
    getFileAsString: (path: string) => Promise.resolve(files[path] ?? ''),
    getFileData: () => undefined,
    getFileText: (path: string) => files[path] ?? undefined,
    getPaths: () => Object.keys(files),
    getPathsStartingWith: (prefix: string) => Object.keys(files).filter((k) => k.startsWith(prefix)),
    hasFile: (path: string) => path in files,
    listFiles: () => Object.keys(files),
  };
}

const layoutXml = (extraAttrs = '', cSldName = 'Blank') => `<?xml version="1.0"?>
<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             type="blank" ${extraAttrs}>
  <p:cSld name="${cSldName}">
    <p:spTree/>
  </p:cSld>
</p:sldLayout>`;

const layoutRels = (masterTarget = '../slideMasters/slideMaster1.xml') => `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
    Target="${masterTarget}"/>
</Relationships>`;

const masterXml = (extraAttrs = '', cSldName = 'Office Theme') => `<?xml version="1.0"?>
<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ${extraAttrs}>
  <p:cSld name="${cSldName}">
    <p:spTree/>
  </p:cSld>
</p:sldMaster>`;

const masterRels = (layoutTargets: string[] = []) => {
  const rels = layoutTargets
    .map((t, i) => `  <Relationship Id="rId${i + 1}"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
    Target="${t}"/>`)
    .join('\n');
  return `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${rels}
</Relationships>`;
};

// ── parseLayouts ──────────────────────────────────────────────────────────────

describe('createMasterLayoutResolver – parseLayouts', () => {
  it('returns empty array for empty path list', () => {
    const resolver = createMasterLayoutResolver(mockZipReader({}));
    expect(resolver.parseLayouts([])).toEqual([]);
  });

  it('skips layout file when XML is not available in zip', () => {
    const zip = mockZipReader({ 'ppt/slideLayouts/slideLayout1.xml': null });
    const result = createMasterLayoutResolver(zip).parseLayouts(['ppt/slideLayouts/slideLayout1.xml']);
    expect(result).toHaveLength(0);
  });

  it('parses id, name, and type from a valid layout file', () => {
    const zip = mockZipReader({
      'ppt/slideLayouts/slideLayout1.xml': layoutXml('', 'Title Slide'),
      'ppt/slideLayouts/_rels/slideLayout1.xml.rels': layoutRels(),
    });
    const [layout] = createMasterLayoutResolver(zip).parseLayouts(['ppt/slideLayouts/slideLayout1.xml']);
    expect(layout.id).toBe('slideLayout1');
    expect(layout.name).toBe('Title Slide');
    expect(layout.type).toBe('blank');
  });

  it('resolves masterId from the layout rels file', () => {
    const zip = mockZipReader({
      'ppt/slideLayouts/slideLayout1.xml': layoutXml(),
      'ppt/slideLayouts/_rels/slideLayout1.xml.rels': layoutRels('../slideMasters/slideMaster1.xml'),
    });
    const [layout] = createMasterLayoutResolver(zip).parseLayouts(['ppt/slideLayouts/slideLayout1.xml']);
    expect(layout.masterId).toBe('slideMaster1');
  });

  it('sets preserve=true when @_preserve="1"', () => {
    const zip = mockZipReader({
      'ppt/slideLayouts/slideLayout1.xml': layoutXml('preserve="1"'),
      'ppt/slideLayouts/_rels/slideLayout1.xml.rels': layoutRels(),
    });
    const [layout] = createMasterLayoutResolver(zip).parseLayouts(['ppt/slideLayouts/slideLayout1.xml']);
    expect(layout.preserve).toBe(true);
  });

  it('sets preserve=undefined when @_preserve is not "1"', () => {
    const zip = mockZipReader({
      'ppt/slideLayouts/slideLayout1.xml': layoutXml('preserve="0"'),
      'ppt/slideLayouts/_rels/slideLayout1.xml.rels': layoutRels(),
    });
    const [layout] = createMasterLayoutResolver(zip).parseLayouts(['ppt/slideLayouts/slideLayout1.xml']);
    expect(layout.preserve).toBeUndefined();
  });

  it('sets userDrawn=true when @_userDrawn="1"', () => {
    const zip = mockZipReader({
      'ppt/slideLayouts/slideLayout1.xml': layoutXml('userDrawn="1"'),
      'ppt/slideLayouts/_rels/slideLayout1.xml.rels': layoutRels(),
    });
    const [layout] = createMasterLayoutResolver(zip).parseLayouts(['ppt/slideLayouts/slideLayout1.xml']);
    expect(layout.userDrawn).toBe(true);
  });

  it('sets matchingName from @_matchingName attribute', () => {
    const zip = mockZipReader({
      'ppt/slideLayouts/slideLayout1.xml': layoutXml('matchingName="Custom Name"'),
      'ppt/slideLayouts/_rels/slideLayout1.xml.rels': layoutRels(),
    });
    const [layout] = createMasterLayoutResolver(zip).parseLayouts(['ppt/slideLayouts/slideLayout1.xml']);
    expect(layout.matchingName).toBe('Custom Name');
  });

  it('falls back to type as name when cSld @_name is empty and matchingName absent', () => {
    const xml = `<?xml version="1.0"?>
<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="title">
  <p:cSld><p:spTree/></p:cSld>
</p:sldLayout>`;
    const zip = mockZipReader({
      'ppt/slideLayouts/slideLayout1.xml': xml,
      'ppt/slideLayouts/_rels/slideLayout1.xml.rels': layoutRels(),
    });
    const [layout] = createMasterLayoutResolver(zip).parseLayouts(['ppt/slideLayouts/slideLayout1.xml']);
    expect(layout.name).toBe('title');
  });

  it('falls back to "custom" type when @_type attribute is absent', () => {
    const xml = `<?xml version="1.0"?>
<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld name="No Type"><p:spTree/></p:cSld>
</p:sldLayout>`;
    const zip = mockZipReader({
      'ppt/slideLayouts/slideLayout1.xml': xml,
      'ppt/slideLayouts/_rels/slideLayout1.xml.rels': layoutRels(),
    });
    const [layout] = createMasterLayoutResolver(zip).parseLayouts(['ppt/slideLayouts/slideLayout1.xml']);
    expect(layout.type).toBe('custom');
  });

  it('parses multiple layout files and returns them all', () => {
    const zip = mockZipReader({
      'ppt/slideLayouts/slideLayout1.xml': layoutXml('', 'Layout 1'),
      'ppt/slideLayouts/_rels/slideLayout1.xml.rels': layoutRels(),
      'ppt/slideLayouts/slideLayout2.xml': layoutXml('', 'Layout 2'),
      'ppt/slideLayouts/_rels/slideLayout2.xml.rels': layoutRels(),
    });
    const layouts = createMasterLayoutResolver(zip).parseLayouts([
      'ppt/slideLayouts/slideLayout1.xml',
      'ppt/slideLayouts/slideLayout2.xml',
    ]);
    expect(layouts).toHaveLength(2);
    expect(layouts[0].name).toBe('Layout 1');
    expect(layouts[1].name).toBe('Layout 2');
  });
});

// ── parseMasters ──────────────────────────────────────────────────────────────

describe('createMasterLayoutResolver – parseMasters', () => {
  it('returns empty array for empty path list', () => {
    expect(createMasterLayoutResolver(mockZipReader({})).parseMasters([])).toEqual([]);
  });

  it('skips master file when XML is not available', () => {
    const zip = mockZipReader({ 'ppt/slideMasters/slideMaster1.xml': null });
    const result = createMasterLayoutResolver(zip).parseMasters(['ppt/slideMasters/slideMaster1.xml']);
    expect(result).toHaveLength(0);
  });

  it('parses id and name from cSld of a valid master file', () => {
    const zip = mockZipReader({
      'ppt/slideMasters/slideMaster1.xml': masterXml('', 'Office Theme'),
      'ppt/slideMasters/_rels/slideMaster1.xml.rels': masterRels(),
    });
    const [master] = createMasterLayoutResolver(zip).parseMasters(['ppt/slideMasters/slideMaster1.xml']);
    expect(master.id).toBe('slideMaster1');
    expect(master.name).toBe('Office Theme');
  });

  it('sets preserve=true when @_preserve="1"', () => {
    const zip = mockZipReader({
      'ppt/slideMasters/slideMaster1.xml': masterXml('preserve="1"'),
      'ppt/slideMasters/_rels/slideMaster1.xml.rels': masterRels(),
    });
    const [master] = createMasterLayoutResolver(zip).parseMasters(['ppt/slideMasters/slideMaster1.xml']);
    expect(master.preserve).toBe(true);
  });

  it('sets preserve=undefined when @_preserve attribute is absent', () => {
    const zip = mockZipReader({
      'ppt/slideMasters/slideMaster1.xml': masterXml(),
      'ppt/slideMasters/_rels/slideMaster1.xml.rels': masterRels(),
    });
    const [master] = createMasterLayoutResolver(zip).parseMasters(['ppt/slideMasters/slideMaster1.xml']);
    expect(master.preserve).toBeUndefined();
  });

  it('extracts layoutIds from slideLayout rels', () => {
    const zip = mockZipReader({
      'ppt/slideMasters/slideMaster1.xml': masterXml(),
      'ppt/slideMasters/_rels/slideMaster1.xml.rels': masterRels([
        '../slideLayouts/slideLayout1.xml',
        '../slideLayouts/slideLayout2.xml',
      ]),
    });
    const [master] = createMasterLayoutResolver(zip).parseMasters(['ppt/slideMasters/slideMaster1.xml']);
    expect(master.layoutIds).toEqual(['slideLayout1', 'slideLayout2']);
  });

  it('returns empty layoutIds when rels has no slideLayout relationships', () => {
    const zip = mockZipReader({
      'ppt/slideMasters/slideMaster1.xml': masterXml(),
      'ppt/slideMasters/_rels/slideMaster1.xml.rels': `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`,
    });
    const [master] = createMasterLayoutResolver(zip).parseMasters(['ppt/slideMasters/slideMaster1.xml']);
    expect(master.layoutIds).toEqual([]);
  });
});
