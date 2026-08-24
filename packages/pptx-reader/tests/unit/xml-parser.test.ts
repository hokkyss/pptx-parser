import { describe, expect, it } from 'vitest';
import { createXmlParser, defaultXmlParser, getXmlChild, getXmlChildren } from '../../lib/xml/xml-parser';

describe('XML Parser Helper Functions', () => {
  it('handles empty XML and returns empty object', () => {
    const parser = createXmlParser();
    expect(parser.parse('')).toEqual({});
    expect(parser.parse('   ')).toEqual({});
  });

  it('getXmlChild handles array and non-array children and namespace prefixes', () => {
    expect(getXmlChild(undefined, 'test')).toBeUndefined();
    expect(getXmlChild({}, 'test')).toBeUndefined();

    const nodeWithArray = {
      'p:sp': [{ id: 1 }, { id: 2 }],
    };
    expect(getXmlChild(nodeWithArray, 'sp')).toEqual({ id: 1 });

    const nodeWithObj = {
      'a:xfrm': { x: 0 },
    };
    expect(getXmlChild(nodeWithObj, 'xfrm')).toEqual({ x: 0 });
  });

  it('getXmlChildren handles arrays, objects, and empty/undefined nodes', () => {
    expect(getXmlChildren(undefined, 'test')).toEqual([]);
    expect(getXmlChildren({}, 'test')).toEqual([]);

    const node = {
      'p:sp': [{ id: 1 }, { id: 2 }],
      'a:single': { key: 'val' },
    };

    const spChildren = getXmlChildren(node, 'sp');
    expect(spChildren).toHaveLength(2);

    const singleChildren = getXmlChildren(node, 'single');
    expect(singleChildren).toHaveLength(1);
    expect(singleChildren[0]).toEqual({ key: 'val' });
  });
});
