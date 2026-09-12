import { describe, expect, it } from 'vitest';
import { createXmlParser, getXmlChild, getXmlChildren } from '../../lib/xml/xml-parser';

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

  it('parses XML documents into ordered XmlElement tree with attributes and text', () => {
    const parser = createXmlParser();
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
    <root attr1="val1" attr2="val2 &amp; more">
      <!-- comment -->
      <child id="1">First &lt;child&gt;</child>
      <child id="2"><![CDATA[Special <cdata> & chars]]></child>
      <standalone flag="true"/>
    </root>`;

    const parsed = parser.parse<{ root: any }>(xml);
    expect(parsed.root).toBeDefined();
    expect(parsed.root.attrs.attr1).toBe('val1');
    expect(parsed.root.attrs.attr2).toBe('val2 & more');
    expect(parsed.root['@_attr1']).toBe('val1');

    // Children order preserved
    const children = parsed.root.children;
    expect(children).toHaveLength(3);
    expect(children[0].tag).toBe('child');
    expect(children[0].attrs.id).toBe('1');
    expect(children[0]['#text']).toBe('First <child>');

    expect(children[1].tag).toBe('child');
    expect(children[1].attrs.id).toBe('2');
    expect(children[1]['#text']).toBe('Special <cdata> & chars');

    expect(children[2].tag).toBe('standalone');
    expect(children[2].attrs.flag).toBe('true');
  });
});
