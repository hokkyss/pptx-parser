import { describe, it, expect } from 'vitest';
import { parseParagraph } from '../../lib/parsers/text-parser';

describe('Bullet Point Parsing', () => {
  it('should parse character bullet points (<a:buChar>)', () => {
    const pNode = {
      'a:pPr': {
        '@_lvl': '0',
        'a:buChar': { '@_char': '•' },
      },
      'a:r': {
        'a:t': 'Bullet item 1',
      },
    };

    const paragraph = parseParagraph(pNode);
    expect(paragraph.properties.bullet).toBeDefined();
    expect(paragraph.properties.bullet?.type).toBe('char');
    expect(paragraph.properties.bullet?.char).toBe('•');
  });

  it('should parse auto-numbering bullets (<a:buAutoNum>)', () => {
    const pNode = {
      'a:pPr': {
        '@_lvl': '1',
        'a:buAutoNum': { '@_type': 'arabicPeriod', '@_startAt': '1' },
      },
      'a:r': {
        'a:t': 'Numbered item 1',
      },
    };

    const paragraph = parseParagraph(pNode);
    expect(paragraph.properties.bullet).toBeDefined();
    expect(paragraph.properties.bullet?.type).toBe('autoNum');
    expect(paragraph.properties.bullet?.autoNumType).toBe('arabicPeriod');
    expect(paragraph.properties.bullet?.startAt).toBe(1);
  });

  it('should handle buNone for explicit no-bullet paragraphs', () => {
    const pNode = {
      'a:pPr': {
        'a:buNone': {},
      },
      'a:r': {
        'a:t': 'Plain text without bullet',
      },
    };

    const paragraph = parseParagraph(pNode);
    expect(paragraph.properties.bullet).toBeDefined();
    expect(paragraph.properties.bullet?.type).toBe('none');
  });
});

import { parseTextBodyXml, parseRunProperties } from '../../lib/parsers/text-parser';

describe('parseTextBodyXml and run color parsing', () => {
  it('parses text body from XML string', () => {
    const xml = '<a:txBody><a:p><a:r><a:t>Hello XML</a:t></a:r></a:p></a:txBody>';
    const body = parseTextBodyXml(xml);
    expect(body.paragraphs[0].runs[0].text).toBe('Hello XML');
  });

  it('parses solidFill srgbClr color in run properties', () => {
    const rPr = {
      'a:solidFill': {
        'a:srgbClr': { '@_val': 'FF0000' },
      },
    };
    const parsed = parseRunProperties(rPr);
    expect(parsed.color).toBe('FF0000');
  });
});

describe('parseParagraph fields and number text nodes', () => {
  it('parses a:fld fields and object #text nodes', () => {
    const pNodeWithField = {
      'a:fld': [
        { 'a:t': 42 },
        { 'a:t': { '#text': 'Preserved Text' } },
      ],
    };
    const parsed = parseParagraph(pNodeWithField);
    expect(parsed.runs).toHaveLength(2);
    expect(parsed.runs[0].text).toBe('42');
    expect(parsed.runs[1].text).toBe('Preserved Text');
  });

  it('covers rich text formatting runs with strikethrough, underline, and size', () => {
    const pNode = {
      'a:pPr': {
        '@_algn': 'ctr',
      },
      'a:r': [
        {
          'a:rPr': { '@_b': '1', '@_i': '1', '@_strike': 'sngStrike', '@_sz': '2400', '@_u': 'sng' },
          'a:t': 'Formatted Text',
        },
      ],
    };
    const parsed = parseParagraph(pNode);
    expect(parsed.runs[0].text).toBe('Formatted Text');
    expect(parsed.runs[0].properties?.bold).toBe(true);
    expect(parsed.runs[0].properties?.italic).toBe(true);
    expect(parsed.runs[0].properties?.fontSize).toBe(2400);
    expect(parsed.runs[0].properties?.underline).toBe(true);
    expect(parsed.runs[0].properties?.strikethrough).toBe(true);
  });

  it('parses <a:br> line break elements with optional run properties', () => {
    const pNode = {
      'a:pPr': {
        '@_lvl': '1',
      },
      'a:r': [
        { 'a:t': 'First line' },
      ],
      'a:br': [
        {
          'a:rPr': { '@_b': '1' },
        },
      ],
    };
    const parsed = parseParagraph(pNode);
    expect(parsed.properties.level).toBe(1);
    expect(parsed.runs).toHaveLength(2);
    expect(parsed.runs[0].text).toBe('First line');
    expect(parsed.runs[1].break).toBe(true);
    expect(parsed.runs[1].properties?.bold).toBe(true);
  });
});

