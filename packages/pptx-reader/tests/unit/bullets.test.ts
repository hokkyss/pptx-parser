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
