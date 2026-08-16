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
