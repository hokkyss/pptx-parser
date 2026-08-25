import { describe, it, expect } from 'vitest';
import { parseTextBody } from '../../lib/parsers/text-parser.js';
import { parseTable } from '../../lib/parsers/table-parser.js';
import { parseSingleShape } from '../../lib/parsers/shape-parser.js';

describe('Clean CamelCase Properties Parsing', () => {
  it('should parse bodyProperties into clean camelCase fields without @_ attributes', () => {
    const txBodyXml = {
      'a:bodyPr': {
        '@_anchor': 'ctr',
        '@_wrap': 'square',
        '@_lIns': '91440',
        '@_tIns': '45720',
        '@_rIns': '91440',
        '@_bIns': '45720',
        '@_numCol': '2',
        '@_spcCol': '182880',
        '@_vert': 'horz',
      },
      'a:p': {
        'a:r': { 'a:t': 'Sample text' },
      },
    };

    const textBody = parseTextBody(txBodyXml);
    expect(textBody.bodyProperties).toBeDefined();
    expect(textBody.bodyProperties.verticalAlignment).toBe('middle');
    expect(textBody.bodyProperties.wrap).toBe('square');
    expect(textBody.bodyProperties.leftInset).toBe(91440);
    expect(textBody.bodyProperties.topInset).toBe(45720);
    expect(textBody.bodyProperties.rightInset).toBe(91440);
    expect(textBody.bodyProperties.bottomInset).toBe(45720);
    expect(textBody.bodyProperties.columns).toBe(2);
    expect(textBody.bodyProperties.columnSpacing).toBe(182880);

    // Verify raw @_ attributes do not leak
    expect('@_wrap' in textBody.bodyProperties).toBe(false);
    expect('@_anchor' in textBody.bodyProperties).toBe(false);
  });

  it('should parse paragraph margin, indentation, and spacing into camelCase fields', () => {
    const pNode = {
      'a:pPr': {
        '@_algn': 'l',
        '@_lvl': '1',
        '@_marL': '342900',
        '@_marR': '171450',
        '@_indent': '-342900',
        'a:spcBef': {
          'a:spcPts': { '@_val': '1200' }, // 12pt = 1200 HundredthsPoint
        },
        'a:spcAft': {
          'a:spcPts': { '@_val': '600' }, // 6pt = 600 HundredthsPoint
        },
        'a:lnSpc': {
          'a:spcPct': { '@_val': '120000' }, // 120% line spacing
        },
      },
      'a:r': {
        'a:t': 'Indented text',
      },
    };

    const textBody = parseTextBody({ 'a:p': pNode });
    const pProps = textBody.paragraphs[0].properties;

    expect(pProps.alignment).toBe('left');
    expect(pProps.level).toBe(1);
    expect(pProps.leftMargin).toBe(342900);
    expect(pProps.rightMargin).toBe(171450);
    expect(pProps.firstLineIndent).toBe(-342900);
    expect(pProps.spaceBefore).toBe(1200);
    expect(pProps.spaceAfter).toBe(600);
    expect(pProps.lineSpacing).toBe(120000);
  });

  it('should parse table cell insets and verticalAlignment into camelCase fields', () => {
    const tblXml = {
      'a:tbl': {
        'a:tblGrid': {
          'a:gridCol': { '@_w': '1000000' },
        },
        'a:tr': {
          '@_h': '500000',
          'a:tc': {
            '@_marL': '91440',
            '@_marT': '45720',
            '@_anchor': 'b',
            'a:txBody': {
              'a:p': {
                'a:r': { 'a:t': 'Cell text' },
              },
            },
          },
        },
      },
    };

    const table = parseTable(tblXml);
    expect(table).toBeDefined();
    const cellProps = table!.rows[0].cells[0].properties;

    expect(cellProps).toBeDefined();
    expect(cellProps?.leftInset).toBe(91440);
    expect(cellProps?.topInset).toBe(45720);
    expect(cellProps?.verticalAlignment).toBe('bottom');

    // Verify raw @_ attributes do not leak
    if (cellProps) {
      expect('@_marL' in cellProps).toBe(false);
      expect('@_anchor' in cellProps).toBe(false);
    }
  });

  it('should parse shape lock flags and visibility state correctly', () => {
    const picNode = {
      'p:nvPicPr': {
        'p:cNvPr': {
          '@_id': '35',
          '@_name': '図 presentation-cover',
          '@_hidden': '1',
        },
        'p:cNvPicPr': {
          'a:picLocks': {
            '@_noGrp': '1',
            '@_noMove': '1',
            '@_noResize': '1',
          },
        },
      },
    };

    const shape = parseSingleShape(picNode, 'picture');
    expect(shape.isVisible).toBe(false);
    expect(shape.isLocked).toBe(true);
    expect(shape.locks).toEqual({
      noGrp: true,
      noMove: true,
      noResize: true,
    });
  });
});
