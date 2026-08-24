import { describe, expect, it } from 'vitest';
import { parseChart } from '../../lib/parsers/chart-parser';

describe('parseChart', () => {
  // ── Guard cases ─────────────────────────────────────────────────────────────

  it('returns undefined when there is no c:chart node', () => {
    const result = parseChart({ 'c:chartSpace': {} });
    expect(result).toBeUndefined();
  });

  it('returns undefined when there is no c:plotArea node', () => {
    const result = parseChart({
      'c:chartSpace': { 'c:chart': {} },
    });
    expect(result).toBeUndefined();
  });

  it('falls back to "unknownChart" when no key ends with "Chart"', () => {
    const result = parseChart({
      'c:chartSpace': {
        'c:chart': { 'c:plotArea': { 'c:catAx': {} } },
      },
    });
    expect(result).toBeDefined();
    expect(result!.chartType).toBe('unknownChart');
    expect(result!.series).toHaveLength(0);
    expect(result!.categories).toHaveLength(0);
  });

  // ── Basic chart types ────────────────────────────────────────────────────────

  it('parses a bar chart object with a single series', () => {
    const chartObj = {
      'c:chartSpace': {
        'c:chart': {
          'c:plotArea': {
            'c:barChart': {
              'c:ser': {
                'c:idx': { '@_val': '0' },
                'c:order': { '@_val': '0' },
                'c:tx': { 'c:v': 'Revenue' },
                'c:val': {
                  'c:numLit': {
                    'c:pt': [{ 'c:v': '100' }, { 'c:v': '200' }],
                  },
                },
                'c:cat': {
                  'c:strLit': {
                    'c:pt': [{ 'c:v': 'Q1' }, { 'c:v': 'Q2' }],
                  },
                },
              },
            },
          },
        },
      },
    };

    const result = parseChart(chartObj);
    expect(result).toBeDefined();
    expect(result!.chartType).toBe('barChart');
    expect(result!.series).toHaveLength(1);
    expect(result!.series[0].name).toBe('Revenue');
    expect(result!.series[0].index).toBe(0);
    expect(result!.series[0].order).toBe(0);
    expect(result!.series[0].values).toEqual([100, 200]);
    expect(result!.categories).toEqual(['Q1', 'Q2']);
  });

  it('parses a line chart with multiple series', () => {
    const chartObj = {
      'c:chartSpace': {
        'c:chart': {
          'c:plotArea': {
            'c:lineChart': {
              'c:ser': [
                {
                  'c:idx': { '@_val': '0' },
                  'c:order': { '@_val': '0' },
                  'c:tx': { 'c:v': 'Series A' },
                  'c:val': { 'c:numLit': { 'c:pt': [{ 'c:v': '10' }, { 'c:v': '20' }] } },
                  'c:cat': { 'c:strLit': { 'c:pt': [{ 'c:v': 'Jan' }, { 'c:v': 'Feb' }] } },
                },
                {
                  'c:idx': { '@_val': '1' },
                  'c:order': { '@_val': '1' },
                  'c:tx': { 'c:v': 'Series B' },
                  'c:val': { 'c:numLit': { 'c:pt': [{ 'c:v': '30' }, { 'c:v': '40' }] } },
                },
              ],
            },
          },
        },
      },
    };

    const result = parseChart(chartObj);
    expect(result!.chartType).toBe('lineChart');
    expect(result!.series).toHaveLength(2);
    expect(result!.series[1].name).toBe('Series B');
    expect(result!.series[1].index).toBe(1);
    // categories taken from the series with more categories (series[0])
    expect(result!.categories).toEqual(['Jan', 'Feb']);
  });

  // ── XML string input ─────────────────────────────────────────────────────────

  it('accepts a raw XML string and parses it', () => {
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:pieChart>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:val>
            <c:numLit>
              <c:pt><c:v>30</c:v></c:pt>
              <c:pt><c:v>70</c:v></c:pt>
            </c:numLit>
          </c:val>
        </c:ser>
      </c:pieChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;

    const result = parseChart(xml);
    expect(result).toBeDefined();
    expect(result!.chartType).toBe('pieChart');
    expect(result!.series).toHaveLength(1);
    expect(result!.series[0].values).toEqual([30, 70]);
  });

  // ── Series name fallbacks ────────────────────────────────────────────────────

  it('uses default name "Series N" when tx node is absent', () => {
    const result = parseChart({
      'c:chartSpace': {
        'c:chart': {
          'c:plotArea': {
            'c:barChart': { 'c:ser': { 'c:idx': { '@_val': '2' } } },
          },
        },
      },
    });
    expect(result!.series[0].name).toBe('Series 3');
  });

  it('uses default name when tx has no v child', () => {
    const result = parseChart({
      'c:chartSpace': {
        'c:chart': {
          'c:plotArea': {
            'c:barChart': { 'c:ser': { 'c:idx': { '@_val': '0' }, 'c:tx': {} } },
          },
        },
      },
    });
    expect(result!.series[0].name).toBe('Series 1');
  });

  it('extracts series name from #text object inside tx v node', () => {
    const result = parseChart({
      'c:chartSpace': {
        'c:chart': {
          'c:plotArea': {
            'c:barChart': {
              'c:ser': {
                'c:idx': { '@_val': '0' },
                'c:tx': { 'c:v': { '#text': 'TextNode Name' } },
              },
            },
          },
        },
      },
    });
    expect(result!.series[0].name).toBe('TextNode Name');
  });

  it('uses non-prefixed "ser" key as fallback', () => {
    const result = parseChart({
      'c:chartSpace': {
        'c:chart': {
          'c:plotArea': {
            'c:barChart': {
              ser: { idx: { '@_val': '0' }, tx: { v: 'NoPrefixName' } },
            },
          },
        },
      },
    });
    expect(result!.series[0].name).toBe('NoPrefixName');
  });

  it('defaults idx and order to 0 when not present', () => {
    const result = parseChart({
      'c:chartSpace': {
        'c:chart': {
          'c:plotArea': { 'c:barChart': { 'c:ser': {} } },
        },
      },
    });
    expect(result!.series[0].index).toBe(0);
    expect(result!.series[0].order).toBe(0);
  });

  // ── numLit / numRef value extraction ─────────────────────────────────────────

  it('extracts values using numRef (not numLit)', () => {
    const result = parseChart({
      'c:chartSpace': {
        'c:chart': {
          'c:plotArea': {
            'c:barChart': {
              'c:ser': {
                'c:val': { 'c:numRef': { 'c:pt': [{ 'c:v': '42' }] } },
              },
            },
          },
        },
      },
    });
    expect(result!.series[0].values).toEqual([42]);
  });

  it('handles NaN numeric value by substituting 0', () => {
    const result = parseChart({
      'c:chartSpace': {
        'c:chart': {
          'c:plotArea': {
            'c:barChart': {
              'c:ser': {
                'c:val': {
                  'c:numLit': { 'c:pt': [{ 'c:v': 'notanumber' }, { 'c:v': '5' }] },
                },
              },
            },
          },
        },
      },
    });
    expect(result!.series[0].values).toEqual([0, 5]);
  });

  it('handles #text object in numeric value points', () => {
    const result = parseChart({
      'c:chartSpace': {
        'c:chart': {
          'c:plotArea': {
            'c:barChart': {
              'c:ser': {
                'c:val': {
                  'c:numLit': { 'c:pt': { 'c:v': { '#text': '99' } } },
                },
              },
            },
          },
        },
      },
    });
    expect(result!.series[0].values).toEqual([99]);
  });

  it('returns empty values when numLit has no pt key', () => {
    const result = parseChart({
      'c:chartSpace': {
        'c:chart': {
          'c:plotArea': {
            'c:barChart': {
              'c:ser': { 'c:val': { 'c:numLit': {} } },
            },
          },
        },
      },
    });
    expect(result!.series[0].values).toEqual([]);
  });

  it('returns empty values when val has no numLit/numRef', () => {
    const result = parseChart({
      'c:chartSpace': {
        'c:chart': {
          'c:plotArea': {
            'c:barChart': { 'c:ser': { 'c:val': {} } },
          },
        },
      },
    });
    expect(result!.series[0].values).toEqual([]);
  });

  // ── strLit / strRef category extraction ──────────────────────────────────────

  it('extracts categories using strRef (not strLit)', () => {
    const result = parseChart({
      'c:chartSpace': {
        'c:chart': {
          'c:plotArea': {
            'c:barChart': {
              'c:ser': {
                'c:cat': {
                  'c:strRef': { 'c:pt': [{ 'c:v': 'Alpha' }, { 'c:v': 'Beta' }] },
                },
              },
            },
          },
        },
      },
    });
    expect(result!.categories).toEqual(['Alpha', 'Beta']);
  });

  it('handles #text object in category string points', () => {
    const result = parseChart({
      'c:chartSpace': {
        'c:chart': {
          'c:plotArea': {
            'c:barChart': {
              'c:ser': {
                'c:cat': {
                  'c:strLit': { 'c:pt': { 'c:v': { '#text': 'TextCat' } } },
                },
              },
            },
          },
        },
      },
    });
    expect(result!.categories).toEqual(['TextCat']);
  });

  it('returns empty categories when strLit has no pt', () => {
    const result = parseChart({
      'c:chartSpace': {
        'c:chart': {
          'c:plotArea': {
            'c:barChart': {
              'c:ser': { 'c:cat': { 'c:strLit': {} } },
            },
          },
        },
      },
    });
    expect(result!.categories).toEqual([]);
  });

  it('returns empty categories when cat node has no strLit/strRef', () => {
    const result = parseChart({
      'c:chartSpace': {
        'c:chart': {
          'c:plotArea': {
            'c:barChart': { 'c:ser': { 'c:cat': {} } },
          },
        },
      },
    });
    expect(result!.categories).toEqual([]);
  });

  // ── Legend extraction ─────────────────────────────────────────────────────────

  it('parses legend with bottom position (b) and no overlay', () => {
    const result = parseChart({
      'c:chartSpace': {
        'c:chart': {
          'c:plotArea': { 'c:barChart': {} },
          'c:legend': { 'c:legendPos': { '@_val': 'b' } },
        },
      },
    });
    expect(result!.legend).toBeDefined();
    expect(result!.legend!.position).toBe('bottom');
    expect(result!.legend!.overlay).toBeFalsy();
  });

  it('maps all legend position codes correctly', () => {
    const posMap: Record<string, string> = {
      b: 'bottom', l: 'left', r: 'right', t: 'top', tr: 'topRight',
    };
    for (const [code, expected] of Object.entries(posMap)) {
      const result = parseChart({
        'c:chartSpace': {
          'c:chart': {
            'c:plotArea': { 'c:barChart': {} },
            'c:legend': { 'c:legendPos': { '@_val': code } },
          },
        },
      });
      expect(result!.legend!.position).toBe(expected);
    }
  });

  it('defaults to "bottom" for unknown legend position code', () => {
    const result = parseChart({
      'c:chartSpace': {
        'c:chart': {
          'c:plotArea': { 'c:barChart': {} },
          'c:legend': { 'c:legendPos': { '@_val': 'unknown' } },
        },
      },
    });
    expect(result!.legend!.position).toBe('bottom');
  });

  it('defaults legend position to "bottom" when legendPos node is absent', () => {
    const result = parseChart({
      'c:chartSpace': {
        'c:chart': {
          'c:plotArea': { 'c:barChart': {} },
          'c:legend': {},
        },
      },
    });
    expect(result!.legend!.position).toBe('bottom');
  });

  it('sets overlay to true when c:overlay @_val is "1"', () => {
    const result = parseChart({
      'c:chartSpace': {
        'c:chart': {
          'c:plotArea': { 'c:barChart': {} },
          'c:legend': {
            'c:legendPos': { '@_val': 'r' },
            'c:overlay': { '@_val': '1' },
          },
        },
      },
    });
    expect(result!.legend!.overlay).toBe(true);
  });

  it('returns no legend when legend node is absent', () => {
    const result = parseChart({
      'c:chartSpace': {
        'c:chart': { 'c:plotArea': { 'c:barChart': {} } },
      },
    });
    expect(result!.legend).toBeUndefined();
  });

  // ── Fallback: direct pass-through (no c:chartSpace wrapper) ──────────────────

  it('works when top-level node has no c:chartSpace wrapper', () => {
    const result = parseChart({
      'c:chart': {
        'c:plotArea': {
          'c:barChart': { 'c:ser': { 'c:tx': { 'c:v': 'Direct' } } },
        },
      },
    });
    expect(result).toBeDefined();
    expect(result!.series[0].name).toBe('Direct');
  });
});
