import { describe, it, expect } from 'vitest';
import { parseAnimations } from '../../lib/parsers/animation-parser';

describe('Animation Parser', () => {
  it('should return empty array if no timing tree', () => {
    const xml = `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"></p:sld>`;
    const animations = parseAnimations(xml);
    expect(animations).toEqual([]);
  });

  it('should parse animEffect and targetShapeId', () => {
    const xml = `
      <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
        <p:timing>
          <p:tnLst>
            <p:par>
              <p:cTn id="1">
                <p:childTnLst>
                  <p:animEffect filter="fade" presetClass="appear" nodeType="onClick">
                    <p:tgtEl>
                      <p:spTgt spid="2"/>
                    </p:tgtEl>
                  </p:animEffect>
                </p:childTnLst>
              </p:cTn>
            </p:par>
          </p:tnLst>
        </p:timing>
      </p:sld>
    `;
    const animations = parseAnimations(xml);
    expect(animations.length).toBe(1);
    expect(animations[0].targetShapeId).toBe('2');
    expect(animations[0].effect).toBe('appear');
  });
});

describe('parseAnimations object and array recursion', () => {
  it('parses from JS object with array timing nodes', () => {
    const obj = {
      'p:sld': {
        'p:timing': {
          'p:tnLst': [
            {
              'p:tgtEl': { 'p:spTgt': { '@_spid': '5' } },
              '@_presetClass': 'fly',
              '@_nodeType': 'afterPrevious',
            },
          ],
        },
      },
    };
    const animations = parseAnimations(obj);
    expect(animations).toHaveLength(1);
    expect(animations[0].targetShapeId).toBe('5');
    expect(animations[0].effect).toBe('fly');
    expect(animations[0].trigger).toBe('afterPrevious');
  });
});
