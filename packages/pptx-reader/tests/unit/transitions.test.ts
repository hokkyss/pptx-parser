import { describe, expect, it } from 'vitest';
import { parseTransition } from '../../lib/parsers/transition-parser';

describe('Slide Transition Parser', () => {
  it('returns undefined when no transition node exists', () => {
    const xml = '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld/></p:sld>';
    expect(parseTransition(xml)).toBeUndefined();
  });

  it('parses basic fade transition', () => {
    const xml = `
      <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
        <p:transition spd="fast">
          <p:fade/>
        </p:transition>
      </p:sld>
    `;
    const transition = parseTransition(xml);
    expect(transition).toBeDefined();
    expect(transition?.type).toBe('fade');
    expect(transition?.speed).toBe('fast');
    expect(transition?.durationMs).toBe(500);
    expect(transition?.advanceOnClick).toBe(true);
  });

  it('parses fade through black transition', () => {
    const xml = `
      <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
        <p:transition>
          <p:fade thruBlk="1"/>
        </p:transition>
      </p:sld>
    `;
    const transition = parseTransition(xml);
    expect(transition?.type).toBe('fade');
    expect(transition?.throughBlack).toBe(true);
  });

  it('parses directional wipe and push transitions', () => {
    const wipeXml = `
      <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
        <p:transition spd="slow">
          <p:wipe dir="r"/>
        </p:transition>
      </p:sld>
    `;
    const wipeTrans = parseTransition(wipeXml);
    expect(wipeTrans?.type).toBe('wipe');
    expect(wipeTrans?.direction).toBe('right');
    expect(wipeTrans?.speed).toBe('slow');
    expect(wipeTrans?.durationMs).toBe(2000);

    const pushXml = `
      <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
        <p:transition advClick="0" advTm="3500">
          <p:push dir="u"/>
        </p:transition>
      </p:sld>
    `;
    const pushTrans = parseTransition(pushXml);
    expect(pushTrans?.type).toBe('push');
    expect(pushTrans?.direction).toBe('up');
    expect(pushTrans?.advanceOnClick).toBe(false);
    expect(pushTrans?.advanceAfterMs).toBe(3500);
  });

  it('parses wheel transition with spokes', () => {
    const xml = `
      <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
        <p:transition>
          <p:wheel spokes="8"/>
        </p:transition>
      </p:sld>
    `;
    const transition = parseTransition(xml);
    expect(transition?.type).toBe('wheel');
    expect(transition?.spokes).toBe(8);
  });
});

describe('parseTransition object input', () => {
  it('parses from JS object node directly', () => {
    const obj = {
      'p:sld': {
        'p:transition': {
          '@_spd': 'med',
          'p:wipe': { '@_dir': 'r' },
        },
      },
    };
    const t = parseTransition(obj);
    expect(t?.type).toBe('wipe');
    expect(t?.direction).toBe('right');
    expect(t?.speed).toBe('med');
  });

  it('parses transition from object with explicit duration and push direction', () => {
    const trans = parseTransition({
      'p:transition': {
        '@_dur': 500,
        '@_spd': 'fast',
        'p:push': { '@_dir': 'r' },
      },
    });
    expect(trans?.type).toBe('push');
    expect(trans?.direction).toBe('right');
    expect(trans?.speed).toBe('fast');
  });
});
