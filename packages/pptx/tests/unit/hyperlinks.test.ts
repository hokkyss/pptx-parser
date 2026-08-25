import { describe, expect, it } from 'vitest';
import { inches, points } from '@hokkyss/pptx-core';
import { Presentation } from '../../lib/presentation';

describe('Fluent Hyperlinks (@hokkyss/pptx)', () => {
  it('should create and roundtrip presentation with text, shape, and action hyperlinks', async () => {
    const pres = Presentation.create();

    // Slide 1: Introduction with external links & slide jump
    const slide1 = pres.addSlide();
    slide1.addText('Hokkyss PPTX Parser', {
      fontSize: points(28),
      h: inches(1),
      hyperlink: { tooltip: 'Official Repo', url: 'https://github.com/hokkyss/pptx-parser' },
      w: inches(8),
      x: inches(1),
      y: inches(1),
    });

    slide1.addText([
      { text: 'Click ' },
      {
        bold: true,
        color: '38BDF8',
        hyperlink: { slideIndex: 2, tooltip: 'Jump to Slide 2' },
        text: 'here to jump to Slide 2',
      },
      { text: ' or navigate using buttons below.' },
    ], {
      fontSize: points(16),
      h: inches(1.5),
      w: inches(8),
      x: inches(1),
      y: inches(2.5),
    });

    slide1.addShape('roundRect', {
      fill: '10B981',
      h: inches(0.8),
      hyperlink: { action: 'nextSlide', tooltip: 'Next Slide' },
      text: 'Next Slide ➔',
      textOptions: { color: 'FFFFFF', fontSize: points(14) },
      w: inches(2.5),
      x: inches(1),
      y: inches(4.5),
    });

    // Slide 2: Destination slide with back action
    const slide2 = pres.addSlide();
    slide2.addText('Slide 2: Welcome!', {
      fontSize: points(24),
      h: inches(1),
      w: inches(8),
      x: inches(1),
      y: inches(1),
    });

    slide2.addShape('roundRect', {
      fill: '64748B',
      h: inches(0.8),
      hyperlink: { action: 'firstSlide', tooltip: 'Return Home' },
      text: '⬅ Back to Start',
      textOptions: { color: 'FFFFFF', fontSize: points(14) },
      w: inches(2.5),
      x: inches(1),
      y: inches(3),
    });

    const buffer = await pres.toBuffer();
    expect(buffer).toBeInstanceOf(Uint8Array);
    expect(buffer.length).toBeGreaterThan(0);

    // Read back and verify
    const reloaded = await Presentation.load(buffer);

    expect(reloaded.slides).toHaveLength(2);

    const s1Elements = reloaded.slides[0].getElements();
    expect(s1Elements).toHaveLength(3);

    // Verify Title Hyperlink
    const titleShape = s1Elements[0];
    expect(titleShape.hyperlink).toMatchObject({
      tooltip: 'Official Repo',
      url: 'https://github.com/hokkyss/pptx-parser',
    });

    // Verify Rich Text Run Hyperlink
    const bodyShape = s1Elements[1];
    const runs = bodyShape.textBody?.paragraphs[0].runs || [];
    expect(runs).toHaveLength(3);
    expect(runs[1].properties?.hyperlink).toMatchObject({
      slideIndex: 2,
      tooltip: 'Jump to Slide 2',
    });

    // Verify Shape Action Hyperlink
    const buttonShape = s1Elements[2];
    expect(buttonShape.hyperlink).toMatchObject({
      action: 'nextSlide',
      tooltip: 'Next Slide',
    });

    // Verify Slide 2 Action Button
    const s2Elements = reloaded.slides[1].getElements();
    const backButton = s2Elements[1];
    expect(backButton.hyperlink).toMatchObject({
      action: 'firstSlide',
      tooltip: 'Return Home',
    });
  });

  it('safely neutralizes dangerous protocol injection in fluent SDK', async () => {
    const pres = Presentation.create();
    const slide = pres.addSlide();

    slide.addText('Safe Link', {
      hyperlink: 'https://example.com',
    });

    slide.addText('Malicious Script Link', {
      hyperlink: 'javascript:alert(document.cookie)',
    });

    slide.addShape('roundRect', {
      h: inches(1),
      hyperlink: {
        tooltip: 'Malicious\r\nTooltip\0Breakout',
        url: 'file:///etc/passwd',
      },
      w: inches(2),
      x: inches(1),
      y: inches(1),
    });

    const buffer = await pres.toBuffer();
    const reloaded = await Presentation.load(buffer);

    const elements = reloaded.slides[0].getElements();
    expect(elements).toHaveLength(3);

    // Safe link preserved
    expect(elements[0].hyperlink).toMatchObject({
      url: 'https://example.com',
    });

    // Malicious javascript link stripped
    expect(elements[1].hyperlink).toBeUndefined();

    // Malicious file link stripped, tooltip sanitized
    expect(elements[2].hyperlink).toBeUndefined();
  });
});
