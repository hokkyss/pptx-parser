import { describe, expect, it } from 'vitest';
import { Presentation } from '../../lib/presentation';

describe('Fluent Hyperlinks (@hokkyss/pptx)', () => {
  it('should create and roundtrip presentation with text, shape, and action hyperlinks', async () => {
    const pres = Presentation.create();

    // Slide 1: Introduction with external links & slide jump
    const slide1 = pres.addSlide();
    slide1.addText('Hokkyss PPTX Parser', {
      fontSize: 28,
      h: 1,
      hyperlink: { tooltip: 'Official Repo', url: 'https://github.com/hokkyss/pptx-parser' },
      w: 8,
      x: 1,
      y: 1,
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
      fontSize: 16,
      h: 1.5,
      w: 8,
      x: 1,
      y: 2.5,
    });

    slide1.addShape('roundRect', {
      fill: '10B981',
      h: 0.8,
      hyperlink: { action: 'nextSlide', tooltip: 'Next Slide' },
      text: 'Next Slide ➔',
      textOptions: { color: 'FFFFFF', fontSize: 14 },
      w: 2.5,
      x: 1,
      y: 4.5,
    });

    // Slide 2: Destination slide with back action
    const slide2 = pres.addSlide();
    slide2.addText('Slide 2: Welcome!', {
      fontSize: 24,
      h: 1,
      w: 8,
      x: 1,
      y: 1,
    });

    slide2.addShape('roundRect', {
      fill: '64748B',
      h: 0.8,
      hyperlink: { action: 'firstSlide', tooltip: 'Return Home' },
      text: '⬅ Back to Start',
      textOptions: { color: 'FFFFFF', fontSize: 14 },
      w: 2.5,
      x: 1,
      y: 3,
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
});
