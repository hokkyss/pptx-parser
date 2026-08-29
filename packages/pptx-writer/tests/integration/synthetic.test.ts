import type { PptxDocument } from '@hokkyss/pptx-core';
import { emu, emuDegree, hundredthsPoint } from '@hokkyss/pptx-core';
import { parsePptx } from '@hokkyss/pptx-reader';
import { describe, expect, it } from 'vitest';
import { writePptx } from '../../lib';

describe('Synthetic Presentation Generation', () => {
  it('generates a valid PPTX document from scratch that parsePptx can read', async () => {
    const doc: PptxDocument = {
      customXml: [],
      media: [],
      metadata: {
        creator: 'Writer Test',
        slideCount: 2,
        slideHeight: emu(6858000),
        slideWidth: emu(12192000), // 16:9 widescreen
        title: 'Quarterly Kickoff',
      },
      slideLayouts: [],
      slideMasters: [],
      slides: [
        {
          animations: [],
          elements: [
            {
              elementType: 'shape',
              fill: {
                solidColor: { type: 'srgb', value: '4F81BD' },
                type: 'solid',
              },
              id: '2',
              isVisible: true,
              name: 'Title Box',
              position: {
                cx: emu(8000000),
                cy: emu(1500000),
                x: emu(1000000),
                y: emu(1000000),
              },
              rotation: emuDegree(0),
              textBody: {
                bodyProperties: {},
                paragraphs: [
                  {
                    properties: {
                      alignment: 'center',
                    },
                    runs: [
                      {
                        properties: {
                          bold: true,
                          color: 'FFFFFF',
                          fontSize: hundredthsPoint(2800),
                        },
                        text: 'Welcome to PPTX Writer',
                      },
                    ],
                  },
                ],
              },
              type: 'shape',
              zIndex: 0,
            },
          ],
          shapes: [],
          slideId: 'rId1',
          slideNumber: 1,
        },
        {
          animations: [],
          elements: [
            {
              elementType: 'table',
              id: '3',
              isVisible: true,
              name: 'Data Table',
              position: {
                cx: emu(6000000),
                cy: emu(3000000),
                x: emu(1000000),
                y: emu(1000000),
              },
              rotation: emuDegree(0),
              table: {
                columnWidths: [emu(3000000), emu(3000000)],
                rows: [
                  {
                    cells: [
                      { textBody: { bodyProperties: {}, paragraphs: [{ properties: {}, runs: [{ properties: {}, text: 'Metric' }] }] } },
                      { textBody: { bodyProperties: {}, paragraphs: [{ properties: {}, runs: [{ properties: {}, text: 'Value' }] }] } },
                    ],
                    height: emu(1000000),
                  },
                  {
                    cells: [
                      { textBody: { bodyProperties: {}, paragraphs: [{ properties: {}, runs: [{ properties: {}, text: 'Growth' }] }] } },
                      { textBody: { bodyProperties: {}, paragraphs: [{ properties: {}, runs: [{ properties: {}, text: '120%' }] }] } },
                    ],
                    height: emu(1000000),
                  },
                ],
              },
              type: 'graphicFrame',
              zIndex: 0,
            },
          ],
          shapes: [],
          slideId: 'rId2',
          slideNumber: 2,
        },
        {
          animations: [],
          elements: [
            {
              audio: {
                mediaId: 'sound_1',
                playback: { loop: true, trigger: 'onClick', volume: 0.8 },
                posterImageId: 'poster_1',
              },
              elementType: 'audio',
              id: '4',
              isVisible: true,
              name: 'Audio Track',
              position: {
                cx: emu(914400),
                cy: emu(914400),
                x: emu(1000000),
                y: emu(1000000),
              },
              rotation: emuDegree(0),
              type: 'picture',
              zIndex: 0,
            },
            {
              elementType: 'video',
              id: '5',
              isVisible: true,
              name: 'Video Clip',
              position: {
                cx: emu(3657600),
                cy: emu(2743200),
                x: emu(3000000),
                y: emu(1000000),
              },
              rotation: emuDegree(0),
              type: 'picture',
              video: {
                mediaId: 'video_1',
                playback: { loop: true, muted: true, trigger: 'automatic' },
              },
              zIndex: 1,
            },
          ],
          shapes: [],
          slideId: 'rId3',
          slideNumber: 3,
        },
      ],
      themes: [],
    };

    doc.media.push(
      {
        data: new Uint8Array([1, 2, 3]),
        fileName: 'track.mp3',
        filename: 'track.mp3',
        id: 'sound_1',
        mimeType: 'audio/mpeg',
        path: 'ppt/media/track.mp3',
      },
      {
        data: new Uint8Array([4, 5, 6]),
        fileName: 'poster.png',
        filename: 'poster.png',
        id: 'poster_1',
        mimeType: 'image/png',
        path: 'ppt/media/poster.png',
      },
      {
        data: new Uint8Array([7, 8, 9]),
        fileName: 'clip.mp4',
        filename: 'clip.mp4',
        id: 'video_1',
        mimeType: 'video/mp4',
        path: 'ppt/media/clip.mp4',
      },
    );

    const buffer = await writePptx(doc);
    expect(buffer).toBeInstanceOf(Uint8Array);
    expect(buffer.length).toBeGreaterThan(0);

    // Read back with pptx-reader
    const parsed = await parsePptx(buffer);
    expect(parsed.slides).toHaveLength(3);
    expect(parsed.metadata.title).toBe('Quarterly Kickoff');
    expect(parsed.metadata.slideWidth).toBe(12192000);
    expect(parsed.metadata.slideHeight).toBe(6858000);

    // Verify slide 1 shape & text
    const s1 = parsed.slides[0];
    expect(s1.elements.length).toBeGreaterThan(0);
    const shape = s1.elements[0];
    expect(shape.elementType).toBe('shape');
    expect(shape.textBody?.paragraphs[0].runs[0].text).toContain('Welcome to PPTX Writer');

    // Verify slide 2 table
    const s2 = parsed.slides[1];
    const tableElem = s2.elements[0];
    expect(tableElem.elementType).toBe('table');
    if (tableElem.elementType === 'table') {
      expect(tableElem.table.rows).toHaveLength(2);
    }

    // Verify slide 3 audio and video
    const s3 = parsed.slides[2];
    expect(s3.elements).toHaveLength(2);
    expect(s3.elements[0].elementType).toBe('audio');
    expect(s3.elements[1].elementType).toBe('video');
  });
});
