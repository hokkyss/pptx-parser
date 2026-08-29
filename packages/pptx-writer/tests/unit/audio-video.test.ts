import { describe, expect, it } from 'vitest';
import { emu } from '@hokkyss/pptx-core';
import {
  buildSlideTiming,
  serializeAudio,
  serializeVideo,
} from '../../lib/serializers/audio-video-serializer';

describe('Audio & Video Serializers (Unit Tests)', () => {
  it('serializes audio element to DrawingML pic element', () => {
    const audioNode = serializeAudio(
      {
        audio: {
          mediaId: 'media-audio-1',
          playback: {
            loop: true,
            trigger: 'onClick',
            volume: 0.8,
          },
        },
        elementType: 'audio',
        id: 'shape-1',
        name: 'Audio 1',
        position: { cx: emu(914400), cy: emu(914400), x: emu(100000), y: emu(200000) },
        type: 'picture',
      },
      'rIdAudioLink',
      'rIdMediaEmbed',
      'rIdPosterImage',
    );

    expect(audioNode).toBeDefined();
    expect((audioNode as any)['p:nvPicPr']['p:nvPr']['a:audioFile']['@_r:link']).toBe('rIdAudioLink');
    expect((audioNode as any)['p:nvPicPr']['p:nvPr']['p:extLst']['p:ext']['p14:media']['@_r:embed']).toBe('rIdMediaEmbed');
    expect((audioNode as any)['p:blipFill']['a:blip']['@_r:embed']).toBe('rIdPosterImage');
  });

  it('serializes video element to DrawingML pic element', () => {
    const videoNode = serializeVideo(
      {
        elementType: 'video',
        id: 'shape-2',
        name: 'Video 2',
        position: { cx: emu(3657600), cy: emu(2743200), x: emu(500000), y: emu(600000) },
        type: 'picture',
        video: {
          mediaId: 'media-video-1',
          playback: {
            loop: true,
            muted: true,
            trigger: 'automatic',
          },
        },
      },
      'rIdVideoLink',
      'rIdMediaEmbed',
      'rIdPosterImage',
    );

    expect(videoNode).toBeDefined();
    expect((videoNode as any)['p:nvPicPr']['p:nvPr']['a:videoFile']['@_r:link']).toBe('rIdVideoLink');
    expect((videoNode as any)['p:nvPicPr']['p:nvPr']['p:extLst']['p:ext']['p14:media']['@_r:embed']).toBe('rIdMediaEmbed');
    expect((videoNode as any)['p:blipFill']['a:blip']['@_r:embed']).toBe('rIdPosterImage');
  });

  it('builds slide timing with animations, automatic media, and onClick media', () => {
    const timing = buildSlideTiming(
      [
        {
          delay: 100,
          duration: 600,
          targetShapeId: 'shape-text-1',
          trigger: 'afterPrevious',
          type: 'fade',
        },
        {
          delay: 0,
          duration: 400,
          targetShapeId: 'shape-text-2',
          trigger: 'withPrevious',
          type: 'flyIn',
        },
        {
          delay: 0,
          duration: 500,
          targetShapeId: 'shape-text-3',
          trigger: 'onClick',
          type: 'zoom',
        },
      ],
      [
        {
          id: 'media-auto',
          mediaType: 'video',
          muted: true,
          playback: {
            loop: true,
            trigger: 'automatic',
            volume: 0.5,
          },
        },
        {
          id: 'media-click',
          mediaType: 'audio',
          muted: false,
          playback: {
            loop: false,
            trigger: 'onClick',
            volume: 0.9,
          },
        },
      ],
    );

    expect(timing).toBeDefined();
    expect((timing as any)['p:tnLst']).toBeDefined();
    expect((timing as any)['p:tnLst']['p:par']).toBeDefined();
  });

  it('builds slide timing with start and end trimming', () => {
    const timing = buildSlideTiming(
      undefined,
      [
        {
          id: 'media-trimmed',
          mediaType: 'video',
          muted: false,
          playback: {
            endTime: 12000,
            loop: true,
            startTime: 2000,
            trigger: 'automatic',
            volume: 1.0,
          },
        },
      ],
    );

    expect(timing).toBeDefined();
    const par = (timing as any)['p:tnLst']['p:par'];
    expect(par).toBeDefined();
  });

  it('embeds default audio and video posters when posterImageId is omitted', async () => {
    const { writePptx } = await import('../../lib');
    const doc: any = {
      customXml: [],
      media: [
        {
          data: new Uint8Array([1, 2, 3]),
          fileName: 'audio.mp3',
          filename: 'audio.mp3',
          id: 'audio_1',
          mimeType: 'audio/mpeg',
          path: 'ppt/media/audio.mp3',
        },
        {
          data: new Uint8Array([4, 5, 6]),
          fileName: 'video.mp4',
          filename: 'video.mp4',
          id: 'video_1',
          mimeType: 'video/mp4',
          path: 'ppt/media/video.mp4',
        },
      ],
      metadata: {
        slideCount: 1,
        slideHeight: emu(6858000),
        slideWidth: emu(12192000),
        title: 'Poster Fallback Deck',
      },
      slideLayouts: [],
      slideMasters: [],
      slides: [
        {
          animations: [],
          elements: [
            {
              audio: { mediaId: 'audio_1' },
              elementType: 'audio',
              id: '1',
              isVisible: true,
              name: 'Audio',
              position: { cx: emu(914400), cy: emu(914400), x: emu(100000), y: emu(100000) },
              type: 'picture',
            },
            {
              elementType: 'video',
              id: '2',
              isVisible: true,
              name: 'Video',
              position: { cx: emu(3657600), cy: emu(2743200), x: emu(200000), y: emu(200000) },
              type: 'picture',
              video: { mediaId: 'video_1' },
            },
          ],
          shapes: [],
          slideId: 'rId1',
          slideNumber: 1,
        },
      ],
      themes: [],
    };

    const buf = await writePptx(doc);
    expect(buf).toBeInstanceOf(Uint8Array);
    expect(buf.length).toBeGreaterThan(0);
  });

  it('returns undefined when no animations and no media elements exist', () => {
    expect(buildSlideTiming(undefined, undefined)).toBeUndefined();
    expect(buildSlideTiming([], [])).toBeUndefined();
  });
});
