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

  it('returns undefined when no animations and no media elements exist', () => {
    expect(buildSlideTiming(undefined, undefined)).toBeUndefined();
    expect(buildSlideTiming([], [])).toBeUndefined();
  });
});
