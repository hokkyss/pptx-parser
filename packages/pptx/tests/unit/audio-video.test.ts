import { describe, expect, it } from 'vitest';
import { inches } from '@hokkyss/pptx-core';
import { Presentation } from '../../lib/presentation';

describe('Audio & Video Media Embedding (Unit & Integration Tests)', () => {
  const dummyAudioData = new Uint8Array([0xff, 0xfb, 0x90, 0x44, 0x00, 0x00]);
  const dummyVideoData = new Uint8Array([0x00, 0x00, 0x00, 0x20, 0x66, 0x74, 0x79, 0x70, 0x6d, 0x70, 0x34, 0x32]);

  it('embeds an audio file with default parameters', () => {
    const pres = Presentation.create();
    const slide = pres.addSlide();

    slide.addAudio(dummyAudioData, {
      fileName: 'background_music.mp3',
      mimeType: 'audio/mpeg',
    });

    const elements = slide.getElements();
    expect(elements.length).toBe(1);

    const audioEl = elements[0];
    expect(audioEl.elementType).toBe('audio');
    if (audioEl.elementType === 'audio') {
      expect(audioEl.audio.mimeType).toBe('audio/mpeg');
      expect(audioEl.audio.mediaId).toBeDefined();
      expect(audioEl.position.cx).toBe(914400); // 1 inch default
      expect(audioEl.position.cy).toBe(914400); // 1 inch default
      expect(audioEl.position.x).toBe(457200); // 0.5 inch default
      expect(audioEl.position.y).toBe(457200); // 0.5 inch default
      expect(audioEl.audio.playback?.trigger).toBe('onClick');
    }

    expect(pres.ast.media.length).toBe(1);
    expect(pres.ast.media[0].fileName).toBe('background_music.mp3');
    expect(pres.ast.media[0].mimeType).toBe('audio/mpeg');
    expect(pres.ast.media[0].data).toEqual(dummyAudioData);
  });

  it('customizes audio playback options (loop, volume, trim, hideWhenDone)', () => {
    const pres = Presentation.create();
    const slide = pres.addSlide();

    slide.addAudio(dummyAudioData, {
      endTime: 15000,
      fileName: 'voiceover.wav',
      h: inches(1.5),
      hideWhenDone: true,
      loop: true,
      mimeType: 'audio/x-wav',
      name: 'Narrator Track',
      showWhenStopped: true,
      startTime: 1000,
      trigger: 'automatic',
      volume: 0.75,
      w: inches(1.5),
      x: inches(2),
      y: inches(3),
    });

    const elements = slide.getElements();
    expect(elements.length).toBe(1);

    const audioEl = elements[0];
    expect(audioEl.elementType).toBe('audio');
    if (audioEl.elementType === 'audio') {
      expect(audioEl.name).toBe('Narrator Track');
      expect(audioEl.position.x).toBe(1828800); // 2 inches
      expect(audioEl.position.y).toBe(2743200); // 3 inches
      expect(audioEl.position.cx).toBe(1371600); // 1.5 inches
      expect(audioEl.position.cy).toBe(1371600); // 1.5 inches
      expect(audioEl.audio.playback?.trigger).toBe('automatic');
      expect(audioEl.audio.playback?.loop).toBe(true);
      expect(audioEl.audio.playback?.volume).toBe(0.75);
      expect(audioEl.audio.playback?.startTime).toBe(1000);
      expect(audioEl.audio.playback?.endTime).toBe(15000);
      expect(audioEl.audio.playback?.hideWhenDone).toBe(true);
      expect(audioEl.audio.playback?.showWhenStopped).toBe(true);
    }
  });

  it('embeds a video file with default dimensions', () => {
    const pres = Presentation.create();
    const slide = pres.addSlide();

    slide.addVideo(dummyVideoData, {
      fileName: 'product_demo.mp4',
      mimeType: 'video/mp4',
    });

    const elements = slide.getElements();
    expect(elements.length).toBe(1);

    const videoEl = elements[0];
    expect(videoEl.elementType).toBe('video');
    if (videoEl.elementType === 'video') {
      expect(videoEl.video.mimeType).toBe('video/mp4');
      expect(videoEl.position.cx).toBe(3657600); // 4 inches default
      expect(videoEl.position.cy).toBe(2743200); // 3 inches default
      expect(videoEl.position.x).toBe(2286000); // 2.5 inches default
      expect(videoEl.position.y).toBe(1371600); // 1.5 inches default
      expect(videoEl.video.playback?.trigger).toBe('onClick');
    }

    expect(pres.ast.media.length).toBe(1);
    expect(pres.ast.media[0].fileName).toBe('product_demo.mp4');
    expect(pres.ast.media[0].mimeType).toBe('video/mp4');
  });

  it('customizes video options (muted, loop, automatic, custom bounds)', () => {
    const pres = Presentation.create();
    const slide = pres.addSlide();

    slide.addVideo(dummyVideoData, {
      endTime: 30000,
      fileName: 'hero_reel.mp4',
      h: inches(4.5),
      hideWhenDone: false,
      loop: true,
      mimeType: 'video/mp4',
      muted: true,
      name: 'Hero Reel',
      startTime: 5000,
      trigger: 'automatic',
      w: inches(8),
      x: inches(1),
      y: inches(1),
    });

    const elements = slide.getElements();
    expect(elements.length).toBe(1);

    const videoEl = elements[0];
    expect(videoEl.elementType).toBe('video');
    if (videoEl.elementType === 'video') {
      expect(videoEl.name).toBe('Hero Reel');
      expect(videoEl.video.muted).toBe(true);
      expect(videoEl.video.playback?.loop).toBe(true);
      expect(videoEl.video.playback?.trigger).toBe('automatic');
      expect(videoEl.video.playback?.startTime).toBe(5000);
      expect(videoEl.video.playback?.endTime).toBe(30000);
      expect(videoEl.position.cx).toBe(7315200); // 8 inches
      expect(videoEl.position.cy).toBe(4114800); // 4.5 inches
    }
  });

  it('enforces ID uniqueness for audio and video elements on the same slide', () => {
    const pres = Presentation.create();
    const slide = pres.addSlide();

    slide.addAudio(dummyAudioData, {
      fileName: 'track1.mp3',
      id: 'media-node-1',
      mimeType: 'audio/mpeg',
    });

    expect(() => {
      slide.addAudio(dummyAudioData, {
        fileName: 'track2.mp3',
        id: 'media-node-1',
        mimeType: 'audio/mpeg',
      });
    }).toThrowError(/Duplicate element ID "media-node-1"/);

    expect(() => {
      slide.addVideo(dummyVideoData, {
        fileName: 'clip1.mp4',
        id: 'media-node-1',
        mimeType: 'video/mp4',
      });
    }).toThrowError(/Duplicate element ID "media-node-1"/);
  });

  it('supports round-trip save and load of embedded audio and video', async () => {
    const pres = Presentation.create();
    const slide = pres.addSlide();

    slide.addAudio(dummyAudioData, {
      fileName: 'audio_sample.mp3',
      id: 'audio-1',
      mimeType: 'audio/mpeg',
      name: 'Sample Audio Track',
      x: inches(0.5),
      y: inches(0.5),
    });

    slide.addVideo(dummyVideoData, {
      fileName: 'video_sample.mp4',
      id: 'video-1',
      mimeType: 'video/mp4',
      name: 'Sample Video Track',
      x: inches(2),
      y: inches(2),
    });

    const buffer = await pres.toBuffer();
    expect(buffer).toBeInstanceOf(Uint8Array);
    expect(buffer.byteLength).toBeGreaterThan(0);

    const loaded = await Presentation.load(buffer);
    const loadedSlides = loaded.slides;
    expect(loadedSlides.length).toBe(1);

    const loadedElements = loadedSlides[0].getElements();
    expect(loadedElements.length).toBe(2);

    const loadedAudio = loadedElements.find((el) => el.elementType === 'audio');
    const loadedVideo = loadedElements.find((el) => el.elementType === 'video');

    expect(loadedAudio).toBeDefined();
    expect(loadedAudio?.name).toBe('Sample Audio Track');
    if (loadedAudio?.elementType === 'audio') {
      expect(loadedAudio.audio.mimeType).toBe('audio/mpeg');
    }

    expect(loadedVideo).toBeDefined();
    expect(loadedVideo?.name).toBe('Sample Video Track');
    if (loadedVideo?.elementType === 'video') {
      expect(loadedVideo.video.mimeType).toBe('video/mp4');
    }

    // Verify media assets were preserved in the package (including default poster image)
    const mediaList = loaded.ast.media;
    expect(mediaList.length).toBeGreaterThanOrEqual(2);
    expect(mediaList.some((m) => m.fileName === 'audio_sample.mp3' || m.filename === 'audio_sample.mp3')).toBe(true);
    expect(mediaList.some((m) => m.fileName === 'video_sample.mp4' || m.filename === 'video_sample.mp4')).toBe(true);
  });
});
