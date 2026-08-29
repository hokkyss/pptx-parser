---
title: "Media & Images"
description: "Embed high-resolution images, native audio tracks, video player containers, and custom poster thumbnails in @hokkyss/pptx."
order: 7
section: "authoring"
---

# Media & Images

`@hokkyss/pptx` provides a native, zero-dependency multimedia embedding engine. Embed images, audio files, and video player containers directly from binary data (`Uint8Array` or `ArrayBuffer`) with comprehensive playback controls and custom poster thumbnails across **Node.js, Web Browsers, and Cloudflare Workers**.

---

## Embedding Images (`addImage`)

Embed PNG, JPEG, SVG, WebP, and GIF images onto your slide canvas.

```typescript
import { Presentation, inches, degrees } from '@hokkyss/pptx';
import * as fs from 'node:fs';

const pres = Presentation.create();
const slide = pres.addSlide();

const imageBytes = fs.readFileSync('hero.png');

slide.addImage(imageBytes, {
  fileName: 'hero.png',
  x: inches(1),
  y: inches(1),
  w: inches(5),
  h: inches(3),
  rotation: degrees(0),
  hyperlink: 'https://example.com',
});
```

### `AddImageOptions`

| Option | Type | Default | Description |
| :--- | :--- | :--- | :--- |
| `fileName` | `string` | **Required** | Asset file name inside `ppt/media/` (e.g. `'hero.png'`). |
| `x` | `Inches` | `inches(1)` | Left offset from slide edge. |
| `y` | `Inches` | `inches(1)` | Top offset from slide edge. |
| `w` | `Inches` | `inches(4)` | Image width. |
| `h` | `Inches` | `inches(3)` | Image height. |
| `rotation` | `Degrees` | `degrees(0)` | Clockwise rotation angle. |
| `hyperlink` | `PptxHyperlink \| string` | `undefined` | Clickable link URL or internal slide jump. |
| `id` | `string` | Auto | Unique element ID on this slide. |
| `name` | `string` | Auto | Descriptive layer name in PowerPoint selection pane. |
| `zIndex` | `number` | Sequential | Visual layering index. |

---

## Embedding Audio (`addAudio`)

Embed background audio tracks, voiceover narrations, and sound effects with full control over playback triggers, looping, volume, and trimming.

```typescript
const audioBytes = fs.readFileSync('soundtrack.mp3');

slide.addAudio(audioBytes, {
  fileName: 'soundtrack.mp3',
  mimeType: 'audio/mpeg',
  x: inches(1.0),
  y: inches(1.0),
  w: inches(1.0),
  h: inches(1.0),
  trigger: 'onClick',    // 'onClick' (default) or 'automatic'
  loop: true,           // Loop indefinitely
  volume: 0.8,          // 0.0 (silent) to 1.0 (full volume)
  startTime: 1000,      // Clip start offset in milliseconds
  endTime: 15000,       // Clip end position in milliseconds
  hideWhenDone: false,  // Whether to hide icon after audio finishes
  showWhenStopped: true // Show speaker icon when stopped
});
```

### Custom Audio Poster / Album Art Thumbnail

By default, an optimized, circular speaker icon is generated automatically. You can provide custom album artwork or podcast cover art via the `poster` option:

```typescript
const albumCoverBytes = fs.readFileSync('album_art.png');

slide.addAudio(audioBytes, {
  fileName: 'podcast_ep1.mp3',
  mimeType: 'audio/mpeg',
  poster: {
    data: albumCoverBytes,
    fileName: 'podcast_art.png',
    mimeType: 'image/png',
  },
  trigger: 'onClick',
});
```

### `AddAudioOptions`

| Option | Type | Default | Description |
| :--- | :--- | :--- | :--- |
| `fileName` | `string` | **Required** | Audio file name inside `ppt/media/` (e.g. `'narration.mp3'`). |
| `mimeType` | `string` | **Required** | MIME type (e.g. `'audio/mpeg'`, `'audio/x-wav'`, `'audio/mp4'`). |
| `poster` | `PptxPosterOption` | Auto speaker | Custom poster frame / thumbnail image object (`{ data, mimeType?, fileName? }`). |
| `trigger` | `'onClick' \| 'automatic'` | `'onClick'` | Playback trigger. `'onClick'` plays on click; `'automatic'` plays on slide load. |
| `loop` | `boolean` | `false` | Whether to repeat playback indefinitely. |
| `volume` | `number` | `0.8` | Audio volume from `0.0` (mute) to `1.0` (maximum). |
| `startTime` | `number` | `undefined` | Start offset in milliseconds (`<p14:trim st="..."/>`). |
| `endTime` | `number` | `undefined` | End position in milliseconds (`<p14:trim end="..."/>`). |
| `hideWhenDone` | `boolean` | `false` | Hide the speaker icon once audio playback finishes. |
| `showWhenStopped` | `boolean` | `true` | Show placeholder icon when stopped. |
| `x` / `y` | `Inches` | `inches(0.5)` | Position coordinates on slide. |
| `w` / `h` | `Inches` | `inches(1.0)` | Dimensions on slide. |

---

## Embedding Video (`addVideo`)

Embed motion graphics, product walkthroughs, and screen recordings in cross-platform video containers.

```typescript
const videoBytes = fs.readFileSync('product_demo.mp4');

slide.addVideo(videoBytes, {
  fileName: 'product_demo.mp4',
  mimeType: 'video/mp4',
  x: inches(2.0),
  y: inches(1.5),
  w: inches(6.0),
  h: inches(4.5),
  trigger: 'automatic', // Starts immediately when slide is presented
  loop: true,          // Loop video playback
  muted: true,         // Mute audio track on presentation
});
```

### Custom Cinematic Video Poster Frame

Provide custom high-resolution poster thumbnails (e.g. title cards, 4K cover graphics) for the video placeholder:

```typescript
const coverJpgBytes = fs.readFileSync('video_thumbnail.jpg');

slide.addVideo(videoBytes, {
  fileName: 'keynote.mp4',
  mimeType: 'video/mp4',
  poster: {
    data: coverJpgBytes,
    fileName: 'keynote_poster.jpg',
    mimeType: 'image/jpeg',
  },
  x: inches(1.5),
  y: inches(1.0),
  w: inches(7.0),
  h: inches(4.0),
  trigger: 'onClick',
});
```

### `AddVideoOptions`

| Option | Type | Default | Description |
| :--- | :--- | :--- | :--- |
| `fileName` | `string` | **Required** | Video file name inside `ppt/media/` (e.g. `'demo.mp4'`). |
| `mimeType` | `string` | **Required** | MIME type (e.g. `'video/mp4'`, `'video/quicktime'`). |
| `poster` | `PptxPosterOption` | Auto player | Custom poster thumbnail image object (`{ data, mimeType?, fileName? }`). |
| `trigger` | `'onClick' \| 'automatic'` | `'onClick'` | Playback trigger. |
| `loop` | `boolean` | `false` | Loop playback continuously. |
| `muted` | `boolean` | `false` | Mute the video audio track during playback. |
| `startTime` | `number` | `undefined` | Trim start offset in milliseconds. |
| `endTime` | `number` | `undefined` | Trim end offset in milliseconds. |
| `hideWhenDone` | `boolean` | `false` | Hide video frame after playback concludes. |
| `x` / `y` | `Inches` | `2.5" / 1.5"` | Position coordinates (centered by default on standard 16\:9 slides). |
| `w` / `h` | `Inches` | `4.0" / 3.0"` | Placeholder dimensions. |

---

## Supported Format Compatibility

| Media Type | Recommended Extension | Supported In | Notes |
| :--- | :--- | :--- | :--- |
| **Video** | `.mp4` (H.264 / AAC) | Windows, macOS, iOS, Android, Web | **Most portable**. Plays natively across all Office platforms. |
| **Video** | `.mov` | macOS, iOS | Uses Apple QuickTime container. |
| **Video** | `.wmv` / `.avi` | Windows | Windows Media Video container. |
| **Audio** | `.mp3` | Windows, macOS, iOS, Android, Web | Universal MPEG audio standard. |
| **Audio** | `.wav` | Windows, macOS | Uncompressed PCM audio. |
| **Audio** | `.m4a` | Windows, macOS, iOS | AAC audio in MPEG-4 container. |
| **Image / Poster** | `.png`, `.jpg`, `.jpeg` | All platforms | Universal lossless and photographic image support. |

---

## Round-Trip Fidelity

When opening an existing `.pptx` file with `@hokkyss/pptx-reader`, embedded audio and video media files together with their custom poster frames are parsed into AST elements with full fidelity:

```typescript
const pres = await Presentation.load(pptxBuffer);
const slide = pres.slides[0];

for (const el of slide.getElements()) {
  if (el.elementType === 'video') {
    console.log(`Video media ID: ${el.video.mediaId}, poster ID: ${el.video.posterImageId}`);
  } else if (el.elementType === 'audio') {
    console.log(`Audio media ID: ${el.audio.mediaId}, trigger: ${el.audio.playback?.trigger}`);
  }
}
```

