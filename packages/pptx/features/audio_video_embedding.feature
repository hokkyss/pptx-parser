Feature: Audio and Video Media Embedding
  As a presentation developer using @hokkyss/pptx
  I want to embed audio and video media files directly into PowerPoint slides
  So that presentations can feature rich interactive sound clips, background music, video demos, and presentations

  Background:
    Given a new presentation instance is initialized
    And a slide is added to the presentation

  # =========================================================================
  # 1. AUDIO EMBEDDING & PLAYBACK CONTROLS
  # =========================================================================

  Scenario: Embedding an audio file with default settings
    Given raw binary audio data representing "soundtrack.mp3"
    When I call slide.addAudio(audioData, options) with:
      | fileName | soundtrack.mp3 |
      | mimeType | audio/mpeg     |
    Then an audio element is registered on the slide with elementType "audio"
    And the audio media asset is added to presentation.ast.media with mimeType "audio/mpeg"
    And the audio placeholder dimensions default to 1" width and 1" height
    And the audio position defaults to (x: 0.5", y: 0.5")

  Scenario: Customizing audio playback triggers and timing controls
    Given raw binary audio data representing "narration.wav"
    When I call slide.addAudio(audioData, options) with:
      | fileName     | narration.wav |
      | mimeType     | audio/x-wav   |
      | trigger      | automatic     |
      | loop         | true          |
      | volume       | 0.8           |
      | startTime    | 1500          |
      | endTime      | 12000         |
      | hideWhenDone | true          |
    Then the audio playback configuration contains:
      | trigger      | automatic |
      | loop         | true      |
      | volume       | 0.8       |
      | startTime    | 1500      |
      | endTime      | 12000     |
      | hideWhenDone | true      |

  # =========================================================================
  # 2. VIDEO EMBEDDING & PLAYBACK CONTROLS
  # =========================================================================

  Scenario: Embedding a video file with default dimensions and position
    Given raw binary video data representing "demo.mp4"
    When I call slide.addVideo(videoData, options) with:
      | fileName | demo.mp4  |
      | mimeType | video/mp4 |
    Then a video element is registered on the slide with elementType "video"
    And the video media asset is added to presentation.ast.media with mimeType "video/mp4"
    And the video placeholder dimensions default to 4" width and 3" height
    And the video position defaults to (x: 2.5", y: 1.5")

  Scenario: Embedding a video with custom dimensions, muted audio, and loop
    Given raw binary video data representing "intro.mp4"
    When I call slide.addVideo(videoData, options) with:
      | fileName | intro.mp4  |
      | mimeType | video/mp4  |
      | x        | 1.0        |
      | y        | 1.0        |
      | w        | 6.0        |
      | h        | 4.5        |
      | muted    | true       |
      | loop     | true       |
      | trigger  | automatic  |
    Then the video element position is (x: 1.0", y: 1.0", w: 6.0", h: 4.5")
    And the video is configured with muted=true and loop=true

  # =========================================================================
  # 3. ID UNIQUENESS & COLLISION PREVENTION
  # =========================================================================

  Scenario: Throwing early on duplicate audio element ID on the same slide
    Given an audio element is added with id "media-track-1"
    When I attempt to add another audio element with id "media-track-1" to the same slide
    Then an error should be thrown matching 'Duplicate element ID "media-track-1" detected on Slide 1'

  Scenario: Throwing early on duplicate video element ID on the same slide
    Given a shape is added with id "player-frame"
    When I attempt to add a video element with id "player-frame" to the same slide
    Then an error should be thrown matching 'Duplicate element ID "player-frame" detected on Slide 1'

  # =========================================================================
  # 4. OPENXML SERIALIZATION & ROUND-TRIP FIDELITY
  # =========================================================================

  Scenario: Serializing audio elements into OpenXML DrawingML and relationship parts
    Given a slide containing an audio element for "sound.mp3"
    When the presentation is exported to PPTX format
    Then the slide XML contains `<p:audioFile @_r:link="rId2"/>`
    And the slide XML contains `<p14:media @_r:embed="rId3"/>` inside `<p:extLst>`
    And the slide relationships contain an audio relationship pointing to `../media/sound.mp3`
    And the slide relationships contain a media relationship pointing to `../media/sound.mp3`
    And `[Content_Types].xml` includes the default MIME mapping for `mp3`

  Scenario: Serializing video elements into OpenXML DrawingML and relationship parts
    Given a slide containing a video element for "clip.mp4"
    When the presentation is exported to PPTX format
    Then the slide XML contains `<p:videoFile @_r:link="rId2"/>`
    And the slide XML contains `<p14:media @_r:embed="rId3"/>` inside `<p:extLst>`
    And the slide relationships contain a video relationship pointing to `../media/clip.mp4`
    And the slide relationships contain a media relationship pointing to `../media/clip.mp4`
    And `[Content_Types].xml` includes the default MIME mapping for `mp4`

  Scenario: Round-trip parsing of embedded audio and video elements
    Given a PPTX file containing embedded audio and video slides
    When the presentation is parsed by @hokkyss/pptx-reader
    Then the parsed AST contains elements with elementType "audio" and "video"
    And the media assets are extracted into document.media with correct binary data and MIME types
