/** Represents an animation applied to a shape */
export interface PptxAnimation {
  /** Delay in milliseconds. OpenXML: `<p:cTn @_delay>` */
  delay?: number;
  /** Duration in milliseconds. OpenXML: `<p:cTn @_dur>` */
  duration?: number;
  /** Effect type. OpenXML: `<p:animEffect @_filter>` */
  effect: string;
  /** Alias for effect type */
  effectType?: string;
  /** Sequence number */
  sequence: number;
  /** Target shape ID. OpenXML: `<p:spTarget @_spid>` */
  targetShapeId: string;
  /** Trigger type (e.g., onClick, withPrevious). OpenXML: `<p:cBhvr><p:cTn @_evt>` */
  trigger: string;
}

/** Standard OpenXML Slide Transition Type names */
export type PptxTransitionType
  = 'blinds'
    | 'checker'
    | 'circle'
    | 'comb'
    | 'cover'
    | 'cut'
    | 'diamond'
    | 'dissolve'
    | 'fade'
    | 'newsflash'
    | 'none'
    | 'plus'
    | 'pull'
    | 'push'
    | 'randomBar'
    | 'split'
    | 'wedge'
    | 'wheel'
    | 'wipe'
    | 'zoom'
    | ({} & string);

/** Direction for directional slide transitions */
export type PptxTransitionDirection = 'down' | 'horz' | 'in' | 'left' | 'out' | 'right' | 'up' | 'vert' | ({} & string);

/** Speed preset for slide transitions */
export type PptxTransitionSpeed = 'fast' | 'med' | 'medium' | 'slow' | ({} & string);

/** Represents a slide transition */
export interface PptxTransition {
  /** Auto-advance time in milliseconds. OpenXML: `<p:transition @_advTm>` */
  advanceAfterMs?: number;
  /** Whether to advance on click. OpenXML: `<p:transition @_advClick>` */
  advanceOnClick?: boolean;
  /** Transition direction (e.g. 'left', 'right', 'up', 'down', 'horz', 'vert', 'in', 'out'). */
  direction?: PptxTransitionDirection;
  /** Legacy duration in milliseconds. */
  duration?: number;
  /** Transition duration in milliseconds. */
  durationMs?: number;
  /** Transition speed preset. OpenXML: `<p:transition @_spd>` */
  speed?: PptxTransitionSpeed;
  /** Spokes count for wheel transition (1, 2, 3, 4, 8) */
  spokes?: number;
  /** For fade transition: whether to fade through black. OpenXML: `<p:fade @_thruBlk="1">` */
  throughBlack?: boolean;
  /** Transition type (e.g., 'fade', 'push', 'wipe'). OpenXML: `<p:transition><p:*>` */
  type: PptxTransitionType;
}
