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

/** Represents a slide transition */
export interface PptxTransition {
  /** Auto-advance time in milliseconds. OpenXML: `<p:transition @_advTm>` */
  advanceAfterMs?: number;
  /** Whether to advance on click. OpenXML: `<p:transition @_advClick>` */
  advanceOnClick?: boolean;
  /** Transition duration in milliseconds. OpenXML: `<p:transition @_spd>` */
  duration?: number;
  /** Transition speed. OpenXML: `<p:transition @_spd>` */
  speed?: string;
  /** Transition type (e.g., 'fade', 'push', 'wipe'). OpenXML: `<p:transition><p:*>` */
  type: string;
}
