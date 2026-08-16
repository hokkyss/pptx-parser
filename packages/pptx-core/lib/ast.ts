import type {
  PptxAnimation,
  PptxTransition,
  PptxTransitionDirection,
  PptxTransitionSpeed,
  PptxTransitionType,
} from './animation';
import type { PptxChart } from './chart';
import type { PptxFill } from './color';
import type { PptxMediaAsset } from './media';
import type { PptxTable } from './table';
import type { PptxHyperlink, PptxHyperlinkAction, PptxTextBody } from './text';
import type { PptxLine, PptxTheme } from './theme';
import type { Emu, EmuDegree, HundredthsPoint, ThousandthsPercent } from './units';

export type {
  Emu,
  EmuDegree,
  HundredthsPoint,
  PptxAnimation,
  PptxChart,
  PptxFill,
  PptxHyperlink,
  PptxHyperlinkAction,
  PptxLine,
  PptxMediaAsset,
  PptxTable,
  PptxTextBody,
  PptxTheme,
  PptxTransition,
  PptxTransitionDirection,
  PptxTransitionSpeed,
  PptxTransitionType,
  ThousandthsPercent,
};

/** Represents background properties */
export interface PptxBackground {
  /** Background fill. OpenXML: `<p:bg><p:bgPr>` */
  fill?: PptxFill;
}

/** Shape lock settings. OpenXML: `<a:spLocks>`, `<a:picLocks>`, `<a:grpSpLocks>` */
export interface PptxShapeLocks {
  /** Prevent adjusting handles. OpenXML: `@_noAdjustHandles="1"` */
  noAdjustHandles?: boolean;
  /** Prevent aspect ratio change. OpenXML: `@_noChangeAspect="1"` */
  noChangeAspect?: boolean;
  /** Prevent shape type change. OpenXML: `@_noChangeShapeType="1"` */
  noChangeShapeType?: boolean;
  /** Prevent cropping (for pictures). OpenXML: `@_noCrop="1"` */
  noCrop?: boolean;
  /** Prevent editing points. OpenXML: `@_noEditPoints="1"` */
  noEditPoints?: boolean;
  /** Prevent grouping. OpenXML: `@_noGrp="1"` */
  noGrp?: boolean;
  /** Prevent movement/position change. OpenXML: `@_noMove="1"` */
  noMove?: boolean;
  /** Prevent resizing. OpenXML: `@_noResize="1"` */
  noResize?: boolean;
  /** Prevent rotation. OpenXML: `@_noRot="1"` */
  noRot?: boolean;
  /** Prevent selection. OpenXML: `@_noSelect="1"` */
  noSelect?: boolean;
  /** Prevent ungrouping (for groups). OpenXML: `@_noUngrp="1"` */
  noUngrp?: boolean;
}

/** Custom geometry path command */
export interface PptxPathCommand {
  /** Path command type ('moveTo' | 'lineTo' | 'cubicTo' | 'close'). OpenXML: `<a:moveTo>`, `<a:lnTo>`, `<a:cubicBezTo>`, `<a:close>` */
  command: string;
  /** Points for the path command. OpenXML: `<a:pt @_x @_y>` */
  points: Array<{ x: Emu; y: Emu }>;
}

/** Represents shape geometry */
export interface PptxGeometry {
  /** Adjustment values map. OpenXML: `<a:avLst><a:gd>` */
  adjustments?: Record<string, number>;
  /** Custom geometry path commands. OpenXML: `<a:custGeom><a:path>` */
  customPathCommands?: PptxPathCommand[];
  /** Custom geometry path data string. OpenXML: `<a:custGeom><a:pathLst>` */
  customPathData?: string;
  /** Preset geometry name (e.g., 'rect', 'ellipse'). OpenXML: `<a:prstGeom @_prst>` */
  presetGeometry?: string;
}

/** Represents a picture */
export interface PptxPicture {
  /** Alpha transparency. OpenXML: `<a:blip><a:alphaModFix>` */
  alpha?: ThousandthsPercent;
  /** Blip fill properties. OpenXML: `<p:blipFill>` */
  blipFill?: Record<string, unknown>;
  /** Image crop settings. OpenXML: `<a:srcRect>` */
  crop?: {
    bottom?: ThousandthsPercent;
    left?: ThousandthsPercent;
    right?: ThousandthsPercent;
    top?: ThousandthsPercent;
  };
  /** Reference ID to the media asset. OpenXML: `<a:blip @_r:embed>` */
  mediaId: string;
}

/** Represents a placeholder within a layout or master */
export interface PptxPlaceholder {
  /** Placeholder index. OpenXML: `<p:ph @_idx>` */
  idx?: number | string;
  /** Placeholder type (e.g., 'title', 'body', 'subTitle', 'ctrTitle'). OpenXML: `<p:ph @_type>` */
  type: string;
}

/** Base properties common to all visual slide elements */
export interface PptxBaseElement {
  /** Direct reference ID to embedded picture asset (e.g. 'rId2'). OpenXML: `<a:blip @_r:embed>` */
  blipEmbedId?: string;
  /** Chart data. OpenXML: `<c:chartSpace>` inside `<p:graphicFrame>` */
  chart?: PptxChart;
  /** Child elements if group shape. OpenXML: `<p:grpSp>` child nodes */
  children?: PptxElement[];
  /** Fill properties. OpenXML: `<p:spPr><a:solidFill>`, `<a:gradFill>`, etc. */
  fill?: PptxFill;
  /** Geometry properties. OpenXML: `<p:spPr><a:prstGeom>` or `<a:custGeom>` */
  geometry?: PptxGeometry;
  /** Element ID. OpenXML: `<p:cNvPr @_id>` */
  id: string;
  /** Shape or element hyperlink. OpenXML: `<p:cNvPr><a:hlinkClick>` */
  hyperlink?: PptxHyperlink | string;
  /** Whether the element is locked in PowerPoint. OpenXML: `<a:spLocks>`, `<a:picLocks>`, `<a:grpSpLocks>` */
  isLocked?: boolean;
  /** Whether the element is visible in PowerPoint rendering. OpenXML: `<p:cNvPr @_hidden="1">` (defaults to true) */
  isVisible: boolean;
  /** Origin layer source when resolved across Master -> Layout -> Slide */
  layerSource?: 'layout' | 'master' | 'slide';
  /** Line properties. OpenXML: `<p:spPr><a:ln>` */
  line?: PptxLine;
  /** Detailed lock settings. OpenXML: `<a:spLocks>`, `<a:picLocks>`, `<a:grpSpLocks>` */
  locks?: PptxShapeLocks;
  /** Element name. OpenXML: `<p:cNvPr @_name>` */
  name: string;
  /** Picture properties. OpenXML: `<p:pic>` */
  picture?: PptxPicture;
  /** Placeholder properties, if this element is a placeholder. OpenXML: `<p:nvPr><p:ph>` */
  placeholder?: PptxPlaceholder;
  /** Position and dimensions in EMU. OpenXML: `<a:off @_x @_y>`, `<a:ext @_cx @_cy>` */
  position: {
    /** Width in EMU. OpenXML: `<a:ext @_cx>` */
    cx: Emu;
    /** Height in EMU. OpenXML: `<a:ext @_cy>` */
    cy: Emu;
    /** X coordinate in EMU. OpenXML: `<a:off @_x>` */
    x: Emu;
    /** Y coordinate in EMU. OpenXML: `<a:off @_y>` */
    y: Emu;
  };
  /** Rotation in EmuDegrees. OpenXML: `<a:xfrm @_rot>` */
  rotation: EmuDegree;
  /** Outer shadow effect. OpenXML: `<p:spPr><a:effectLst><a:outerShdw>` */
  shadow?: PptxShadow;
  /** Shape type name (e.g. 'rect', 'roundRect', 'ellipse') */
  shapeType?: string;
  /** Table data. OpenXML: `<a:tbl>` inside `<p:graphicFrame>` */
  table?: PptxTable;
  /** Text body, if the element contains text. OpenXML: `<p:txBody>` */
  textBody?: PptxTextBody;
  /** Low-level OpenXML container tag. OpenXML: `<p:sp>`, `<p:pic>`, `<p:graphicFrame>`, `<p:grpSp>`, `<p:cxnSp>` */
  type: 'connector' | 'graphicFrame' | 'group' | 'picture' | 'shape';
  /** 0-based z-index (back-to-front rendering layer order in container) */
  zIndex: number;
}

/** Represents an outer shadow effect. OpenXML: `<a:effectLst><a:outerShdw>` */
export interface PptxShadow {
  /** Alignment. OpenXML: `@_algn` (e.g. 'tl', 'b', 'ctr', 'br') */
  alignment?: string;
  /** Blur radius in EMU. OpenXML: `@_blurRad` */
  blurRadius?: Emu;
  /** Shadow color hex (e.g. '000000'). OpenXML: `<a:srgbClr @_val>` */
  color?: string;
  /** Direction angle in EmuDegrees. OpenXML: `@_dir` */
  direction?: EmuDegree;
  /** Distance in EMU. OpenXML: `@_dist` */
  distance?: Emu;
  /** Opacity decimal (0.0 to 1.0). OpenXML: `<a:alpha @_val>` */
  opacity?: number;
  /** Rotate shadow with shape. OpenXML: `@_rotWithShape` */
  rotateWithShape?: boolean;
}

/** Chart element (embedded via graphicFrame) */
export interface PptxChartElement extends PptxBaseElement {
  chart: PptxChart;
  elementType: 'chart';
  type: 'graphicFrame';
}

/** Connector line element */
export interface PptxConnectorElement extends PptxBaseElement {
  elementType: 'connector';
  type: 'connector';
}

/** Group shape container element */
export interface PptxGroupElement extends PptxBaseElement {
  children: PptxElement[];
  elementType: 'group';
  type: 'group';
}

/** Picture element */
export interface PptxPictureElement extends PptxBaseElement {
  elementType: 'picture';
  picture: PptxPicture;
  type: 'picture';
}

/** AutoShape or Text Box element */
export interface PptxShapeElement extends PptxBaseElement {
  elementType: 'shape';
  /** Whether the shape is explicitly marked as a text box. OpenXML: `<p:cNvSpPr @_txBox="1">` */
  isTextBox?: boolean;
  type: 'shape';
}

/** Table element (embedded via graphicFrame) */
export interface PptxTableElement extends PptxBaseElement {
  elementType: 'table';
  table: PptxTable;
  type: 'graphicFrame';
}

/**
 * Universal Discriminated Union for any visual slide element (Shape, Text Box, Picture, Table, Chart, Group, Connector).
 */
export type PptxElement
  = | PptxChartElement
    | PptxConnectorElement
    | PptxGroupElement
    | PptxPictureElement
    | PptxShapeElement
    | PptxTableElement;

/** Alias for PptxElement */
export type PptxShape = PptxElement;

/** Represents custom XML or package auxiliary parts */
export interface PptxCustomXmlPart {
  /** Optional binary content for embeddings and binary streams */
  binaryData?: Uint8Array;
  /** Target file path in package. OpenXML: `customXml/item1.xml` */
  path: string;
  /** Raw XML content */
  xmlString?: string;
}

/** Document metadata */
export interface PptxMetadata {
  /** Creation date. OpenXML: `<dcterms:created>` */
  created?: Date;
  /** Document creator. OpenXML: `<dc:creator>` */
  creator?: string;
  /** Last modified by. OpenXML: `<cp:lastModifiedBy>` */
  lastModifiedBy?: string;
  /** Last modification date. OpenXML: `<dcterms:modified>` */
  modified?: Date;
  /** Revision number. OpenXML: `<cp:revision>` */
  revision?: number;
  /** Total number of slides */
  slideCount: number;
  /** Slide height in EMU. OpenXML: `<p:sldSz @_cy>` */
  slideHeight: Emu;
  /** Slide width in EMU. OpenXML: `<p:sldSz @_cx>` */
  slideWidth: Emu;
  /** Document title. OpenXML: `<dc:title>` */
  title?: string;
}

/** Represents a single slide */
export interface PptxSlide {
  /** Animations on the slide. OpenXML: `<p:timing>` */
  animations: PptxAnimation[];
  /** Slide background properties. OpenXML: `<p:bg>` */
  background?: PptxBackground;
  /** Visual elements on the slide (alias for shapes) */
  elements: PptxElement[];
  /** Reference to the layout used by this slide. OpenXML: `<p:slide>` rel to `slideLayout*.xml` */
  layoutId?: string;
  /** Notes associated with the slide. OpenXML: `ppt/notesSlides/notesSlide*.xml` */
  notes?: string;
  /** Structured text body for rich speaker notes with formatting. OpenXML: `<p:notes><p:txBody>` */
  notesBody?: PptxTextBody;
  /** Optional raw OpenXML content for exact round-trip fidelity */
  rawXml?: string;
  /** Optional raw relationship XML content (.rels) */
  relsXml?: string;
  /** Shapes on the slide. OpenXML: `<p:spTree>` child nodes */
  shapes: PptxShape[];
  /** Internal slide ID. OpenXML: `rId*` from `presentation.xml.rels` */
  slideId: string;
  /** 1-based slide number */
  slideNumber: number;
  /** Slide transition properties. OpenXML: `<p:transition>` */
  transition?: PptxTransition;
}

/** Represents a slide layout */
export interface PptxSlideLayout {
  /** Visual elements on the layout (alias for shapes) */
  elements: PptxElement[];
  /** Slide layout ID (e.g. 'slideLayout1') */
  id: string;
  /** Reference to the parent slide master ID */
  masterId: string;
  /** Matching name for layout pairing. OpenXML: `<p:sldLayout @_matchingName>` */
  matchingName?: string;
  /** Layout name. OpenXML: `<p:cSld @_name>` or `<p:sldLayout @_matchingName>` */
  name: string;
  /** Whether the layout is preserved when unused. OpenXML: `<p:sldLayout @_preserve>` */
  preserve?: boolean;
  /** Optional raw OpenXML content for exact round-trip fidelity */
  rawXml?: string;
  /** Optional raw relationship XML content (.rels) */
  relsXml?: string;
  /** Background shapes and placeholders */
  shapes: PptxShape[];
  /** Layout type (e.g., 'title', 'custom'). OpenXML: `<p:sldLayout @_type>` */
  type: string;
  /** Whether the layout was custom user-drawn. OpenXML: `<p:sldLayout @_userDrawn>` */
  userDrawn?: boolean;
}

/** Represents a slide master */
export interface PptxSlideMaster {
  /** Visual elements on the master (alias for shapes) */
  elements: PptxElement[];
  /** Slide master ID (e.g. 'slideMaster1') */
  id: string;
  /** Layout IDs associated with this master */
  layoutIds: string[];
  /** Slide master name. OpenXML: `<p:cSld @_name>` or `<p:sldMaster @_name>` */
  name?: string;
  /** Whether the master is preserved when unused. OpenXML: `<p:sldMaster @_preserve>` */
  preserve?: boolean;
  /** Optional raw OpenXML content for exact round-trip fidelity */
  rawXml?: string;
  /** Optional raw relationship XML content (.rels) */
  relsXml?: string;
  /** Background shapes and elements */
  shapes: PptxShape[];
  /** Theme applied to the master */
  theme?: PptxTheme;
}

/** Represents a parsed PPTX document. */
export interface PptxDocument {
  /** Custom XML data parts. OpenXML: `customXml/*` */
  customXml: PptxCustomXmlPart[];
  /** Embedded media assets. OpenXML: `ppt/media/*` */
  media: PptxMediaAsset[];
  /** Document metadata. OpenXML: `docProps/core.xml` & `ppt/presentation.xml` */
  metadata: PptxMetadata;
  /** Slide layouts. OpenXML: `ppt/slideLayouts/slideLayout*.xml` */
  slideLayouts: PptxSlideLayout[];
  /** Slide masters. OpenXML: `ppt/slideMasters/slideMaster*.xml` */
  slideMasters: PptxSlideMaster[];
  /** Slides in presentation order. OpenXML: `ppt/slides/slide*.xml` */
  slides: PptxSlide[];
  /** Themes. OpenXML: `ppt/theme/theme*.xml` */
  themes: PptxTheme[];
}
