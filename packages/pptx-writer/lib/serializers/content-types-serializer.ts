import { el, serializeXml, type XmlElement } from '../xml/xml-element';

export interface ContentTypesOptions {
  customPartOverrides?: Array<{ contentType: string; partName: string }>;
  hasCharts?: boolean;
  layoutCount?: number;
  layoutNames?: string[];
  masterCount?: number;
  masterNames?: string[];
  mediaExtensions?: string[];
  slideCount?: number;
  themeCount?: number;
  themeNames?: string[];
}

const MIME_EXTENSION_MAP: Record<string, string> = {
  avi: 'video/x-msvideo',
  bmp: 'image/bmp',
  gif: 'image/gif',
  jpeg: 'image/jpeg',
  jpg: 'image/jpeg',
  m4a: 'audio/mp4',
  mov: 'video/quicktime',
  mp3: 'audio/mpeg',
  mp4: 'video/mp4',
  png: 'image/png',
  rels: 'application/vnd.openxmlformats-package.relationships+xml',
  svg: 'image/svg+xml',
  tif: 'image/tiff',
  tiff: 'image/tiff',
  wav: 'audio/x-wav',
  wma: 'audio/x-ms-wma',
  wmf: 'image/x-wmf',
  wmv: 'video/x-ms-wmv',
  xml: 'application/xml',
};

/**
 * Serializes `[Content_Types].xml` with defaults and part overrides.
 * @param options
 */
export function serializeContentTypes(options: ContentTypesOptions = {}): string {
  const defaultList: Array<{ ContentType: string; Extension: string }> = [
    { ContentType: MIME_EXTENSION_MAP.rels, Extension: 'rels' },
    { ContentType: MIME_EXTENSION_MAP.xml, Extension: 'xml' },
  ];

  // Add media extensions to defaults if present
  const extensions = new Set(options.mediaExtensions ?? ['png', 'jpeg', 'jpg']);
  for (const ext of extensions) {
    const cleanExt = ext.toLowerCase().replace(/^\./, '');
    const mime = MIME_EXTENSION_MAP[cleanExt] ?? `image/${cleanExt}`;
    if (!defaultList.some((d) => d.Extension === cleanExt)) {
      defaultList.push({
        ContentType: mime,
        Extension: cleanExt,
      });
    }
  }

  const overridesList: Array<{ ContentType: string; PartName: string }> = [
    {
      ContentType: 'application/vnd.openxmlformats-package.core-properties+xml',
      PartName: '/docProps/core.xml',
    },
    {
      ContentType: 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
      PartName: '/docProps/app.xml',
    },
    {
      ContentType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml',
      PartName: '/ppt/presentation.xml',
    },
    {
      ContentType: 'application/vnd.openxmlformats-officedocument.presentationml.presProps+xml',
      PartName: '/ppt/presProps.xml',
    },
    {
      ContentType: 'application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml',
      PartName: '/ppt/tableStyles.xml',
    },
    {
      ContentType: 'application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml',
      PartName: '/ppt/viewProps.xml',
    },
  ];

  // Slide overrides
  const slideCount = options.slideCount ?? 1;
  for (let i = 1; i <= slideCount; i++) {
    overridesList.push({
      ContentType: 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml',
      PartName: `/ppt/slides/slide${i}.xml`,
    });
  }

  // Layout overrides
  if (options.layoutNames && options.layoutNames.length > 0) {
    for (const name of options.layoutNames) {
      const fileName = name.endsWith('.xml') ? name : `${name}.xml`;
      overridesList.push({
        ContentType: 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml',
        PartName: `/ppt/slideLayouts/${fileName}`,
      });
    }
  } else {
    const layoutCount = options.layoutCount ?? 1;
    for (let i = 1; i <= layoutCount; i++) {
      overridesList.push({
        ContentType: 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml',
        PartName: `/ppt/slideLayouts/slideLayout${i}.xml`,
      });
    }
  }

  // Master overrides
  if (options.masterNames && options.masterNames.length > 0) {
    for (const name of options.masterNames) {
      const fileName = name.endsWith('.xml') ? name : `${name}.xml`;
      overridesList.push({
        ContentType: 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml',
        PartName: `/ppt/slideMasters/${fileName}`,
      });
    }
  } else {
    const masterCount = options.masterCount ?? 1;
    for (let i = 1; i <= masterCount; i++) {
      overridesList.push({
        ContentType: 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml',
        PartName: `/ppt/slideMasters/slideMaster${i}.xml`,
      });
    }
  }

  // Theme overrides
  if (options.themeNames && options.themeNames.length > 0) {
    for (const name of options.themeNames) {
      const fileName = name.endsWith('.xml') ? name : `${name}.xml`;
      overridesList.push({
        ContentType: 'application/vnd.openxmlformats-officedocument.theme+xml',
        PartName: `/ppt/theme/${fileName}`,
      });
    }
  } else {
    const themeCount = options.themeCount ?? 1;
    for (let i = 1; i <= themeCount; i++) {
      overridesList.push({
        ContentType: 'application/vnd.openxmlformats-officedocument.theme+xml',
        PartName: `/ppt/theme/theme${i}.xml`,
      });
    }
  }

  // Custom part overrides
  if (options.customPartOverrides) {
    for (const part of options.customPartOverrides) {
      overridesList.push({
        ContentType: part.contentType,
        PartName: part.partName.startsWith('/') ? part.partName : `/${part.partName}`,
      });
    }
  }

  const defaultNodes: XmlElement[] = defaultList.map((d) =>
    el('Default', {
      ContentType: d.ContentType,
      Extension: d.Extension,
    }),
  );

  const overrideNodes: XmlElement[] = overridesList.map((o) =>
    el('Override', {
      ContentType: o.ContentType,
      PartName: o.PartName,
    }),
  );

  const root = el('Types', {
    xmlns: 'http://schemas.openxmlformats.org/package/2006/content-types',
  }, [...defaultNodes, ...overrideNodes]);

  return serializeXml(root);
}
