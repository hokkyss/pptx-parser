import { serializeXml } from '../xml/xml-builder';

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
  bmp: 'image/bmp',
  gif: 'image/gif',
  jpeg: 'image/jpeg',
  jpg: 'image/jpeg',
  png: 'image/png',
  rels: 'application/vnd.openxmlformats-package.relationships+xml',
  svg: 'image/svg+xml',
  tif: 'image/tiff',
  tiff: 'image/tiff',
  wmf: 'image/x-wmf',
  xml: 'application/xml',
};

/**
 * Serializes `[Content_Types].xml` with defaults and part overrides.
 * @param options
 */
export function serializeContentTypes(options: ContentTypesOptions = {}): string {
  const defaults: Array<{ '@_ContentType': string; '@_Extension': string }> = [
    { '@_ContentType': MIME_EXTENSION_MAP.rels, '@_Extension': 'rels' },
    { '@_ContentType': MIME_EXTENSION_MAP.xml, '@_Extension': 'xml' },
  ];

  // Add media extensions to defaults if present
  const extensions = new Set(options.mediaExtensions ?? ['png', 'jpeg', 'jpg']);
  for (const ext of extensions) {
    const cleanExt = ext.toLowerCase().replace(/^\./, '');
    const mime = MIME_EXTENSION_MAP[cleanExt] ?? `image/${cleanExt}`;
    if (!defaults.some((d) => d['@_Extension'] === cleanExt)) {
      defaults.push({
        '@_ContentType': mime,
        '@_Extension': cleanExt,
      });
    }
  }

  const overrides: Array<{ '@_ContentType': string; '@_PartName': string }> = [
    {
      '@_ContentType': 'application/vnd.openxmlformats-package.core-properties+xml',
      '@_PartName': '/docProps/core.xml',
    },
    {
      '@_ContentType': 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
      '@_PartName': '/docProps/app.xml',
    },
    {
      '@_ContentType': 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml',
      '@_PartName': '/ppt/presentation.xml',
    },
    {
      '@_ContentType': 'application/vnd.openxmlformats-officedocument.presentationml.presProps+xml',
      '@_PartName': '/ppt/presProps.xml',
    },
    {
      '@_ContentType': 'application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml',
      '@_PartName': '/ppt/tableStyles.xml',
    },
    {
      '@_ContentType': 'application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml',
      '@_PartName': '/ppt/viewProps.xml',
    },
  ];

  // Slide overrides
  const slideCount = options.slideCount ?? 1;
  for (let i = 1; i <= slideCount; i++) {
    overrides.push({
      '@_ContentType': 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml',
      '@_PartName': `/ppt/slides/slide${i}.xml`,
    });
  }

  // Layout overrides
  if (options.layoutNames && options.layoutNames.length > 0) {
    for (const name of options.layoutNames) {
      const fileName = name.endsWith('.xml') ? name : `${name}.xml`;
      overrides.push({
        '@_ContentType': 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml',
        '@_PartName': `/ppt/slideLayouts/${fileName}`,
      });
    }
  } else {
    const layoutCount = options.layoutCount ?? 1;
    for (let i = 1; i <= layoutCount; i++) {
      overrides.push({
        '@_ContentType': 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml',
        '@_PartName': `/ppt/slideLayouts/slideLayout${i}.xml`,
      });
    }
  }

  // Master overrides
  if (options.masterNames && options.masterNames.length > 0) {
    for (const name of options.masterNames) {
      const fileName = name.endsWith('.xml') ? name : `${name}.xml`;
      overrides.push({
        '@_ContentType': 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml',
        '@_PartName': `/ppt/slideMasters/${fileName}`,
      });
    }
  } else {
    const masterCount = options.masterCount ?? 1;
    for (let i = 1; i <= masterCount; i++) {
      overrides.push({
        '@_ContentType': 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml',
        '@_PartName': `/ppt/slideMasters/slideMaster${i}.xml`,
      });
    }
  }

  // Theme overrides
  if (options.themeNames && options.themeNames.length > 0) {
    for (const name of options.themeNames) {
      const fileName = name.endsWith('.xml') ? name : `${name}.xml`;
      overrides.push({
        '@_ContentType': 'application/vnd.openxmlformats-officedocument.theme+xml',
        '@_PartName': `/ppt/theme/${fileName}`,
      });
    }
  } else {
    const themeCount = options.themeCount ?? 1;
    for (let i = 1; i <= themeCount; i++) {
      overrides.push({
        '@_ContentType': 'application/vnd.openxmlformats-officedocument.theme+xml',
        '@_PartName': `/ppt/theme/theme${i}.xml`,
      });
    }
  }

  // Custom part overrides
  if (options.customPartOverrides) {
    for (const part of options.customPartOverrides) {
      overrides.push({
        '@_ContentType': part.contentType,
        '@_PartName': part.partName.startsWith('/') ? part.partName : `/${part.partName}`,
      });
    }
  }

  const root = {
    Types: {
      '@_xmlns': 'http://schemas.openxmlformats.org/package/2006/content-types',
      Default: defaults,
      Override: overrides,
    },
  };

  return serializeXml(root);
}
