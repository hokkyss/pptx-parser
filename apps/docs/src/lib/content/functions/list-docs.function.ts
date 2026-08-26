import { createServerFn } from '@tanstack/react-start';
import { setResponseHeader } from '@tanstack/react-start/server';
import { getAllDocs } from '../content-manifest';
import { listDocsRequestDto, listDocsResponseDto } from '../dto/list-docs.dto';

const sectionLabels = {
  authoring: 'Authoring Guide',
  'core-concepts': 'Core Concepts',
  'getting-started': 'Getting Started',
  'parsing-and-mutation': 'Parsing & Mutation',
  performance: 'Performance & Benchmarks',
  pptx: '@hokkyss/pptx',
  'pptx-core': '@hokkyss/pptx-core',
  'pptx-reader': '@hokkyss/pptx-reader',
  'pptx-writer': '@hokkyss/pptx-writer',
  'runtimes-and-deploy': 'Runtimes & Deployment',
} as const satisfies Record<string, string>;

const sectionOrder: (keyof typeof sectionLabels)[] = [
  'getting-started',
  'core-concepts',
  'parsing-and-mutation',
  'authoring',
  'performance',
  'runtimes-and-deploy',
  'pptx',
  'pptx-core',
  'pptx-reader',
  'pptx-writer',
];

const getDocSection = (pathStr: string) => {
  const parts = pathStr.split('/');
  if ((parts[0] === 'docs' || parts[0] === 'api-reference') && parts[1]) {
    return parts[1];
  }
  return parts[0];
};

const listDocsFunction = createServerFn({
  method: 'GET',
})
  .validator(listDocsRequestDto)
  .handler(({ data }) => {
    setResponseHeader('X-Cache-Maxage', '2592000');
    setResponseHeader('X-Stale-After', '604800');

    const allDocs = getAllDocs();
    const filteredDocs = data.section
      ? allDocs.filter((d) => d.path.startsWith(data.section!))
      : allDocs;

    const grouped: Record<string, { description?: string; order?: number; path: string; title: string }[]> = {};

    for (const doc of filteredDocs) {
      const sec = getDocSection(doc.path);
      if (!grouped[sec]) {
        grouped[sec] = [];
      }
      grouped[sec].push({
        description: doc.description,
        order: doc.order ?? 99,
        path: doc.path,
        title: doc.title,
      });
    }

    const sections = Object.entries(grouped)
      .sort(([aKey], [bKey]) => {
        const aIdx = sectionOrder.indexOf(aKey as keyof typeof sectionLabels);
        const bIdx = sectionOrder.indexOf(bKey as keyof typeof sectionLabels);
        const resolvedA = aIdx === -1 ? 999 : aIdx;
        const resolvedB = bIdx === -1 ? 999 : bIdx;
        return resolvedA - resolvedB;
      })
      .map(([secKey, items]) => ({
        items: items.sort((a, b) => (a.order ?? 99) - (b.order ?? 99)),
        title: sectionLabels[secKey as keyof typeof sectionLabels] || secKey.toUpperCase(),
      }));

    return listDocsResponseDto.parse({ sections });
  });

export default listDocsFunction;
