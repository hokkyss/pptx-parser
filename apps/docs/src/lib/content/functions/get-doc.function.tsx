import { createServerFn } from '@tanstack/react-start';
import { renderServerComponent } from '@tanstack/react-start/rsc';
import { setResponseHeader } from '@tanstack/react-start/server';
import MarkdownRenderer from '../../../components/markdown-renderer.component';
import { getDocByPath } from '../content-manifest';
import { getDocRequestDto, getDocResponseDto } from '../dto/get-doc.dto';

const getDocFunction = createServerFn({
  method: 'GET',
})
  .validator(getDocRequestDto)
  .handler(async ({ data }) => {
    // 30 days edge cache, 7 days client cache
    setResponseHeader('X-Cache-Maxage', '2592000');
    setResponseHeader('X-Stale-After', '604800');

    const doc = getDocByPath(data.path);

    if (!doc) {
      throw new Error(`Documentation page not found: ${data.path}`);
    }

    const Renderable = await renderServerComponent(
      <MarkdownRenderer content={doc.content} />,
    );

    return getDocResponseDto.parse({
      content: doc.content,
      description: doc.description,
      frontmatter: doc.frontmatter,
      order: doc.order,
      package: doc.package,
      path: doc.path,
      Renderable,
      slug: doc.slug,
      title: doc.title,
      toc: doc.toc,
    });
  });

export default getDocFunction;
