import { createServerFn } from '@tanstack/react-start';
import { setResponseHeader } from '@tanstack/react-start/server';
import { getDocByPath } from '../content-manifest';
import { getDocRequestDto, getDocResponseDto } from '../dto/get-doc.dto';

const getDocFunction = createServerFn({
  method: 'GET',
})
  .validator(getDocRequestDto)
  .handler(({ data }) => {
    // 30 days edge cache, 7 days client cache
    setResponseHeader('X-Cache-Maxage', '2592000');
    setResponseHeader('X-Stale-After', '604800');

    const doc = getDocByPath(data.path);

    if (!doc) {
      throw new Error(`Documentation page not found: ${data.path}`);
    }

    return getDocResponseDto.parse(doc);
  });

export default getDocFunction;
