import { createFileRoute, redirect } from '@tanstack/react-router';

export const Route = createFileRoute('/docs/')({
  beforeLoad: () => {
    throw redirect({
      params: { _splat: 'getting-started/overview' },
      to: '/docs/$',
    });
  },
});
