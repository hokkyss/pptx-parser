import { createFileRoute, redirect } from '@tanstack/react-router';

export const Route = createFileRoute('/api/')({
  beforeLoad: () => {
    throw redirect({
      params: { _splat: 'pptx/presentation' },
      to: '/api/$',
    });
  },
});
