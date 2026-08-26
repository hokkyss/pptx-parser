import { createFileRoute, redirect } from '@tanstack/react-router';

export const Route = createFileRoute('/api-reference/')({
  beforeLoad: () => {
    throw redirect({
      params: { _splat: 'pptx/presentation' },
      to: '/api-reference/$',
    });
  },
});
