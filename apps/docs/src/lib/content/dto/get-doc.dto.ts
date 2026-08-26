import type { RenderableServerComponent } from '@tanstack/react-start/rsc';
import type { ReactNode } from 'react';
import { z } from 'zod/v4';

export const getDocRequestDto = z.object({
  path: z.string(),
});

export type GetDocRequestDto = z.infer<typeof getDocRequestDto>;

export const tocItemSchema = z.object({
  id: z.string(),
  level: z.number(),
  text: z.string(),
});

export const getDocResponseDto = z.object({
  content: z.string(),
  description: z.string().optional(),
  frontmatter: z.object({
    description: z.string().optional(),
    order: z.number().optional(),
    pkg: z.string().optional(),
    title: z.string().optional(),
  }),
  order: z.number().optional(),
  package: z.string().optional(),
  path: z.string(),
  Renderable: z.custom<RenderableServerComponent<ReactNode>>(),
  slug: z.string(),
  title: z.string(),
  toc: z.array(tocItemSchema),
});

export type GetDocResponseDto = z.infer<typeof getDocResponseDto>;
