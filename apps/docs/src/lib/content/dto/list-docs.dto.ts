import { z } from 'zod/v4';

export const listDocsRequestDto = z.object({
  section: z.string().optional(),
});

export type ListDocsRequestDto = z.infer<typeof listDocsRequestDto>;

export const navItemSchema = z.object({
  description: z.string().optional(),
  order: z.number().optional(),
  path: z.string(),
  title: z.string(),
});

export const navSectionSchema = z.object({
  items: z.array(navItemSchema),
  title: z.string(),
});

export const listDocsResponseDto = z.object({
  sections: z.array(navSectionSchema),
});

export type ListDocsResponseDto = z.infer<typeof listDocsResponseDto>;
