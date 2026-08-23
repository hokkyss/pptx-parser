import type { RenderableServerComponent } from '@tanstack/react-start/rsc';
import { ReactNode } from 'react';
import { z } from 'zod/v4';

export const renderTeaserRequestDto = z.object({
});

export type RenderTeaserRequestDto = z.input<typeof renderTeaserRequestDto>;

export const renderTeaserResponseDto = z.custom<RenderableServerComponent<ReactNode>>();

export type RenderTeaserResponseDto = z.infer<typeof renderTeaserResponseDto>;
